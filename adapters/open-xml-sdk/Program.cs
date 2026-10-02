// PPTBench container helper for the Open XML SDK (NuGet DocumentFormat.OpenXml 3.5.1).
//
// Usage: PptBenchOpenXml <feature-matrix|template-mutation|chart-data> < input.pptx > output.pptx
// Reads the whole input deck from stdin, edits it through the SDK's typed DOM, and writes the
// complete output package to stdout. Diagnostics go to stderr only. Exit codes: 0 success,
// 1 edit failure, 2 usage error / unknown lane.
//
// Lanes mirror src/pptbench/worker.py:
//   feature-matrix     open the deck and save it unchanged.
//   template-mutation  slide 5 table "PPTBenchTable" cells (1,1)/(2,2), slide 7 shape
//                      "PPTBenchBullets" paragraph 1; the first existing run keeps its
//                      properties, extra runs/breaks/fields in the paragraph are dropped.
//   chart-data         slide 1 bar, slide 2 scatter, slide 3 bubble: rewrite the c:strCache /
//                      c:numCache points and the embedded workbook cells the series formulas
//                      reference (opened as a SpreadsheetDocument and fed back to the part).

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace PptBenchOpenXml;

internal static class Program
{
    private static readonly string[] Lanes = { "feature-matrix", "template-mutation", "chart-data" };

    private static int Main(string[] args)
    {
        if (args.Length != 1 || !Lanes.Contains(args[0]))
        {
            Console.Error.WriteLine($"usage: PptBenchOpenXml <{string.Join("|", Lanes)}> < input.pptx > output.pptx");
            return 2;
        }

        try
        {
            using var buffer = new MemoryStream();
            using (var stdin = Console.OpenStandardInput())
            {
                stdin.CopyTo(buffer);
            }
            if (buffer.Length == 0)
            {
                throw new InvalidDataException("stdin contained no presentation bytes");
            }
            buffer.Position = 0;

            using (var document = PresentationDocument.Open(buffer, true))
            {
                switch (args[0])
                {
                    case "template-mutation":
                        MutateTemplate(document);
                        break;
                    case "chart-data":
                        ReplaceChartData(document);
                        break;
                }
                document.Save();
            }

            var output = buffer.ToArray();
            using (var stdout = Console.OpenStandardOutput())
            {
                stdout.Write(output, 0, output.Length);
                stdout.Flush();
            }
            return 0;
        }
        catch (Exception exc)
        {
            Console.Error.WriteLine($"open-xml-sdk {args[0]} failed: {exc}");
            return 1;
        }
    }

    // ---- template-mutation -------------------------------------------------------------

    private static void MutateTemplate(PresentationDocument document)
    {
        var slides = SlidesInOrder(document);

        var tableFrame = Single(
            Root(SlideAt(slides, 4).Slide, "slide 5").Descendants<P.GraphicFrame>()
                .Where(frame => frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value == "PPTBenchTable"),
            "graphic frame PPTBenchTable on slide 5");
        var table = tableFrame.Graphic?.GraphicData?.GetFirstChild<A.Table>()
            ?? throw new InvalidDataException("PPTBenchTable does not contain an a:tbl");
        SetCellText(TableCell(table, 1, 1), "UPDATED-TABLE-A");
        SetCellText(TableCell(table, 2, 2), "UPDATED-TABLE-B");

        var bullets = Single(
            Root(SlideAt(slides, 6).Slide, "slide 7").Descendants<P.Shape>()
                .Where(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "PPTBenchBullets"),
            "shape PPTBenchBullets on slide 7");
        var paragraphs = bullets.TextBody?.Elements<A.Paragraph>().ToList()
            ?? throw new InvalidDataException("PPTBenchBullets has no text body");
        if (paragraphs.Count < 2)
        {
            throw new InvalidDataException("PPTBenchBullets has no paragraph index 1");
        }
        SetParagraphText(paragraphs[1], "UPDATED-BULLET");
    }

    private static A.TableCell TableCell(A.Table table, int row, int column)
    {
        var rows = table.Elements<A.TableRow>().ToList();
        if (row >= rows.Count)
        {
            throw new InvalidDataException($"table has no row {row}");
        }
        var cells = rows[row].Elements<A.TableCell>().ToList();
        if (column >= cells.Count)
        {
            throw new InvalidDataException($"table row {row} has no column {column}");
        }
        return cells[column];
    }

    private static void SetCellText(A.TableCell cell, string value)
    {
        var body = cell.TextBody ?? cell.AppendChild(new A.TextBody(new A.BodyProperties(), new A.ListStyle()));
        var paragraphs = body.Elements<A.Paragraph>().ToList();
        if (paragraphs.Count == 0)
        {
            paragraphs.Add(body.AppendChild(new A.Paragraph()));
        }
        foreach (var extra in paragraphs.Skip(1))
        {
            extra.Remove();
        }
        SetParagraphText(paragraphs[0], value);
    }

    private static void SetParagraphText(A.Paragraph paragraph, string value)
    {
        foreach (var inline in paragraph.ChildElements.Where(child => child is A.Break || child is A.Field).ToList())
        {
            inline.Remove();
        }
        var runs = paragraph.Elements<A.Run>().ToList();
        if (runs.Count == 0)
        {
            var run = new A.Run(new A.Text(value));
            var end = paragraph.GetFirstChild<A.EndParagraphRunProperties>();
            if (end is null)
            {
                paragraph.AppendChild(run);
            }
            else
            {
                paragraph.InsertBefore(run, end);
            }
            return;
        }
        foreach (var extra in runs.Skip(1))
        {
            extra.Remove();
        }
        var text = runs[0].Text ?? runs[0].AppendChild(new A.Text());
        text.Text = value;
    }

    // ---- chart-data --------------------------------------------------------------------

    private sealed record CellEdit(string Formula, string[] Values, bool IsText);

    private static void ReplaceChartData(PresentationDocument document)
    {
        var slides = SlidesInOrder(document);

        {
            var chartPart = ChartOnSlide(SlideAt(slides, 0));
            var bar = Single(PlotArea(chartPart).Elements<C.BarChart>(), "c:barChart on slide 1");
            var series = Single(bar.Elements<C.BarChartSeries>(), "bar series");
            RequireSeriesName(series, "Revenue");
            var categories = series.GetFirstChild<C.CategoryAxisData>()?.GetFirstChild<C.StringReference>()
                ?? throw new InvalidDataException("bar series categories are not a c:strRef");
            var values = series.GetFirstChild<C.Values>()?.GetFirstChild<C.NumberReference>()
                ?? throw new InvalidDataException("bar series values are not a c:numRef");
            var edits = new List<CellEdit>
            {
                SetStrings(categories, new[] { "East", "42", "84" }),
                SetNumbers(values, new double[] { 42, 84, 126 }),
            };
            UpdateWorkbook(chartPart, edits);
        }

        {
            var chartPart = ChartOnSlide(SlideAt(slides, 1));
            var scatter = Single(PlotArea(chartPart).Elements<C.ScatterChart>(), "c:scatterChart on slide 2");
            var series = Single(scatter.Elements<C.ScatterChartSeries>(), "scatter series");
            RequireSeriesName(series, "Trend");
            var edits = new List<CellEdit>
            {
                SetNumbers(NumberReference(series.GetFirstChild<C.XValues>(), "scatter x"), new double[] { 5, 21, 55 }),
                SetNumbers(NumberReference(series.GetFirstChild<C.YValues>(), "scatter y"), new double[] { 13, 34, 89 }),
            };
            UpdateWorkbook(chartPart, edits);
        }

        {
            var chartPart = ChartOnSlide(SlideAt(slides, 2));
            var bubble = Single(PlotArea(chartPart).Elements<C.BubbleChart>(), "c:bubbleChart on slide 3");
            var series = Single(bubble.Elements<C.BubbleChartSeries>(), "bubble series");
            RequireSeriesName(series, "Pipeline");
            var edits = new List<CellEdit>
            {
                SetNumbers(NumberReference(series.GetFirstChild<C.XValues>(), "bubble x"), new double[] { 5, 34, 8 }),
                SetNumbers(NumberReference(series.GetFirstChild<C.YValues>(), "bubble y"), new double[] { 13, 55, 13 }),
                SetNumbers(NumberReference(series.GetFirstChild<C.BubbleSize>(), "bubble size"), new double[] { 21, 89, 21 }),
            };
            UpdateWorkbook(chartPart, edits);
        }
    }

    private static ChartPart ChartOnSlide(SlidePart slidePart)
    {
        var reference = Single(Root(slidePart.Slide, "chart slide").Descendants<C.ChartReference>(), "chart graphic frame");
        var id = reference.Id?.Value ?? throw new InvalidDataException("chart reference has no r:id");
        return slidePart.GetPartById(id) as ChartPart
            ?? throw new InvalidDataException($"relationship {id} is not a chart part");
    }

    private static C.PlotArea PlotArea(ChartPart chartPart) =>
        chartPart.ChartSpace?.GetFirstChild<C.Chart>()?.PlotArea
        ?? throw new InvalidDataException("chart has no c:plotArea");

    private static void RequireSeriesName(OpenXmlCompositeElement series, string expected)
    {
        var name = series.GetFirstChild<C.SeriesText>()?.Descendants<C.StringPoint>().FirstOrDefault()?.NumericValue?.Text;
        if (name != expected)
        {
            throw new InvalidDataException($"expected series '{expected}', found '{name}'");
        }
    }

    private static C.NumberReference NumberReference(OpenXmlCompositeElement? source, string label) =>
        source?.GetFirstChild<C.NumberReference>()
        ?? throw new InvalidDataException($"{label} data is not a c:numRef");

    private static CellEdit SetStrings(C.StringReference reference, string[] values)
    {
        var formula = reference.Formula?.Text ?? throw new InvalidDataException("c:strRef has no formula");
        var cache = reference.StringCache ?? throw new InvalidDataException("c:strRef has no c:strCache");
        foreach (var point in cache.Elements<C.StringPoint>().ToList())
        {
            point.Remove();
        }
        var count = cache.PointCount ?? cache.PrependChild(new C.PointCount());
        count.Val = (uint)values.Length;
        OpenXmlElement anchor = count;
        for (var index = 0; index < values.Length; index++)
        {
            anchor = cache.InsertAfter(new C.StringPoint(new C.NumericValue(values[index])) { Index = (uint)index }, anchor);
        }
        return new CellEdit(formula, values, IsText: true);
    }

    private static CellEdit SetNumbers(C.NumberReference reference, double[] numbers)
    {
        var formula = reference.Formula?.Text ?? throw new InvalidDataException("c:numRef has no formula");
        var cache = reference.NumberingCache ?? throw new InvalidDataException("c:numRef has no c:numCache");
        var values = numbers.Select(number => number.ToString(CultureInfo.InvariantCulture)).ToArray();
        foreach (var point in cache.Elements<C.NumericPoint>().ToList())
        {
            point.Remove();
        }
        var count = cache.PointCount;
        if (count is null)
        {
            count = new C.PointCount();
            var format = cache.FormatCode;
            if (format is null)
            {
                cache.PrependChild(count);
            }
            else
            {
                cache.InsertAfter(count, format);
            }
        }
        count.Val = (uint)values.Length;
        OpenXmlElement anchor = count;
        for (var index = 0; index < values.Length; index++)
        {
            anchor = cache.InsertAfter(new C.NumericPoint(new C.NumericValue(values[index])) { Index = (uint)index }, anchor);
        }
        return new CellEdit(formula, values, IsText: false);
    }

    // ---- embedded workbook -------------------------------------------------------------

    private static void UpdateWorkbook(ChartPart chartPart, IReadOnlyList<CellEdit> edits)
    {
        var embedded = chartPart.EmbeddedPackagePart
            ?? throw new InvalidDataException($"chart {chartPart.Uri} has no embedded workbook package");

        using var workbookBytes = new MemoryStream();
        using (var source = embedded.GetStream(FileMode.Open, FileAccess.Read))
        {
            source.CopyTo(workbookBytes);
        }
        workbookBytes.Position = 0;

        using (var workbook = SpreadsheetDocument.Open(workbookBytes, true))
        {
            var workbookPart = workbook.WorkbookPart ?? throw new InvalidDataException("embedded workbook has no workbook part");
            foreach (var edit in edits)
            {
                var (sheetName, coordinates) = ParseFormula(edit.Formula);
                if (coordinates.Count != edit.Values.Length)
                {
                    throw new InvalidDataException($"formula {edit.Formula} covers {coordinates.Count} cells, expected {edit.Values.Length}");
                }
                var sheet = Single(
                    Root(workbookPart.Workbook, "workbook").Descendants<S.Sheet>().Where(candidate => candidate.Name?.Value == sheetName),
                    $"worksheet '{sheetName}'");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(
                    sheet.Id?.Value ?? throw new InvalidDataException($"sheet '{sheetName}' has no r:id"));
                for (var index = 0; index < coordinates.Count; index++)
                {
                    var cell = Single(
                        Root(worksheetPart.Worksheet, $"worksheet '{sheetName}'").Descendants<S.Cell>().Where(candidate => candidate.CellReference?.Value == coordinates[index]),
                        $"cell {sheetName}!{coordinates[index]}");
                    if (edit.IsText)
                    {
                        SetTextCell(workbookPart, cell, edit.Values[index]);
                    }
                    else
                    {
                        SetNumberCell(cell, edit.Values[index]);
                    }
                }
            }
            workbook.Save();
        }

        workbookBytes.Position = 0;
        embedded.FeedData(workbookBytes);
    }

    private static void SetNumberCell(S.Cell cell, string value)
    {
        cell.InlineString?.Remove();
        cell.CellFormula?.Remove();
        cell.DataType = null;
        cell.CellValue = new S.CellValue(value);
    }

    private static void SetTextCell(WorkbookPart workbookPart, S.Cell cell, string value)
    {
        cell.CellFormula?.Remove();
        var table = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (table is null)
        {
            cell.CellValue?.Remove();
            cell.DataType = S.CellValues.InlineString;
            cell.InlineString = new S.InlineString(new S.Text(value));
            return;
        }

        var wasShared = cell.DataType?.Value == S.CellValues.SharedString;
        var items = table.Elements<S.SharedStringItem>().ToList();
        var index = items.FindIndex(item => !item.Elements<S.Run>().Any() && item.Text?.Text == value);
        if (index < 0)
        {
            table.AppendChild(new S.SharedStringItem(new S.Text(value)));
            index = items.Count;
        }
        table.UniqueCount = (uint)table.Elements<S.SharedStringItem>().Count();
        if (!wasShared && table.Count is not null)
        {
            table.Count = table.Count.Value + 1;
        }
        cell.InlineString?.Remove();
        cell.DataType = S.CellValues.SharedString;
        cell.CellValue = new S.CellValue(index.ToString(CultureInfo.InvariantCulture));
    }

    private static readonly Regex RangePattern = new(@"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", RegexOptions.CultureInvariant);

    private static (string Sheet, List<string> Coordinates) ParseFormula(string formula)
    {
        var bang = formula.LastIndexOf('!');
        if (bang <= 0)
        {
            throw new InvalidDataException($"unsupported series formula: {formula}");
        }
        var sheet = formula[..bang];
        if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[^1] == '\'')
        {
            sheet = sheet[1..^1].Replace("''", "'");
        }
        var match = RangePattern.Match(formula[(bang + 1)..]);
        if (!match.Success)
        {
            throw new InvalidDataException($"unsupported series formula: {formula}");
        }
        var startColumn = ColumnNumber(match.Groups[1].Value);
        var startRow = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
        var endColumn = match.Groups[3].Success ? ColumnNumber(match.Groups[3].Value) : startColumn;
        var endRow = match.Groups[4].Success ? int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture) : startRow;
        var coordinates = new List<string>();
        for (var row = startRow; row <= endRow; row++)
        {
            for (var column = startColumn; column <= endColumn; column++)
            {
                coordinates.Add(ColumnName(column) + row.ToString(CultureInfo.InvariantCulture));
            }
        }
        return (sheet, coordinates);
    }

    private static int ColumnNumber(string column) => column.Aggregate(0, (total, letter) => total * 26 + letter - 'A' + 1);

    private static string ColumnName(int number)
    {
        var name = "";
        while (number > 0)
        {
            var remainder = (number - 1) % 26;
            name = (char)('A' + remainder) + name;
            number = (number - 1) / 26;
        }
        return name;
    }

    // ---- shared ------------------------------------------------------------------------

    private static List<SlidePart> SlidesInOrder(PresentationDocument document)
    {
        var presentationPart = document.PresentationPart ?? throw new InvalidDataException("package has no presentation part");
        var slideIds = Root(presentationPart.Presentation, "presentation").SlideIdList?.Elements<P.SlideId>()
            ?? throw new InvalidDataException("presentation has no p:sldIdLst");
        return slideIds
            .Select(slideId => (SlidePart)presentationPart.GetPartById(
                slideId.RelationshipId?.Value ?? throw new InvalidDataException("p:sldId has no r:id")))
            .ToList();
    }

    private static SlidePart SlideAt(List<SlidePart> slides, int index) =>
        index < slides.Count ? slides[index] : throw new InvalidDataException($"deck has no slide index {index}");

    private static T Single<T>(IEnumerable<T> candidates, string label)
    {
        var found = candidates.Take(2).ToList();
        return found.Count == 1
            ? found[0]
            : throw new InvalidDataException($"expected exactly one {label}, found {(found.Count == 0 ? "none" : "several")}");
    }

    private static T Root<T>(T? root, string label) where T : class =>
        root ?? throw new InvalidDataException($"{label} part has no root element");
}
