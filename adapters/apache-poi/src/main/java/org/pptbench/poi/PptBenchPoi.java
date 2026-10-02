/*
 * PPTBench container helper for Apache POI XSLF (org.apache.poi:poi-ooxml).
 *
 * Usage: PptBenchPoi <feature-matrix|template-mutation|chart-data> < input.pptx > output.pptx
 *
 * Reads the whole input deck from stdin with XMLSlideShow, applies the lane edit through
 * POI's public XSLF/XDDF API, and writes the complete output deck to stdout. Diagnostics go
 * to stderr only; System.out is redirected to stderr before POI is touched so no library
 * output can corrupt the package stream. Exit codes: 0 success, 1 lane failure, 2 usage.
 *
 * Bubble charts: POI 5.5.1 ships XDDFBubbleChartData, but XDDFChart.getChartSeries() does not
 * enumerate bubble charts, so the helper wraps the plot area's CTBubbleChart (typed XMLBeans
 * model) in XDDFBubbleChartData. XDDFChart.plot() fills only category/value columns of the
 * embedded workbook, so bubble-size cells are written through XDDFChart.getWorkbook(); the
 * workbook is persisted by XDDFChart.commit() -> saveWorkbook() when the deck is written.
 */
package org.pptbench.poi;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileDescriptor;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBubbleChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBubbleChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;

public final class PptBenchPoi {
    private PptBenchPoi() {
    }

    public static void main(String[] args) {
        PrintStream stdout = new PrintStream(new FileOutputStream(FileDescriptor.out), false);
        System.setOut(System.err);
        if (args.length != 1) {
            System.err.println("usage: PptBenchPoi <feature-matrix|template-mutation|chart-data>");
            System.exit(2);
        }
        String lane = args[0];
        if (!lane.equals("feature-matrix") && !lane.equals("template-mutation") && !lane.equals("chart-data")) {
            System.err.println("unknown lane: " + lane);
            System.exit(2);
        }
        try {
            run(lane, new BufferedInputStream(new FileInputStream(FileDescriptor.in), 1 << 16), stdout);
        } catch (Throwable error) {
            System.err.println("apache-poi " + lane + " failed: " + error);
            error.printStackTrace(System.err);
            System.exit(1);
        }
        System.exit(0);
    }

    private static void run(String lane, InputStream input, OutputStream stdout) throws Exception {
        try (XMLSlideShow deck = new XMLSlideShow(input)) {
            switch (lane) {
                case "feature-matrix" -> {
                    // pure open/save round trip
                }
                case "template-mutation" -> mutateTemplate(deck);
                case "chart-data" -> replaceChartData(deck);
                default -> throw new IllegalArgumentException("unsupported lane: " + lane);
            }
            OutputStream out = new BufferedOutputStream(stdout, 1 << 16);
            deck.write(out);
            out.flush();
        }
    }

    // ---- template-mutation -------------------------------------------------------------

    private static void mutateTemplate(XMLSlideShow deck) {
        XSLFTable table = shapeNamed(slide(deck, 4), "PPTBenchTable", XSLFTable.class);
        replaceCellText(table.getCell(1, 1), "UPDATED-TABLE-A");
        replaceCellText(table.getCell(2, 2), "UPDATED-TABLE-B");

        XSLFTextShape bullets = shapeNamed(slide(deck, 6), "PPTBenchBullets", XSLFTextShape.class);
        List<XSLFTextParagraph> paragraphs = bullets.getTextParagraphs();
        if (paragraphs.size() < 2) {
            throw new IllegalStateException("PPTBenchBullets has no paragraph index 1");
        }
        replaceParagraphText(paragraphs.get(1), "UPDATED-BULLET");
    }

    /** Cell text becomes one paragraph; the first paragraph and its first run keep their properties. */
    private static void replaceCellText(XSLFTableCell cell, String text) {
        if (cell == null) {
            throw new IllegalStateException("table cell is missing");
        }
        List<XSLFTextParagraph> paragraphs = new ArrayList<>(cell.getTextParagraphs());
        if (paragraphs.isEmpty()) {
            cell.setText(text);
            return;
        }
        for (int i = paragraphs.size() - 1; i >= 1; i--) {
            cell.removeTextParagraph(paragraphs.get(i));
        }
        replaceParagraphText(paragraphs.get(0), text);
    }

    /** Sets the text of the paragraph's first run and drops the remaining runs. */
    private static void replaceParagraphText(XSLFTextParagraph paragraph, String text) {
        List<XSLFTextRun> runs = new ArrayList<>(paragraph.getTextRuns());
        if (runs.isEmpty()) {
            paragraph.addNewTextRun().setText(text);
            return;
        }
        runs.get(0).setText(text);
        for (int i = runs.size() - 1; i >= 1; i--) {
            if (!paragraph.removeTextRun(runs.get(i))) {
                throw new IllegalStateException("could not remove extra text run");
            }
        }
    }

    // ---- chart-data --------------------------------------------------------------------

    private static void replaceChartData(XMLSlideShow deck) throws Exception {
        replaceCategoryChart(chart(deck, 0));
        replaceScatterChart(chart(deck, 1));
        replaceBubbleChart(chart(deck, 2));
    }

    private static void replaceCategoryChart(XSLFChart chart) {
        XDDFBarChartData data = single(chart.getChartSeries(), XDDFBarChartData.class, "bar chart");
        XDDFChartData.Series series = singleSeries(data, "Revenue");
        XDDFDataSource<?> oldCategories = series.getCategoryData();
        XDDFNumericalDataSource<? extends Number> oldValues = series.getValuesData();
        String categoryRange = oldCategories.getDataRangeReference();
        String valueRange = oldValues.getDataRangeReference();

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromArray(
                new String[] {"East", "42", "84"}, categoryRange, column(categoryRange));
        XDDFNumericalDataSource<Integer> values = numbers(new Integer[] {42, 84, 126}, valueRange, oldValues);
        series.replaceData(categories, values);
        chart.plot(data);
    }

    private static void replaceScatterChart(XSLFChart chart) {
        XDDFScatterChartData data = single(chart.getChartSeries(), XDDFScatterChartData.class, "scatter chart");
        XDDFChartData.Series series = singleSeries(data, "Trend");
        XDDFDataSource<?> oldX = series.getCategoryData();
        XDDFNumericalDataSource<? extends Number> oldY = series.getValuesData();

        XDDFNumericalDataSource<Integer> x = numbers(new Integer[] {5, 21, 55}, oldX.getDataRangeReference(), oldX);
        XDDFNumericalDataSource<Integer> y = numbers(new Integer[] {13, 34, 89}, oldY.getDataRangeReference(), oldY);
        series.replaceData(x, y);
        chart.plot(data);
    }

    private static void replaceBubbleChart(XSLFChart chart) throws Exception {
        CTPlotArea plotArea = chart.getCTChart().getPlotArea();
        if (plotArea.sizeOfBubbleChartArray() != 1) {
            throw new IllegalStateException("expected exactly one bubble chart, found " + plotArea.sizeOfBubbleChartArray());
        }
        CTBubbleChart ctBubble = plotArea.getBubbleChartArray(0);
        if (ctBubble.sizeOfAxIdArray() != 2) {
            throw new IllegalStateException("bubble chart must reference exactly two axes");
        }
        XDDFChartAxis xAxis = axis(chart, ctBubble.getAxIdArray(0).getVal());
        XDDFChartAxis yAxis = axis(chart, ctBubble.getAxIdArray(1).getVal());
        if (!(yAxis instanceof XDDFValueAxis valueAxis)) {
            throw new IllegalStateException("bubble chart Y axis is not a value axis");
        }
        XDDFBubbleChartData data = new XDDFBubbleChartData(chart, ctBubble, xAxis, valueAxis);
        XDDFBubbleChartData.Series series = (XDDFBubbleChartData.Series) singleSeries(data, "Pipeline");
        XDDFDataSource<?> oldX = series.getCategoryData();
        XDDFNumericalDataSource<? extends Number> oldY = series.getValuesData();
        if (!series.getCTBubbleSer().isSetBubbleSize()) {
            throw new IllegalStateException("bubble series has no bubbleSize reference");
        }
        XDDFNumericalDataSource<Double> oldSizes =
                XDDFDataSourcesFactory.fromDataSource(series.getCTBubbleSer().getBubbleSize());
        String sizeRange = oldSizes.getDataRangeReference();

        Integer[] sizeValues = {21, 89, 21};
        XDDFNumericalDataSource<Integer> x = numbers(new Integer[] {5, 34, 8}, oldX.getDataRangeReference(), oldX);
        XDDFNumericalDataSource<Integer> y = numbers(new Integer[] {13, 55, 13}, oldY.getDataRangeReference(), oldY);
        XDDFNumericalDataSource<Integer> sizes = numbers(sizeValues, sizeRange, oldSizes);
        series.replaceData(x, y);
        series.setBubbleSizes(sizes);
        chart.plot(data);

        // XDDFChart.plot() fills category/value columns only; write bubble sizes to the embedded
        // workbook through the same workbook object that commit() saves.
        XSSFSheet sheet = chart.getWorkbook().getSheetAt(0);
        CellReference[] cells = new AreaReference(sizeRange, SpreadsheetVersion.EXCEL2007).getAllReferencedCells();
        if (cells.length != sizeValues.length) {
            throw new IllegalStateException("bubble size range " + sizeRange + " does not hold " + sizeValues.length + " points");
        }
        for (int i = 0; i < cells.length; i++) {
            requireSheet(sheet, cells[i], sizeRange);
            XSSFRow row = sheet.getRow(cells[i].getRow());
            if (row == null) {
                row = sheet.createRow(cells[i].getRow());
            }
            XSSFCell cell = row.getCell(cells[i].getCol());
            if (cell == null) {
                cell = row.createCell(cells[i].getCol());
            }
            cell.setCellValue(sizeValues[i].doubleValue());
        }
    }

    /** Array source bound to the existing formula/column, keeping the existing number format. */
    private static XDDFNumericalDataSource<Integer> numbers(Integer[] values, String range, XDDFDataSource<?> previous) {
        XDDFNumericalDataSource<Integer> source = XDDFDataSourcesFactory.fromArray(values, range, column(range));
        source.setFormatCode(previous.getFormatCode());
        return source;
    }

    /**
     * Column index for XDDFChart.fillSheet, which always writes data rows starting at sheet row 2;
     * refuse ranges it would write to the wrong cells.
     */
    private static int column(String range) {
        AreaReference area = new AreaReference(range, SpreadsheetVersion.EXCEL2007);
        CellReference first = area.getFirstCell();
        if (first.getRow() != 1 || area.getLastCell().getCol() != first.getCol()) {
            throw new IllegalStateException("chart range " + range + " is not a single column starting at row 2");
        }
        return first.getCol();
    }

    private static void requireSheet(XSSFSheet sheet, CellReference cell, String range) {
        String sheetName = cell.getSheetName();
        if (sheetName != null && !sheetName.equals(sheet.getSheetName())) {
            throw new IllegalStateException("range " + range + " does not target workbook sheet " + sheet.getSheetName());
        }
    }

    private static XDDFChartAxis axis(XSLFChart chart, long id) {
        for (XDDFChartAxis axis : chart.getAxes()) {
            if (axis.getId() == id) {
                return axis;
            }
        }
        throw new IllegalStateException("chart axis " + id + " is missing");
    }

    private static XDDFChartData.Series singleSeries(XDDFChartData data, String expectedName) {
        if (data.getSeriesCount() != 1) {
            throw new IllegalStateException("expected one series named " + expectedName + ", found " + data.getSeriesCount());
        }
        return data.getSeries(0);
    }

    private static <T extends XDDFChartData> T single(List<XDDFChartData> all, Class<T> type, String label) {
        if (all.size() != 1 || !type.isInstance(all.get(0))) {
            throw new IllegalStateException("expected exactly one " + label + ", found " + all);
        }
        return type.cast(all.get(0));
    }

    // ---- shared lookup -----------------------------------------------------------------

    private static XSLFSlide slide(XMLSlideShow deck, int index) {
        List<XSLFSlide> slides = deck.getSlides();
        if (index >= slides.size()) {
            throw new IllegalStateException("deck has no slide index " + index);
        }
        return slides.get(index);
    }

    private static <T extends XSLFShape> T shapeNamed(XSLFSlide slide, String name, Class<T> type) {
        for (XSLFShape shape : slide.getShapes()) {
            if (name.equals(shape.getShapeName())) {
                if (!type.isInstance(shape)) {
                    throw new IllegalStateException(name + " is " + shape.getClass().getSimpleName() + ", not " + type.getSimpleName());
                }
                return type.cast(shape);
            }
        }
        throw new IllegalStateException("required shape is missing: " + name);
    }

    private static XSLFChart chart(XMLSlideShow deck, int slideIndex) {
        XSLFShape first = slide(deck, slideIndex).getShapes().get(0);
        if (!(first instanceof XSLFGraphicFrame frame) || !frame.hasChart()) {
            throw new IllegalStateException("slide index " + slideIndex + " first shape is not a chart frame");
        }
        return frame.getChart();
    }
}
