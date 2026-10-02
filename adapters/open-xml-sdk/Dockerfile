FROM mcr.microsoft.com/dotnet/sdk:9.0.318 AS build
ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1
WORKDIR /src
COPY PptBenchOpenXml.csproj ./
RUN dotnet restore PptBenchOpenXml.csproj --source https://api.nuget.org/v3/index.json
COPY Program.cs ./
RUN dotnet publish PptBenchOpenXml.csproj --configuration Release --no-restore --output /out

FROM mcr.microsoft.com/dotnet/runtime:9.0.20
LABEL org.pptbench.adapter="open-xml-sdk" \
      org.pptbench.library="DocumentFormat.OpenXml" \
      org.pptbench.library-version="3.5.1" \
      org.pptbench.runtime=".NET 9.0.20"
ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1
WORKDIR /app
COPY --from=build /out/ ./
USER $APP_UID
ENTRYPOINT ["dotnet", "/app/PptBenchOpenXml.dll"]
