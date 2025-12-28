FROM mcr.microsoft.com/dotnet/sdk:10.0

WORKDIR /app

# Solution-level build config
COPY Directory.Build.props Directory.Build.targets rules.ruleset stylecop.json ./

# Copy project files first to leverage Docker cache
COPY OpenXmlPowerTools/OpenXmlPowerTools.csproj OpenXmlPowerTools/
COPY OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj OpenXmlPowerTools.Tests/
COPY OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj OpenXmlPowerTools.CLI/
COPY GoldenFileGenerator/GoldenFileGenerator.csproj GoldenFileGenerator/

# Restore dependencies
RUN dotnet restore OpenXmlPowerTools/OpenXmlPowerTools.csproj && \
    dotnet restore OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj && \
    dotnet restore OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj && \
    dotnet restore GoldenFileGenerator/GoldenFileGenerator.csproj

# Copy only C# sources and required test assets
COPY OpenXmlPowerTools/ OpenXmlPowerTools/
COPY OpenXmlPowerTools.Tests/ OpenXmlPowerTools.Tests/
COPY OpenXmlPowerTools.CLI/ OpenXmlPowerTools.CLI/
COPY GoldenFileGenerator/ GoldenFileGenerator/
COPY TestFiles/ TestFiles/

# Build in Release mode
RUN dotnet build OpenXmlPowerTools/OpenXmlPowerTools.csproj --no-restore --configuration Release && \
    dotnet build OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj --no-restore --configuration Release && \
    dotnet build OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj --no-restore --configuration Release && \
    dotnet build GoldenFileGenerator/GoldenFileGenerator.csproj --no-restore --configuration Release

# Default command runs tests
CMD ["dotnet", "test", "--no-build", "--configuration", "Release", "--verbosity", "normal"]
