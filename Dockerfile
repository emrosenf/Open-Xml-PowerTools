FROM mcr.microsoft.com/dotnet/sdk:10.0

WORKDIR /app

# Copy project files first to leverage Docker cache
COPY OpenXmlPowerTools/OpenXmlPowerTools.csproj OpenXmlPowerTools/
COPY OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj OpenXmlPowerTools.Tests/
COPY OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj OpenXmlPowerTools.CLI/
COPY OpenXmlPowerTools.Packaging/OpenXmlPowerTools.Packaging.csproj OpenXmlPowerTools.Packaging/
COPY OpenXmlPowerTools.Packaging.Tests/OpenXmlPowerTools.Packaging.Tests.csproj OpenXmlPowerTools.Packaging.Tests/
COPY GoldenFileGenerator/GoldenFileGenerator.csproj GoldenFileGenerator/

# Restore dependencies
RUN dotnet restore OpenXmlPowerTools/OpenXmlPowerTools.csproj && \
    dotnet restore OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj && \
    dotnet restore OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj && \
    dotnet restore OpenXmlPowerTools.Packaging/OpenXmlPowerTools.Packaging.csproj && \
    dotnet restore OpenXmlPowerTools.Packaging.Tests/OpenXmlPowerTools.Packaging.Tests.csproj && \
    dotnet restore GoldenFileGenerator/GoldenFileGenerator.csproj

# Copy the rest of the source code
COPY . .

# Build in Release mode
RUN dotnet build OpenXmlPowerTools/OpenXmlPowerTools.csproj --no-restore --configuration Release && \
    dotnet build OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj --no-restore --configuration Release && \
    dotnet build OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj --no-restore --configuration Release && \
    dotnet build OpenXmlPowerTools.Packaging/OpenXmlPowerTools.Packaging.csproj --no-restore --configuration Release && \
    dotnet build OpenXmlPowerTools.Packaging.Tests/OpenXmlPowerTools.Packaging.Tests.csproj --no-restore --configuration Release && \
    dotnet build GoldenFileGenerator/GoldenFileGenerator.csproj --no-restore --configuration Release

# Default command runs tests
CMD ["dotnet", "test", "--no-build", "--configuration", "Release", "--verbosity", "normal"]
