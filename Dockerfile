FROM mcr.microsoft.com/dotnet/sdk:10.0

WORKDIR /app

# Copy project files first to leverage Docker cache
COPY OpenXmlPowerTools/OpenXmlPowerTools.csproj OpenXmlPowerTools/
COPY OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj OpenXmlPowerTools.Tests/
COPY OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj OpenXmlPowerTools.CLI/

# Restore dependencies
RUN dotnet restore OpenXmlPowerTools/OpenXmlPowerTools.csproj && \
    dotnet restore OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj && \
    dotnet restore OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj

# Copy the rest of the source code
COPY . .

# Build in Release mode
RUN dotnet build OpenXmlPowerTools/OpenXmlPowerTools.csproj --no-restore --configuration Release && \
    dotnet build OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj --no-restore --configuration Release && \
    dotnet build OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj --no-restore --configuration Release

# Default command runs the same test filter as GitHub Actions
CMD ["dotnet", "test", "OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj", "--no-build", "--configuration", "Release", "--verbosity", "normal", "--filter", "FullyQualifiedName~WcTests|FullyQualifiedName~FormattingChange|FullyQualifiedName~SmlComparer|FullyQualifiedName~PmlComparer"]
