#!/bin/bash
# Generate golden files for TypeScript port TDD
# Uses Docker to run the C# GoldenFileGenerator with .NET 10

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# Build Docker image if needed
if [[ "$(docker images -q openxml-powertools:latest 2>/dev/null)" == "" ]] || [[ "$1" == "--rebuild" ]]; then
    echo "Building Docker image..."
    docker build -t openxml-powertools:latest .
fi

# Create output directory
mkdir -p redline-js/tests/golden

# Run the golden file generator
echo "Generating golden files..."
docker run --rm \
    -v "$SCRIPT_DIR/TestFiles:/app/TestFiles:ro" \
    -v "$SCRIPT_DIR/redline-js/tests/golden:/app/redline-js/tests/golden" \
    openxml-powertools:latest \
    dotnet run --project GoldenFileGenerator --no-build --configuration Release

echo "Done! Golden files written to redline-js/tests/golden/"
