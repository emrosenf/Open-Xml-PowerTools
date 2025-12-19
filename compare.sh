#!/bin/bash
set -e

IMAGE_NAME="openxml-powertools-tests"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

usage() {
    echo "Usage: $0 <document1> <document2> [output]"
    echo ""
    echo "Compare two Office documents and produce a comparison result."
    echo "Supports Word (.docx), PowerPoint (.pptx), and Excel (.xlsx) files."
    echo ""
    echo "Arguments:"
    echo "  document1    Path to the original document"
    echo "  document2    Path to the modified document"
    echo "  output       Output path (optional, defaults to ./comparison-result.[ext])"
    echo ""
    echo "Options:"
    echo "  --rebuild    Force rebuild of Docker image"
    echo "  -h, --help   Show this help message"
    echo ""
    echo "Examples:"
    echo "  $0 original.docx modified.docx"
    echo "  $0 original.pptx modified.pptx result.pptx"
    echo "  $0 original.xlsx modified.xlsx"
    echo "  $0 --rebuild original.docx modified.docx"
}

REBUILD=false

# Parse arguments
POSITIONAL_ARGS=()
while [[ $# -gt 0 ]]; do
    case $1 in
        --rebuild)
            REBUILD=true
            shift
            ;;
        -h|--help)
            usage
            exit 0
            ;;
        *)
            POSITIONAL_ARGS+=("$1")
            shift
            ;;
    esac
done

set -- "${POSITIONAL_ARGS[@]}"

if [ $# -lt 2 ]; then
    echo -e "${RED}Error: Two document paths are required${NC}"
    usage
    exit 1
fi

DOC1="$1"
DOC2="$2"
OUTPUT="${3:-comparison-result.docx}"

# Validate input files exist
if [ ! -f "$DOC1" ]; then
    echo -e "${RED}Error: Document not found: $DOC1${NC}"
    exit 1
fi

if [ ! -f "$DOC2" ]; then
    echo -e "${RED}Error: Document not found: $DOC2${NC}"
    exit 1
fi

# Get absolute paths
DOC1_ABS=$(cd "$(dirname "$DOC1")" && pwd)/$(basename "$DOC1")
DOC2_ABS=$(cd "$(dirname "$DOC2")" && pwd)/$(basename "$DOC2")
OUTPUT_DIR=$(cd "$(dirname "$OUTPUT")" 2>/dev/null && pwd || pwd)
OUTPUT_NAME=$(basename "$OUTPUT")
OUTPUT_ABS="$OUTPUT_DIR/$OUTPUT_NAME"

# Check if Docker is running
if ! docker info > /dev/null 2>&1; then
    echo -e "${RED}Error: Docker is not running. Please start Docker and try again.${NC}"
    exit 1
fi

# Build the image if needed
if $REBUILD || ! docker image inspect "$IMAGE_NAME" > /dev/null 2>&1; then
    echo -e "${YELLOW}Building Docker image...${NC}"
    SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
    docker build -t "$IMAGE_NAME" "$SCRIPT_DIR"
    echo -e "${GREEN}Docker image built successfully.${NC}"
fi

echo -e "${YELLOW}Comparing documents...${NC}"
echo "  Original: $DOC1_ABS"
echo "  Modified: $DOC2_ABS"
echo "  Output:   $OUTPUT_ABS"
echo ""

# Run comparison in Docker
docker run --rm \
    -v "$DOC1_ABS:/input/doc1.docx:ro" \
    -v "$DOC2_ABS:/input/doc2.docx:ro" \
    -v "$OUTPUT_DIR:/output" \
    "$IMAGE_NAME" \
    dotnet run --project OpenXmlPowerTools.CLI/OpenXmlPowerTools.CLI.csproj --no-build --configuration Release -- \
    compare /input/doc1.docx /input/doc2.docx -o "/output/$OUTPUT_NAME"

echo ""
echo -e "${GREEN}Comparison complete!${NC}"
echo "Output saved to: $OUTPUT_ABS"
