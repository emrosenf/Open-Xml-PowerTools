#!/bin/bash
set -e

IMAGE_NAME="openxml-powertools-tests"
CONTAINER_NAME="openxml-powertools-test-run"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

usage() {
    echo "Usage: $0 [options]"
    echo ""
    echo "Options:"
    echo "  --filter <filter>   Custom test filter (default: same as GitHub Actions CI)"
    echo "  --all               Run all tests (no filter)"
    echo "  --rebuild           Force rebuild of Docker image"
    echo "  --shell             Start a shell in the container instead of running tests"
    echo "  -h, --help          Show this help message"
    echo ""
    echo "Examples:"
    echo "  $0                           # Run CI tests (WcTests, FormattingChange, SmlComparer)"
    echo "  $0 --all                     # Run all tests"
    echo "  $0 --filter 'FullyQualifiedName~WcTests'  # Run only WcTests"
    echo "  $0 --shell                   # Open a shell for manual testing"
}

# Default values
FILTER="FullyQualifiedName~WcTests|FullyQualifiedName~FormattingChange|FullyQualifiedName~SmlComparer"
REBUILD=false
SHELL_MODE=false
NO_FILTER=false

# Parse arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        --filter)
            FILTER="$2"
            shift 2
            ;;
        --all)
            NO_FILTER=true
            shift
            ;;
        --rebuild)
            REBUILD=true
            shift
            ;;
        --shell)
            SHELL_MODE=true
            shift
            ;;
        -h|--help)
            usage
            exit 0
            ;;
        *)
            echo -e "${RED}Unknown option: $1${NC}"
            usage
            exit 1
            ;;
    esac
done

# Check if Docker is running
if ! docker info > /dev/null 2>&1; then
    echo -e "${RED}Error: Docker is not running. Please start Docker and try again.${NC}"
    exit 1
fi

# Build the image if needed
if $REBUILD || ! docker image inspect "$IMAGE_NAME" > /dev/null 2>&1; then
    echo -e "${YELLOW}Building Docker image...${NC}"
    docker build -t "$IMAGE_NAME" .
    echo -e "${GREEN}Docker image built successfully.${NC}"
else
    echo -e "${GREEN}Using existing Docker image. Use --rebuild to force rebuild.${NC}"
fi

# Clean up any existing container with the same name
docker rm -f "$CONTAINER_NAME" 2>/dev/null || true

if $SHELL_MODE; then
    echo -e "${YELLOW}Starting shell in container...${NC}"
    docker run -it --rm --name "$CONTAINER_NAME" "$IMAGE_NAME" /bin/bash
else
    echo -e "${YELLOW}Running tests...${NC}"

    if $NO_FILTER; then
        docker run --rm --name "$CONTAINER_NAME" "$IMAGE_NAME" \
            dotnet test OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj \
            --no-build --configuration Release --verbosity normal
    else
        docker run --rm --name "$CONTAINER_NAME" "$IMAGE_NAME" \
            dotnet test OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj \
            --no-build --configuration Release --verbosity normal \
            --filter "$FILTER"
    fi

    echo -e "${GREEN}Tests completed.${NC}"
fi
