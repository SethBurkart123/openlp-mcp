#!/bin/bash

# OpenLP MCP Plugin Packaging Script
# This script creates a distributable ZIP package of the plugin

set -e

echo "🔧 OpenLP MCP Plugin Packaging Script"
echo "======================================"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print colored output
print_status() {
    echo -e "${GREEN}✓${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}⚠${NC} $1"
}

print_error() {
    echo -e "${RED}❌${NC} $1"
}

# Check if uv is available
if ! command -v uv &> /dev/null; then
    print_error "uv is required but not installed."
    echo "Please install uv: https://github.com/astral-sh/uv"
    exit 1
fi

print_status "uv found: $(uv --version)"

# Clean up any existing build
if [ -d "mcp" ]; then
    print_status "Cleaning existing build directory..."
    rm -rf mcp
fi

if [ -f "mcp-plugin.zip" ]; then
    print_status "Removing existing ZIP file..."
    rm -f mcp-plugin.zip
fi

# Create plugin directory structure
print_status "Creating plugin directory structure..."
mkdir -p mcp/vendor

# Install fastmcp dependency to vendor directory
print_status "Installing fastmcp dependency to vendor directory..."
uv pip install fastmcp --target mcp/vendor --system

# Copy plugin source files
print_status "Copying plugin source files..."
cp -r src/* mcp/

# Add vendor path loading to __init__.py
print_status "Patching __init__.py to load vendored dependencies..."
cat << 'EOF' >> mcp/__init__.py

# Add vendor directory to Python path for bundled dependencies
import os
import sys
vendor_path = os.path.join(os.path.dirname(__file__), 'vendor')
if vendor_path not in sys.path:
    sys.path.insert(0, vendor_path)
EOF

# Verify package structure
print_status "Verifying package structure..."
required_files=("mcp/__init__.py" "mcp/mcpplugin.py" "mcp/worker.py" "mcp/tools.py" "mcp/conversion.py" "mcp/url_utils.py")

for file in "${required_files[@]}"; do
    if [ ! -f "$file" ]; then
        print_error "Missing required file: $file"
        exit 1
    fi
done

if [ ! -d "mcp/vendor" ]; then
    print_error "Vendor directory not found"
    exit 1
fi

# Check if fastmcp is in vendor directory
if ! ls mcp/vendor/ | grep -q "fastmcp\|mcp"; then
    print_error "fastmcp not found in vendor directory"
    echo "Vendor contents:"
    ls -la mcp/vendor/
    exit 1
fi

print_status "Package structure verified"

# Create ZIP package
print_status "Creating ZIP package..."
zip -r mcp-plugin.zip mcp/

# Get file size
file_size=$(du -h mcp-plugin.zip | cut -f1)
print_status "ZIP package created: mcp-plugin.zip ($file_size)"

echo ""
echo "======================================"
echo -e "${GREEN}🎉 Build completed successfully!${NC}"
echo ""
echo "To install the plugin:"
echo "1. Extract mcp-plugin.zip to get the 'mcp' folder"
echo "2. Copy the 'mcp' folder to your OpenLP plugins directory:"
echo "   • Windows: C:\\Program Files\\OpenLP\\plugins\\"
echo "   • Linux: /usr/share/openlp/plugins/ or ~/.local/share/openlp/plugins/"
echo "   • macOS: Right-click OpenLP.app → Show Package Contents → Contents/MacOS/plugins/"
echo "3. Restart OpenLP and enable the plugin in Settings → Manage Plugins"
echo ""
echo "The plugin includes all dependencies and requires no additional setup!"