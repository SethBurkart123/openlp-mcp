#!/usr/bin/env python3
"""
Build script for OpenLP MCP Plugin

This script builds the plugin package with all dependencies vendored,
ready for distribution as a ZIP file that users can install directly.
"""

import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


def run_command(cmd, cwd=None):
    """Run a command and return the result."""
    try:
        result = subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True, cwd=cwd)
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        print(f"Error running command: {cmd}")
        print(f"Error: {e.stderr}")
        sys.exit(1)


def clean_build_dir():
    """Clean any existing build directory."""
    if os.path.exists('mcp'):
        print("Cleaning existing build directory...")
        shutil.rmtree('mcp')
    
    if os.path.exists('mcp-plugin.zip'):
        print("Removing existing ZIP file...")
        os.remove('mcp-plugin.zip')


def create_plugin_structure():
    """Create the plugin directory structure."""
    print("Creating plugin directory structure...")
    os.makedirs('mcp', exist_ok=True)
    os.makedirs('mcp/vendor', exist_ok=True)


def install_dependencies():
    """Install fastmcp dependency to vendor directory using uv."""
    print("Installing fastmcp dependency to vendor directory...")
    
    # Check if uv is available
    try:
        run_command("uv --version")
    except:
        print("Error: uv not found. Please install uv first: https://github.com/astral-sh/uv")
        sys.exit(1)
    
    # Install fastmcp to vendor directory
    run_command(f"uv pip install fastmcp --target mcp/vendor --system")
    print("✓ fastmcp installed to vendor directory")


def copy_plugin_files():
    """Copy plugin source files to the build directory."""
    print("Copying plugin source files...")
    
    src_dir = Path('src')
    dest_dir = Path('mcp')
    
    for file_path in src_dir.glob('**/*.py'):
        relative_path = file_path.relative_to(src_dir)
        dest_file = dest_dir / relative_path
        dest_file.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(file_path, dest_file)
        print(f"  ✓ Copied {relative_path}")


def patch_init_file():
    """Add vendor path loading to __init__.py."""
    print("Patching __init__.py to load vendored dependencies...")
    
    init_file = Path('mcp/__init__.py')
    
    # Read existing content
    content = init_file.read_text()
    
    # Add vendor path loading at the beginning (after initial comments)
    vendor_code = """
# Add vendor directory to Python path for bundled dependencies
import os
import sys
vendor_path = os.path.join(os.path.dirname(__file__), 'vendor')
if vendor_path not in sys.path:
    sys.path.insert(0, vendor_path)
"""
    
    # Find where to insert (after the docstring)
    lines = content.split('\n')
    insert_idx = 0
    
    # Skip initial comments and docstring
    in_docstring = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith('"""') or stripped.startswith("'''"):
            if not in_docstring:
                in_docstring = True
            elif stripped.endswith('"""') or stripped.endswith("'''"):
                in_docstring = False
                insert_idx = i + 1
                break
        elif not in_docstring and stripped and not stripped.startswith('#'):
            insert_idx = i
            break
        elif not in_docstring and not stripped:
            continue
    
    # Insert the vendor loading code
    lines.insert(insert_idx, vendor_code)
    
    # Write back
    init_file.write_text('\n'.join(lines))
    print("✓ __init__.py patched to load vendor dependencies")


def create_zip_package():
    """Create the ZIP package for distribution."""
    print("Creating ZIP package...")
    
    with zipfile.ZipFile('mcp-plugin.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk('mcp'):
            for file in files:
                file_path = os.path.join(root, file)
                arc_path = os.path.relpath(file_path, '.')
                zipf.write(file_path, arc_path)
                print(f"  ✓ Added {arc_path}")
    
    file_size = os.path.getsize('mcp-plugin.zip') / (1024 * 1024)
    print(f"✓ ZIP package created: mcp-plugin.zip ({file_size:.1f} MB)")


def verify_package():
    """Verify the package structure."""
    print("Verifying package structure...")
    
    required_files = [
        'mcp/__init__.py',
        'mcp/mcpplugin.py',
        'mcp/worker.py',
        'mcp/tools.py',
        'mcp/conversion.py',
        'mcp/url_utils.py',
        'mcp/vendor'
    ]
    
    missing_files = []
    for req_file in required_files:
        if not os.path.exists(req_file):
            missing_files.append(req_file)
    
    if missing_files:
        print("❌ Missing required files:")
        for file in missing_files:
            print(f"  - {file}")
        sys.exit(1)
    
    # Check vendor directory has fastmcp
    vendor_contents = os.listdir('mcp/vendor')
    if not any('fastmcp' in item or 'mcp' in item for item in vendor_contents):
        print("❌ fastmcp not found in vendor directory")
        print(f"Vendor contents: {vendor_contents}")
        sys.exit(1)
    
    print("✓ Package structure verified")


def main():
    """Main build function."""
    print("Building OpenLP MCP Plugin...")
    print("=" * 50)
    
    try:
        clean_build_dir()
        create_plugin_structure()
        install_dependencies()
        copy_plugin_files()
        patch_init_file()
        verify_package()
        create_zip_package()
        
        print("\n" + "=" * 50)
        print("✅ Build completed successfully!")
        print("\nTo install the plugin:")
        print("1. Extract mcp-plugin.zip to get the 'mcp' folder")
        print("2. Copy the 'mcp' folder to your OpenLP plugins directory")
        print("3. Restart OpenLP and enable the plugin in Settings → Manage Plugins")
        
    except KeyboardInterrupt:
        print("\n❌ Build interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Build failed: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()