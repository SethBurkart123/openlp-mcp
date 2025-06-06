> [!IMPORTANT]
> Please note, currently I haven't gotten it to work outside of building the plugin bundled with Openlp. Check out the progress [here](https://gitlab.com/openlp/openlp/-/issues/2072).

# OpenLP MCP Plugin

A powerful OpenLP plugin that provides Model Context Protocol (MCP) server functionality, allowing AI models to fully control OpenLP including automated service creation.

## 🚀 Features

- **Complete Service Management**: Create, load, save, and manage OpenLP services programmatically
- **Smart Media Support**: Handle images, videos, and audio with automatic plugin routing
- **PowerPoint/PDF Integration**: Auto-convert presentations with LibreOffice or fallback converters
- **Live Control**: Real-time slide navigation and theme management
- **URL Downloads**: Support for remote media files, services, and theme backgrounds
- **AI-Friendly API**: Structured MCP interface for seamless AI integration

## 📋 Requirements

- OpenLP 3.1+
- No additional Python packages required (all dependencies bundled)

## 📦 Installation

### Method 1: Download from Releases (Recommended)

1. Go to the [Releases page](../../releases)
2. Download the latest `mcp-plugin.zip` file
3. Extract the ZIP to get the `mcp` folder
4. Copy the `mcp` folder to your OpenLP plugins directory:

   **Windows:**
   ```
   C:\Program Files\OpenLP\plugins\mcp\
   ```

   **Linux:**
   ```
   /usr/share/openlp/plugins/mcp/
   # OR for user installation:
   ~/.local/share/openlp/plugins/mcp/
   ```

   **macOS:**
   ```
   Right-click OpenLP.app → Show Package Contents → Contents/MacOS/plugins/mcp/
   ```

5. Restart OpenLP
6. Go to **Settings → Manage Plugins**
7. Find "MCP" in the plugin list and check the box to enable it
8. Click **OK** to save

### Method 2: Build from Source

If you want to build the plugin yourself:

```bash
# Clone the repository
git clone https://github.com/yourusername/openLP_mcp_plugin.git
cd openLP_mcp_plugin

# Build using Python script (requires uv)
python build.py

# OR build using shell script (requires uv and zip)
./package.sh

# This creates mcp-plugin.zip which you can install as above
```

## 🔧 Development Setup

For development and testing:

```bash
# Install uv if you haven't already
curl -LsSf https://astral.sh/uv/install.sh | sh

# Clone and build
git clone https://github.com/yourusername/openLP_mcp_plugin.git
cd openLP_mcp_plugin
python build.py
```

## 🤖 Usage with AI Models

The plugin starts an MCP server on `http://127.0.0.1:8765` that AI models can connect to. The server provides tools for:

### Service Management
- `create_service` - Create new services with structured data
- `add_service_item` - Add songs, images, videos, presentations to services
- `load_service` - Load existing service files
- `save_service` - Save services to specific locations
- `get_service_items` - Retrieve current service structure

### Media Handling
- `add_image` - Add images with automatic sizing and theme support
- `add_video` - Add videos with playback controls
- `add_audio` - Add audio files to services
- `add_presentation` - Handle PowerPoint/PDF files with auto-conversion

### Live Control
- `go_to_slide` - Navigate to specific slides during live presentation
- `next_slide` / `previous_slide` - Control slide progression
- `get_current_slide` - Check current slide position

### Example AI Prompt
```
"Create a church service with the attached PowerPoint presentation and background music. Set up the service order with welcome, songs, sermon slides, and closing prayer."
```

## 🛠 Supported File Types

### Media Files
- **Images**: JPG, PNG, GIF, BMP, TIFF, WebP
- **Videos**: MP4, AVI, MOV, MKV, WMV, FLV
- **Audio**: MP3, WAV, OGG, FLAC, M4A

### Presentations
- **PowerPoint**: PPT, PPTX (converted to PDF automatically)
- **PDF**: Direct support for PDF presentations
- **LibreOffice**: ODP, ODS, ODT (if LibreOffice available)

### Remote URLs
- Direct download support for all media types from HTTP/HTTPS/FTP URLs
- Intelligent content-type detection
- Works with modern hosting services (Unsplash, CDNs, etc.)

## 🔍 Troubleshooting

### Plugin Not Loading
1. Ensure you copied the entire `mcp` folder (not just its contents)
2. Check that the folder is in the correct plugins directory
3. Verify OpenLP was restarted after installation
4. Look for error messages in OpenLP's log files

### MCP Server Not Starting
1. Check OpenLP logs for MCP-related errors
2. Ensure no other service is using port 8765
3. Verify the plugin is enabled in Manage Plugins

### PowerPoint Conversion Issues
1. Install LibreOffice for best conversion quality
2. The plugin will fall back to python-pptx + reportlab if LibreOffice unavailable
3. Large presentations may cause temporary GUI responsiveness issues

### Missing Dependencies
- All dependencies are bundled - no manual installation required
- If you see import errors, rebuild the plugin package

## 📝 Plugin Structure

```
mcp/
├── __init__.py          # Plugin initialization with vendor loading
├── mcpplugin.py         # Main plugin class
├── worker.py            # Qt worker for thread-safe operations
├── tools.py             # MCP tools and server setup
├── conversion.py        # PowerPoint/PDF conversion utilities
├── url_utils.py         # URL download and file handling
└── vendor/              # Bundled dependencies
    └── fastmcp/         # MCP server library
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test with the build script
5. Submit a pull request

## 📄 License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

## 🐛 Issues & Support

- Report bugs in the [Issues section](../../issues)
- For OpenLP-specific questions, visit the [OpenLP forums](https://forums.openlp.org/)
- For MCP protocol questions, see the [MCP documentation](https://modelcontextprotocol.io/)

## ⚡ Advanced Features

### URL Support
The plugin automatically handles remote files:
```python
# These all work automatically:
add_image("https://example.com/image.jpg")
add_video("https://cdn.example.com/video.mp4")
load_service("https://mysite.com/service.osz")
```

### Theme Integration
- Automatically applies appropriate themes based on content type
- Supports theme background image URLs
- Maintains theme consistency across service items

### Batch Operations
- Process multiple files at once
- Efficient handling of large presentations
- Automatic cleanup of temporary files

### Error Recovery
- Graceful handling of network issues
- Fallback conversion methods for presentations
- Detailed logging for troubleshooting

---

**Made with ❤️ for the OpenLP community**
