# awskirocreateppts
custom power to create presentations
# PowerPoint Creator - Kiro Custom Power

A comprehensive Kiro custom power for creating, editing, and manipulating PowerPoint presentations with AI assistance through natural language commands.

## 📋 Overview

This custom power enables users to create professional PowerPoint presentations programmatically using the Model Context Protocol (MCP). It provides 32+ specialized tools for presentation management, slide operations, content creation, and professional design enhancement.

## 🚀 Features

### Core Capabilities
- ✅ **Create & Manage Presentations** - New presentations, templates, and file operations
- ✅ **Slide Operations** - Add, delete, duplicate, and move slides with various layouts
- ✅ **Rich Content Creation** - Text, images, tables, charts, shapes, and multimedia
- ✅ **Professional Design** - Themes, color schemes, typography, and visual enhancements
- ✅ **Advanced Formatting** - Text styling, shape formatting, and layout optimization

### Supported Content Types
- 📝 **Text Management** - Titles, bullet points, formatted text with rich styling
- 🖼️ **Images** - File uploads, URL imports, and image enhancement effects
- 📊 **Charts** - Bar, line, pie, and other chart types with data visualization
- 📋 **Tables** - Structured data with custom formatting and styling
- 🔷 **Shapes** - Geometric shapes, connectors, and custom graphics
- 🎨 **Backgrounds** - Solid colors, gradients, and image backgrounds

## 🛠️ Installation & Setup

### Prerequisites
- Python 3.6 or higher
- Kiro IDE with MCP support
- Virtual environment (recommended)

### Installation Steps

1. **Create Virtual Environment**
   ```bash
   python -m venv powerpoint-env
   source powerpoint-env/bin/activate  # On Windows: powerpoint-env\Scripts\activate
   ```

2. **Install MCP Server**
   ```bash
   pip install office-powerpoint-mcp-server
   ```

3. **Configure MCP in Kiro**
   Add to your `mcp.json` configuration:
   ```json
   {
     "mcpServers": {
       "powerpoint": {
         "command": "/path/to/powerpoint-env/bin/python",
         "args": ["-m", "ppt_mcp_server"]
       }
     }
   }
   ```

4. **Restart Kiro** to load the new power

## 📚 Available Tools

### Presentation Management
- `create_presentation` - Create new presentations
- `save_presentation` - Save to file
- `open_presentation` - Load existing files
- `get_presentation_info` - Presentation metadata

### Slide Operations
- `add_slide` - Add slides with layouts and backgrounds
- `get_slide_info` - Slide metadata and structure
- `extract_slide_text` - Text content extraction
- `optimize_slide_text` - Automatic text optimization

### Content Creation
- `populate_placeholder` - Fill slide placeholders
- `add_bullet_points` - Bulleted lists
- `manage_text` - Advanced text operations
- `manage_image` - Image insertion and enhancement
- `add_table` - Tables with formatting
- `add_chart` - Data visualization charts
- `add_shape` - Geometric shapes and graphics

### Professional Design
- `apply_professional_design` - Themes and styling
- `list_slide_templates` - Available templates
- `apply_slide_template` - Template application
- `create_slide_from_template` - Template-based slide creation
- `auto_generate_presentation` - AI-powered generation

### Advanced Features
- `manage_hyperlinks` - Link management
- `update_chart_data` - Dynamic chart updates
- `add_connector` - Shape connectors
- `manage_slide_transitions` - Slide animations
- `manage_fonts` - Typography optimization

## 🎯 Usage Examples

### Basic Presentation Creation
```python
# Create new presentation
create_presentation()

# Add title slide
add_slide(layout_index=0, title="My Presentation")

# Add content slide with bullets
add_slide(layout_index=1, title="Key Points")
add_bullet_points(slide_index=1, placeholder_idx=1, 
                 bullet_points=["Point 1", "Point 2", "Point 3"])

# Save presentation
save_presentation("my_presentation.pptx")
```

### Professional Design Application
```python
# Apply professional theme
apply_professional_design(operation="theme", color_scheme="modern_blue")

# Enhance specific slide
apply_professional_design(operation="enhance", slide_index=0, 
                         enhance_title=True, enhance_content=True)

# Optimize text for readability
optimize_slide_text(slide_index=1, auto_resize=True, optimize_spacing=True)
```

### Template-Based Creation
```python
# Create from template
create_slide_from_template(template_id="title_slide", 
                          color_scheme="corporate_gray",
                          content_mapping={"title": "My Title", "subtitle": "Subtitle"})
```

## 🎨 Design Features

### Color Schemes
- `modern_blue` - Professional blue theme
- `corporate_gray` - Business-focused gray palette
- `elegant_green` - Nature-inspired green tones
- `warm_red` - Dynamic red accents

### Template Types
- `title_slide` - Title and subtitle layout
- `text_with_image` - Content with supporting visuals
- `bullet_points` - List-based content
- `chart_slide` - Data visualization focus

## 📁 Project Structure

```
powerpoint-creator/
├── POWER.md                    # Power documentation
├── README.md                   # This file
├── mcp.json                    # MCP server configuration
├── powerpoint-env/             # Python virtual environment
└── steering/                   # Workflow guides
    ├── create-presentation.md  # Creation workflows
    ├── edit-presentation.md    # Editing workflows
    └── visual-content.md       # Visual design guides
```

## 🔧 Configuration

### MCP Server Settings
The power uses a Python-based MCP server that communicates with PowerPoint through the python-pptx library. Configuration options include:

- **Command Path**: Path to Python executable in virtual environment
- **Module**: `ppt_mcp_server` - The main server module
- **Environment Variables**: Optional logging and debug settings

### Steering Files
The power includes workflow guides in the `steering/` directory:
- **create-presentation.md** - Step-by-step creation workflows
- **edit-presentation.md** - Editing and modification patterns
- **visual-content.md** - Design and visual enhancement guides

## 🚀 Best Practices

### Presentation Creation
1. **Plan Structure First** - Outline slides before creating content
2. **Use Consistent Layouts** - Match layout types to content purpose
3. **Apply Themes Early** - Set color schemes before detailed formatting
4. **Optimize Text Regularly** - Use text optimization for readability

### Content Management
1. **Leverage Templates** - Use predefined templates for consistency
2. **Batch Operations** - Apply formatting to multiple slides at once
3. **Save Frequently** - Regular saves prevent data loss
4. **Test Presentations** - Verify output in PowerPoint after creation

### Performance Tips
1. **Virtual Environment** - Use isolated Python environment
2. **Image Optimization** - Compress images before insertion
3. **Minimal Shapes** - Avoid excessive shape complexity
4. **Template Reuse** - Create reusable template configurations

## 🤝 Contributing

### Development Setup
1. Fork the repository
2. Create feature branch
3. Set up development environment
4. Make changes and test thoroughly
5. Submit pull request with detailed description

### Areas for Contribution
- Additional chart types and data visualization options
- Enhanced template library with more design options
- Performance optimizations for large presentations
- Integration with external design assets and icon libraries
- Advanced animation and transition effects

## 📝 License

This project is licensed under the MIT License. See LICENSE file for details.

## 🐛 Troubleshooting

### Common Issues
- **Python Path Errors**: Ensure correct virtual environment path in mcp.json
- **Module Not Found**: Verify ppt_mcp_server installation
- **Permission Errors**: Check file system permissions for save operations
- **Template Errors**: Validate template IDs and content mapping

### Debug Mode
Enable debug logging by setting environment variables:
```json
{
  "env": {
    "FASTMCP_LOG_LEVEL": "DEBUG"
  }
}
```

## 📞 Support

- **Issues**: Report bugs and feature requests via GitHub Issues
- **Documentation**: Check POWER.md for detailed tool documentation
- **Community**: Join discussions in the Kiro community forums

---

*Created for Kiro IDE - Empowering developers with AI-assisted presentation creation*
