---
name: powerpoint-creator
description: Create, edit, and manipulate PowerPoint presentations with AI assistance
keywords:
  - powerpoint
  - pptx
  - presentation
  - slides
  - slide deck
  - ppt
  - microsoft powerpoint
  - create presentation
  - generate slides
---

# PowerPoint Creator Power

This power enables you to create, edit, and manipulate PowerPoint presentations using natural language commands.

## Onboarding

Before using this power, ensure you have:

1. **Python 3.6+** installed on your system
2. **Install the MCP server**:
   ```bash
   pip install office-powerpoint-mcp-server
   ```

## Available Tools

This power provides 32 tools organized into specialized modules:

### Presentation Management
- `create_presentation` - Create a new PowerPoint presentation
- `save_presentation` - Save the presentation to a file
- `get_presentation_info` - Get information about the current presentation

### Slide Operations
- `add_slide` - Add a new slide with specified layout
- `delete_slide` - Remove a slide by index
- `duplicate_slide` - Duplicate an existing slide
- `move_slide` - Move a slide to a different position
- `get_slide_count` - Get the total number of slides

### Content Management
- `add_title` - Add or update slide title
- `add_text_box` - Add a text box to a slide
- `add_bullet_points` - Add bulleted text content
- `add_image` - Insert an image from file path or URL
- `add_table` - Create and insert a table
- `add_chart` - Add charts (bar, line, pie, etc.)
- `add_shape` - Add shapes (rectangle, oval, arrow, etc.)

### Formatting
- `format_text` - Apply text formatting (font, size, color, bold, italic)
- `format_shape` - Style shapes with fill, border, effects
- `set_background` - Set slide background color or image

### Template & Layout
- `list_layouts` - Get available slide layouts
- `apply_template` - Apply a template to the presentation

## Steering

When working on presentations:

1. **Start with structure** - Plan your slides before creating content
2. **Use appropriate layouts** - Match slide layouts to content type
3. **Keep it simple** - Avoid cluttered slides; use bullet points
4. **Visual hierarchy** - Use consistent fonts, sizes, and colors
5. **Save frequently** - Always call `save_presentation` when done

## Example Workflows

### Create a Basic Presentation
```
1. Create a new presentation
2. Add a title slide with your presentation title
3. Add content slides with bullet points
4. Add images or charts as needed
5. Save the presentation
```

### Add Visual Content
```
1. Use add_image to insert graphics
2. Use add_chart for data visualization
3. Use add_table for structured data
4. Apply consistent formatting
```

## Best Practices

- Always specify clear file paths when saving
- Use descriptive titles and content
- Test the presentation after creation
- Consider your audience when designing slides
