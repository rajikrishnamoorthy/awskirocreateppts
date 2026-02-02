# Editing Existing Presentations

Follow this guidance when modifying or updating existing PowerPoint files.

## Opening and Loading

### Load an Existing Presentation
- Use the appropriate tool to open a .pptx file
- Verify the file loads correctly
- Check the current slide count with `get_slide_count`

### Review Current Content
- Use `get_presentation_info` to understand the structure
- Note existing layouts and formatting
- Identify slides that need modification

## Common Editing Tasks

### Reordering Slides
- Use `move_slide` to change slide order
- Plan the new sequence before moving
- Test navigation after reordering

### Modifying Content
- Update text using text box tools
- Replace images with new versions
- Update chart data if needed
- Refresh tables with current information

### Adding New Slides
- Use `add_slide` to insert new content
- Match the layout style of existing slides
- Position new slides appropriately with `move_slide`

### Removing Content
- Use `delete_slide` to remove entire slides
- Be cautious - deletions cannot be undone without re-opening

## Maintaining Consistency

### When Editing
- Match existing fonts and sizes
- Use the same color palette
- Maintain consistent spacing and alignment
- Follow established slide structure

### Quality Checks
- Review all slides after editing
- Check for orphaned content or broken layouts
- Verify all links and references
- Test any embedded media

## Saving Changes

### Save Options
- Save to the same file to update
- Save to a new file to preserve original
- Use descriptive filenames with dates/versions

### Before Saving
- Review all changes
- Check for any unintended modifications
- Confirm the correct save location
