# PowerPoint Presentation Updater

This Python script is designed to update specific values within a PowerPoint presentation. It utilizes the `python-pptx` library to access and modify the content of a PowerPoint presentation.

## Prerequisites

Before using the script, ensure the following:

- **Presentation in PPT Format**: Make sure you have the presentation in PowerPoint format (PPT).
- **Placeholders in the Presentation**: Ensure that there are placeholders in the presentation slides that need to be replaced with new values. Placeholders can be denoted by using a specific symbol or character (e.g., `#`).
- **New Data in Dictionary Form**: Prepare the new data that needs to be updated in the presentation in the form of a dictionary.

## Usage

1. Place your PowerPoint presentation file (`YourPresentation.pptx`) in the same directory as the script.
2. Update the script as needed, including specifying the new data in the appropriate format.
3. Run the script.

The script will update the specified values in the PowerPoint presentation and save the updated version as `Updated Presentation.pptx`.

## Functionality

- **Load the Presentation**: The script loads the PowerPoint presentation using the `python-pptx` library.
- **View Slides**: It provides functionality to view the slides along with their titles and content.
- **Update Specific Slide Values**: The script updates specific values within each slide based on the provided data.
- **Preserve Original Formatting**: The script preserves the original formatting of the slides while updating the content.
- **Save Updated Presentation**: The updated presentation is saved as `Updated Presentation.pptx`.

## Examples

### Updating Slides with Slide-Specific Data

You can update specific slides with slide-specific data. For example, suppose you want to update slide 4 with the number of duplicated songs and the total number of songs removed, and slide 10 with the number of topics. You can specify these data in the script and run it to update the presentation accordingly.

### Updating Slides with General Data

You can also update slides with general data that applies to multiple slides. For instance, if you have placeholders like "Number of duplicated songs," "Total number of songs removed," and "Number of topics" across different slides, you can provide the new data in a dictionary and update all relevant slides simultaneously.
