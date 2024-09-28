# PowerPoint Macros

This directory contains VBA macros for Microsoft PowerPoint.

## Available Macros

1. SetDefaultColors
   Sets a predefined color scheme for the active presentation.
   Usage: Run the macro to apply the color scheme.

2. MultipleCropping
   Applies uniform cropping to multiple selected shapes.
   Usage: Select multiple shapes, then run the macro. The first selected shape's crop values will be applied to all.

3. GetRGBColor (Function)
   Returns RGB color values by name.
   Usage: color = GetRGBColor("blue")

For detailed implementation, see [PowerPointMacros.vba](./PowerPointMacros.vba).

## Setup Instructions

1. Open PowerPoint and press Alt+F11 to open the Visual Basic Editor.
2. In the Project Explorer, right-click on your presentation and select Insert > Module.
3. Copy and paste the code from PowerPointMacros.vba into the new module.
4. Save the presentation as a .pptm file (PowerPoint Macro-Enabled Presentation).

## Usage

To run a macro:
1. Go to Developer tab > Macros.
2. Select the desired macro and click Run.

## Customization

To add macros to the toolbar:
1. Right-click on the ribbon and select Customize the Ribbon.
2. Under Choose commands from, select Macros.
3. Add your macro to the desired tab or create a new group.

To create keyboard shortcuts:
1. File > Options > Customize Ribbon.
2. Click Customize next to Keyboard shortcuts.
3. Under Categories, select Macros.
4. Select your macro and assign a shortcut key.
