# VBA-UserForm-to-twinBASIC
A VBIDE add-in that converts VBA UserForms for use in [twinBASIC](https://twinbasic.com/preview.html).

***

The [twinBASIC](https://twinbasic.com/preview.html) (in Beta) does not yet support VBA UserForms. It does however have its own excellent (VB6 compatible) native Form designer and associated controls. This simple VBIDE add-in for MS Office applications converts a UserForm, its controls, images, and code into a twinBASIC form that can be imported directly into twinBASIC. 

This tool, compiled in twinBASIC, queries the state of the UserForm and each of its child controls at design time and builds the closest twinBASIC equivalent. For non-MS Forms controls or MS Forms controls not supported (see below), a twinBASIC Label or Frame control is substituted to flag the missing control. UserForm code is translated and exported in a format that can be imported into twinBASIC, along with the form. 

The resulting imported form and code may have to be tweaked in twinBASIC to work as desired, but at least the position and most property states will be converted, saving time and tedious effort.

**MS Forms controls supported**: Label, TextBox, CommandButton, Frame, CheckBox, ComboBox, ListBox, OptionButton, Image, ScrollBar, ToggleButton, and SpinButton.

**MS Forms controls not yet supported**: TabStrip, and MultiPage.

**Extract Image Resources Option**: User can optionally extract and save to file all image resources (mouse icons and picture bitmaps) that are embedded in each UserForm for later use in twinBASIC Resources.

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/nested_controls.png" alt="NestedControls" width=95% height=95%>

**Example**: Comparison of VBA UserForm (left) and converted twinBASIC form with VisualStyles=False (right)

## Requirements:

- 64-bit MS Windows
- MS Office 2010 or later, 32/64-bit

## Quick How-To-Use

1) Make sure to close all MS Office applications.
2) Download and run the [tBUserformConverterSetup.exe Inno installer](https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/tree/main/dist).
3) Open an MS Office document that contains UserForm(s) to be converted.
4) Open the Visual Basic for Application IDE.
5) You should see the "twinBASIC Tools" menu item on the far right of the main menu bar.
6) If menu not visible, then click on "Add-ins->Add-in Manager".
7) In Add-in Manager window, click on tbUserFormConverter Add-in, and then make sure "Loaded/Unloaded" is checked - this will toggle on the "twinBASIC Tools" menu item.
8) Click on "twinBASIC Tools" menu, then "Convert UserForm".
9) In the dialog that pops up, select the UserForm(s) that you want to convert, and then hit Convert button.
10) You will be prompted where to save the processed twinBASIC files - there should be two resulting files per UserForm - a .tbform and a .twin file.
11) Import the twinBASIC files into twinBASIC IDE (Project->Add->Import File(s)…).

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/addin%20window.png" alt="AddinManagerDialog" width=50% height=50%>

**Add-in Manager**: You can change the load behavoir of the Add-in by clicking  Add-ins->Add-in Manager

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/VBIDE%20Menu.png" alt="Menu" width=50% height=50%>

**VBIDE Menu**: After install, the twinBasic Tools menu item should show on the right side of the VBA menu.
    
<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/converter%20dialog.png" alt="ConverterDialog" width=50% height=50%>

**Convert Dialog**: The UserForm Converter dialog allows to select the UserForm(s) for conversion.

## Acknowledgements

- Wayne Phillips' [twinBASIC](https://twinbasic.com/preview.html) and Sample 4: MyVBEAddin
- Tim Hall's [JsonConverter](https://github.com/VBA-tools/VBA-JSON/tree/master)
- Krool's [VBCCR](https://github.com/Kr00l/VBCCR) twinBASIC package
- R. Beltran's [ArrayList](https://github.com/Theadd/ArrayList) twinBASIC package
- Jordan Russell's [Inno Setup](https://jrsoftware.org/isinfo.php) and Bill Stewart's [UninsIS](https://github.com/Bill-Stewart/UninsIS)
