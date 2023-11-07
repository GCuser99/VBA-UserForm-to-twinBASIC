# VBA-UserForm-to-twinBASIC
A VBIDE add-in that converts VBA UserForms for use in [twinBASIC](https://twinbasic.com/preview.html).

***

The [twinBASIC](https://twinbasic.com/preview.html) (in Beta) does not yet support VBA UserForms. It does however have its own excellent native Form designer and associated controls. This simple VBIDE add-in for MS Office applications converts (as much as is possible) a UserForm, its controls, and code into a **twinBASIC** form that can be imported directly into **twinBASIC**. 

This tool (compiled in **twinBASIC**) queries the state of the UserForm and each of its child controls at design time and builds the closest **twinBASIC** equivalent. For non-MS Forms controls or MS Forms controls not supported (see below), a **twinBASIC** Label or Frame control is substituted to flag the missing control. UserForm code is (at least partially) translated and exported in a format that can be imported into **twinBASIC**, along with the form. 

The resulting imported form and code may have to be tweaked in **twinBASIC** to work as desired, but at least the position and most property states will be converted, saving some time and tedious effort.

**MS Forms controls supported**: Label, TextBox, CommandButton, Frame, CheckBox, ComboBox, ListBox, OptionButton, Image, ScrollBar, ToggleButton, and SpinButton.

**MS Forms controls not yet supported**: TabStrip, and MultiPage.

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/nested_controls.png" alt="NestedControls" width=95% height=95%>

**Example**: Comparison of VBA UserForm (left) and converted twinBASIC form (right)

## Requirements:

- 64-bit MS Windows
- MS Office 2010 or later, 32/64-bit

## Quick How-To-Use

1) Depending on the bit-ness of your Office app, copy either 32-bit or 64-bit version of files in the [dist folder](https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/tree/main/dist) of this repo to a location of your choice.
2) Make sure to close all MS Office applications.
3) Run the appropriate register*.bat file to register the add-in.
4) Open an MS Office document that contains UserForm(s) to be converted.
5) Open the Visual Basic for Application IDE.
6) You should see the "twinBasic Tools" menu item on the far right of the main menu bar.
7) If menu not visible, then click on Add-ins --> Add-in Manager.
8) In Add-in Manager window, click on tbUserFormConverter Add-in, and then make sure "Loaded/Unloaded" is checked - this will toggle on the "twinBasic Tools" menu item.
9) Click on "twinBasic Tools" menu, then "Convert UserForm".
10) In the dialog that pops up, select the UserForm(s) that you want to convert, and then hit Convert button.
11) You will be prompted where to save the processed twinBASIC files - there should be two resulting files per UserForm - a .tbForm and a .twin file.
12) Import the twinBASIC files into twinBASIC IDE for inspection and use.

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/addin%20window.png" alt="AddinManagerDialog" width=50% height=50%>

**Add-in Manager**: You can change the load behavoir of the Add-in by clicking  Add-ins --> Add-in Manager

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/VBIDE%20Menu.png" alt="Menu" width=50% height=50%>

**VBIDE Menu**: After install, the twinBasic Tools menu item should show on the right side of the VBA menu.
    
<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/converter%20dialog.png" alt="ConverterDialog" width=50% height=50%>

**Convert Dialog**: The UserForm Converter dialog allows to select the UserForm(s) for conversion.

## Things Yet to Do:

- Create Inno installer

## Acknowledgements

- Wayne Phillips' [twinBASIC](https://twinbasic.com/preview.html) and Sample 4: MyVBEAddin
- Tim Hall's [JsonConverter](https://github.com/VBA-tools/VBA-JSON/tree/master)
- Mike Wolfe's [createGUID](https://nolongerset.com/createguid/)
- Krool's [VBCCR](https://github.com/Kr00l/VBCCR)
- R. Beltran's [ArrayList](https://github.com/Theadd/ArrayList)
