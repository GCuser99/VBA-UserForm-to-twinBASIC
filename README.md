# VBA-UserForm-to-twinBASIC
A VBIDE add-in (complied with [twinBASIC](https://twinbasic.com/preview.html)) that converts VBA UserForms for use in **twinBASIC**.

The [twinBASIC](https://twinbasic.com/preview.html) IDE and compiler (under development) does not yet support VBA UserForms. It does however have its own excellent native Form designer and associated controls. This simple VBIDE add-in for MS Office applications converts (as much as is possible) a UserForm and its controls into a **twinBASIC** form that can be imported directly into **twinBASIC**. 

The macro queries the state of the UserForm and each of its child controls at design time and builds the closest **twinBASIC** equivalent. For non-MSForm controls or MSForm controls not supported (see below), a **twinBASIC** label control is substituted to flag the missing control. UserForm code is (at least partially) translated and exported in a format that can be imported into **twinBASIC**, along with the form. 

The resulting imported form and code may have to be tweaked in **twinBASIC** to work as desired, but at least the position and most property states will be converted, saving time and some tedious effort.

MS Forms controls supported: Label, TextBox, CommandButton, Frame, CheckBox, ComboBox, ListBox, OptionButton, Image, ScrollBar, SpinButton

MS Forms controls not yet supported: ToggleButton, TabStrip, MultiPage

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/nested_controls.png" alt="NestedControls" width=75% height=75%>

**Example**: Comparison of VBA UserForm (left) and converted twinBASIC form (right)

## Requirements:

- Windows 64-bit
- MS Office 2010 or later, 32/64-bit

## Things Yet to Do:

- Handle picture data for Image and other controls that accept it
- Create Inno installer

## Quick How-To-Use

1) Depending on the bit-ness of your Office app, copy either 32-bit or 64-bit version of files in the [dist folder] (https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/tree/main/dist) of this repo to a location of your choice.
2) Make sure to close all MS Office applications.
3) Run the appropriate register*.bat file to register the add-in.
4) Open an MS Office document that contains UserForms (such as the test.xlsm file in the [test folder] (https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/tree/main/test)
5) Open the Visual Basic Editor
6) You should see the twinBasic Tools menu item on the far right of the main menu bar.
7) If menu not visible, then click on Add-ins --> Add-in Manager
8) In Add-in Manager window, click on tbUserFormConverter, and then make sure "Loaded/Unloaded" is checked - this will toggle on the twinBasic Tools menu item.
9) Click on twinBasic Tools menu, then Convert UserForm
10) In the dialog that pops up, Select the UserForms that you want to convert, and then hit Convert button
11) You will be prompted where to save the processed twinBASIC files - there should be two resulting files per UserForm - a .tbForm and a .twin file.
12) Import the twinBASIC files into twinBASIC IDE.

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/VBIDE%20Menu.png" alt="Menu" width=35% height=35%>

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/addin%20window.png" alt="AddinManagerDialog" width=35% height=35%>
    
<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/converter%20dialog.png" alt="ConverterDialog" width=35% height=35%>

