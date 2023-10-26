# VBA-UserForm-to-twinBASIC
A VBIDE add-in (complied with [twinBASIC](https://twinbasic.com/preview.html)) that converts VBA UserForms for use in **twinBASIC**.

The [twinBASIC](https://twinbasic.com/preview.html) IDE and compiler (under development) does not yet support VBA UserForms. It does however have its own excellent native Form designer and associated controls. This simple VBIDE add-in for MS Office applications converts (as much as is possible) a UserForm and its controls into a **twinBASIC** form that can be imported directly into **twinBASIC**. 

The macro queries the state of the UserForm and each of its child controls at design time and builds the closest **twinBASIC** equivalent. For non-MSForm controls or MSForm controls not supported (see below), a **twinBASIC** label control is substituted to flag the missing control. UserForm code is also exported in a format that can be imported into **twinBASIC**. This code may have to be edited in **twinBASIC** to work, but at least the position and most property states will be converted, saving time and some tedious effort.

<img src="https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/nested_controls.png" alt="NestedControls" width=75% height=75%>

MS Forms controls supported:
- Label
- TextBox
- CommandButton
- Frame
- CheckBox
- ComboBox
- ListBox
- OptionButton
- Image
- ScrollBar
- SpinButton

MS Forms controls not yet supported:
- ToggleButton
- TabStrip
- MultiPage

Things Yet to Do:
- Export Image source from IPicture to format readable to **twinBASIC**.
- etc etc
