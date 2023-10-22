# VBA-UserForm-to-twinBASIC
A simple VBE add-in for converting VBA UserForms for use in twinBASIC

The exciting new twinBASIC IDE and compiler (under development) does not yet support VBA UserForms. It does however have its own excellent native Form designer and associated controls. This simple VBE add-in for Excel converts (as much as is possible) a UserForm and its controls into a twinBASIC form that can be imported directly into twinBASIC. 

The macro queries the state of the UserForm and each of its child controls at design time and builds the closest twinBASIC equivalent. For non MS Form controls or MS Form controls not supported (see below), a twinBASIC label control is substituted. UserForm code is also exported in a format that can be imported into twinBASIC. This code will very likely have to be edited in twinBASIC to work, but at least the position and most property states will be converted, saving some tedious effort and time.

Show a few examples:

MS Forms controls supported:
- Label
- EditBox
- Command Button
- List Box
- Image
- etc, etc

MS Forms controls not yet supported:
- Toggle Button
- Tab Strip
- MultiPage

Things Yet to Do:
- Export Image source from IPicture to format readable to twinBASIC
- etc etc
