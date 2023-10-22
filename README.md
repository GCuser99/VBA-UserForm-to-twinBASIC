# VBA-UserForm-to-twinBASIC
A simple VBE add-in for converting VBA UserForms for use in [twinBASIC](https://twinbasic.com/preview.html).

The exciting new [twinBASIC](https://twinbasic.com/preview.html) IDE and compiler (under development) does not yet support VBA UserForms. It does however have its own excellent native Form designer and associated controls. This simple VBE add-in for Excel converts (as much as is possible) a UserForm and its controls into a [twinBASIC](https://twinbasic.com/preview.html) form that can be imported directly into [twinBASIC](https://twinbasic.com/preview.html). 

The macro queries the state of the UserForm and each of its child controls at design time and builds the closest [twinBASIC](https://twinbasic.com/preview.html) equivalent. For non MS Form controls or MS Form controls not supported (see below), a [twinBASIC](https://twinbasic.com/preview.html) label control is substituted. UserForm code is also exported in a format that can be imported into [twinBASIC](https://twinbasic.com/preview.html). This code will very likely have to be edited in [twinBASIC](https://twinbasic.com/preview.html) to work, but at least the position and most property states will be converted, saving some tedious effort and time.

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
- Export Image source from IPicture to format readable to [twinBASIC](https://twinbasic.com/preview.html)
- etc etc
