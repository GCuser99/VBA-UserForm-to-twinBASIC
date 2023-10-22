# VBA-UserForm-to-twinBASIC
A simple VBE add-in for converting VBA UserForms for use in [twinBASIC](https://twinbasic.com/preview.html).

The [twinBASIC](https://twinbasic.com/preview.html) IDE and compiler (under development) does not yet support VBA UserForms. It does however have its own excellent native Form designer and associated controls. This simple VBE add-in for Excel converts (as much as is possible) a UserForm and its controls into a **twinBASIC** form that can be imported directly into **twinBASIC**. 

The macro queries the state of the UserForm and each of its child controls at design time and builds the closest **twinBASIC** equivalent. For non MS Form controls or MS Form controls not supported (see below), a **twinBASIC** label control is substituted to flag the missing control. UserForm code is also exported in a format that can be imported into **twinBASIC**. This code will very likely have to be edited in **twinBASIC** to work, but at least the position and most property states will be converted, saving time and some tedious effort.

<img src="[https://github.com/GCuser99/SeleniumVBA/blob/main/dev/logo/logo.png](https://github.com/GCuser99/VBA-UserForm-to-twinBASIC/blob/main/images/nested_controls.png)" alt="NestedControls" width=33% height=33%>

MS Forms controls supported:
- Label
- Text Box
- Command Button
- Frame
- Check Box
- Combo Box
- List Box
- Option Button
- Image
- Scroll Bar

MS Forms controls not yet supported:
- Toggle Button
- Tab Strip
- Multi Page

Things Yet to Do:
- Export Image source from IPicture to format readable to **twinBASIC**.
- etc etc
