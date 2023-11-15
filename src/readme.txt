IMPORTANT INFORMATION – PLEASE READ BEFORE INSTALLING

--------------------------------------------------------------------------------------------------------------------------
What are the system requirements for the tBUserFormConverter Add-in?

64-bit MS Windows
32- or 64-bit MS Office

--------------------------------------------------------------------------------------------------------------------------
Where should the tBUserFormConverter Add-in be installed on my User Account?

This installer will install tBUserFormConverter Add-in in any location that you select. It is recommended that it be installed in C:\Users\[user name]\AppData\Local, which is the default location, in order to prevent the Add-in DLL file from being inadvertently moved or deleted after it has been registered. Moving or deleting the DLL after installation without first properly uninstalling will break any code that references the DLL, and render it unusable.

The installer will place a tBUserFormConverter Add-in shortcut on your Desktop to provide easy access to the installation folder, uninstaller file, readme instructions, and example resources.  

--------------------------------------------------------------------------------------------------------------------------
What happens if I upgrade to a new version, or reinstall to a new location?

If upgrading to a new version, or re-installing with the same version but choosing a different install location, this installer will first unregister and then remove the older version.

--------------------------------------------------------------------------------------------------------------------------
How do I uninstall the tBUserFormConverter Add-in from my system?

To uninstall tBUserFormConverter Add-in from your system, you can run the unins000.exe file in the installation directory. This unregisters the tBUserFormConverter Add-in and removes all the original files from your system.

--------------------------------------------------------------------------------------------------------------------------
Once I have installed the tBUserFormConverter Add-in, what is the quickest way to get started?

1. Open an MS Office document that contains UserForm(s) to be converted.
2. Open the Visual Basic for Application IDE.
3. You should see the "twinBasic Tools" menu item on the far right of the main menu bar.
4. If menu is not visible, then click on “Add-ins --> Add-in Manager”.
5. In Add-in Manager window, click on tbUserFormConverter Add-in, and then make sure "Loaded/Unloaded" is checked - this will toggle on the "twinBasic Tools" menu item.
6. Click on "twinBasic Tools" menu, then "Convert UserForm".
7. In the dialog that pops up, select the UserForm(s) that you want to convert, and then hit Convert button.
8. You will be prompted where to save the processed twinBASIC files - there should be two resulting files per UserForm - a .tbform and a .twin file.
9. Import the twinBASIC files into twinBASIC IDE (Project-->Add-->Import File(s)…).

--------------------------------------------------------------------------------------------------------------------------
How do access the provided example Userform?

Use the installed tBUserFormConverter Add-in shortcut of your Desktop to navigate to the examples folder in the installation directory. You can move or copy these example files (or the entire example directory) to another location on your system. But be sure NOT to move any other files, such as the DLL, and/or the uninstaller file. Import the example .frm into your VBA project.


