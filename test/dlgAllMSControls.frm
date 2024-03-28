VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgAllMSControls 
   Caption         =   "All Supported Controls (VBA)"
   ClientHeight    =   4535
   ClientLeft      =   90
   ClientTop       =   425
   ClientWidth     =   6150
   OleObjectBlob   =   "dlgAllMSControls.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgAllMSControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
    If CheckBox1.Value Then
        TextBox1.Enabled = True
    Else
        TextBox1.Enabled = False
    End If
End Sub

Private Sub ComboBox1_Change()
    Select Case ComboBox1.List(ComboBox1.ListIndex)
    Case "Yellow"
        Frame13.BackColor = vbYellow
    Case "Magenta"
        Frame13.BackColor = vbMagenta
    Case "Cyan"
        Frame13.BackColor = vbCyan
    End Select
End Sub

Private Sub CommandButton1_Click()
    MsgBox "You cannot be trusted with the nuclear codes!"
End Sub

Private Sub ListBox1_Click()
    Label4.Caption = ListBox1.List(ListBox1.ListIndex) & " Selected"
End Sub

Private Sub OptionButton1_Click()
    Label1.Caption = "Answer A) is Incorrect!"
    Label1.Visible = True
End Sub

Private Sub OptionButton2_Click()
    Label1.Caption = "Answer B) is Incorrect!"
    Label1.Visible = True
End Sub

Private Sub ScrollBar2_Change()
    Label5.Top = ScrollBar2.Value
End Sub

Private Sub SpinButton1_Change()
    Label2.Caption = "Count: " & SpinButton1.Value
End Sub

Private Sub ToggleButton1_Click()
    ToggleButton2.Value = Not ToggleButton1.Value
    If ToggleButton1.Value Then
        ToggleButton1.Caption = "On"
        ToggleButton2.Caption = "Off"
    Else
        ToggleButton1.Caption = "Off"
        ToggleButton2.Caption = "On"
    End If
End Sub

Private Sub ToggleButton2_Click()
    ToggleButton1.Value = Not ToggleButton2.Value
    If ToggleButton2.Value Then
        ToggleButton2.Caption = "On"
        ToggleButton1.Caption = "Off"
    Else
        ToggleButton2.Caption = "Off"
        ToggleButton1.Caption = "On"
    End If
End Sub

Private Sub UserForm_Initialize()
    ListBox1.AddItem "Orange"
    ListBox1.AddItem "Apple"
    ListBox1.AddItem "Grape"
    ComboBox1.AddItem "Yellow"
    ComboBox1.AddItem "Magenta"
    ComboBox1.AddItem "Cyan"
    TextBox1.Enabled = False
    CheckBox1.Value = False
    ComboBox1.ListIndex = 0
    ComboBox1_Change
    Label1.Visible = False
    ScrollBar2.Min = Label5.Top
    ScrollBar2.Max = Label6.Top
    ScrollBar2.SmallChange = (ScrollBar2.Max - ScrollBar2.Min) / 10
    ScrollBar2.LargeChange = ScrollBar2.SmallChange
    ScrollBar2.Value = ScrollBar2.Min
End Sub
