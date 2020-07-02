Sub CreateCheckBoxes()

'Declare variables
Dim c As Range
Dim chkBox As CheckBox
Dim ansBoxDefault As Long
Dim chkBoxRange As Range
Dim chkBoxDefault As Boolean

'Ingore errors if user clicks Cancel or X
On Error Resume Next

'Use Input Box to select cells
Set chkBoxRange = Application.InputBox(Prompt:="Select cell range", _
    Title:="Create checkboxes", Type:=8)

'Exit the code if user clicks Cancel or X
If Err.Number <> 0 Then Exit Sub

'Use MessageBox to select checked or unchecked
ansBoxDefault = MsgBox("Should the boxes be checked?", vbYesNoCancel, _
    "Create checkboxes")
If ansBoxDefault = vbYes Then chkBoxDefault = True
If ansBoxDefault = vbNo Then chkBoxDefault = False
If ansBoxDefault = vbCancel Then Exit Sub

'Turn error checking back on
On Error GoTo 0

'Loop through each cell in the selected cells
For Each c In chkBoxRange

    'Create the checkbox
    Set chkBox = chkBoxRange.Parent.CheckBoxes.Add(0, 1, 1, 0)

    With chkBox

        'Set the position of the checkbox based on the cell
        .Top = c.Top + c.Height / 2 - chkBox.Height / 2
        .Left = c.Left + c.Width / 2 - chkBox.Width / 2

        'Set the name of the checkbox based on the cell address
        .Name = c.Address

        'Set the linked cell to the cell with the checkbox
        .LinkedCell = c.Offset(0, 0).Address(external:=True)

        'Enable the checkBox to be used when worksheet protection applied
        .Locked = False

        'Set the caption to blank
        .Caption = ""

    End With

    'Set the cell to the default value
    c.Value = chkBoxDefault

    'Hide the value in the cell with Number Formatting
    c.NumberFormat = ";;;"

Next c

End Sub