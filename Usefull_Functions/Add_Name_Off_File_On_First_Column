Sub File_Loop_Example()
    'Excel VBA code to loop through files in a folder with Excel VBA

    Dim MyFolder As String, MyFile As String

    'Opens a file dialog box for user to select a folder

    With Application.FileDialog(msoFileDialogFolderPicker)
       .AllowMultiSelect = False
       .Show
       MyFolder = .SelectedItems(1)
       Err.Clear
    End With

    'stops screen updating, calculations, events, and statsu bar updates to help code run faster
    'you'll be opening and closing many files so this will prevent your screen from displaying that

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    'This section will loop through and open each file in the folder you selected
    'and then close that file before opening the next file

    MyFile = Dir(MyFolder & "\", vbReadOnly)

    Do While MyFile <> ""
        DoEvents
        On Error GoTo 0
        Workbooks.Open FileName:=MyFolder & "\" & MyFile, UpdateLinks:=False
            Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove        
            Range("A1").Value = "Planilha"
            Range("A2:A200").Value = ActiveWorkbook.Name
    0
        Workbooks(MyFile).Close SaveChanges:=True
        MyFile = Dir
    Loop

    'turns settings back on that you turned off before looping folders

    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationManual

    End Sub
