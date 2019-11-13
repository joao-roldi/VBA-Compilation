Sub LoopAllExcelFilesInFolder()
Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    'What you want to do with each Woorkbook
      Call Rename_and_Save()
    
    'Save and Close Workbook
      wb.Close SaveChanges:=True
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Function Rename_and_Save()
Dim workbookName As String
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
ActiveSheet.Name = "BD_FAT"
workbookName = ActiveWorkbook.Name
ActiveWorkbook.SaveAs "C:\Users\joaor\Desktop\BSA - Comissões e Premiações\Relatórios SAP\" & workbookName, fileformat:=56
End Function
