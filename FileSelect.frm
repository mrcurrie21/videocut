VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileSelect 
   Caption         =   "Import Select"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8265
   OleObjectBlob   =   "FileSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Filter_Change()
Dim wsAS As Worksheet
Set wsAS = Application.ThisWorkbook.ActiveSheet
Dim wsS As Worksheet
Set wsS = Application.ThisWorkbook.Sheets("setup")

wsAS.Range("B3") = Filter.Value
Call Path_Change
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Activate()
  MakeFormResizable
End Sub
Private Sub Browse_Click()
Dim diaFolder As FileDialog
Dim File_Path As String

Dim wsAS As Worksheet
Set wsAS = Application.ThisWorkbook.ActiveSheet
Dim wsS As Worksheet
Set wsS = Application.ThisWorkbook.Sheets("setup")

Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)

If wsAS.Range("B2") <> "" Then
    File_Path = wsAS.Range("B2")
ElseIf wsS.Range("B5") <> "" Then
    File_Path = wsS.Range("B5")
Else
    File_Path = Application.DefaultFilePath
End If
    
    With diaFolder
        .AllowMultiSelect = False
        .InitialFileName = File_Path
        .Title = "Movie File Location"
        If .Show = -1 Then
            File_Path = diaFolder.SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Path.Value = File_Path
    wsAS.Range("B2") = Path.Value
    
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub clearData_Click()
Dim answer As Integer

Dim wsAS As Worksheet
Set wsAS = Application.ThisWorkbook.ActiveSheet
Dim wsS As Worksheet
Set wsS = Application.ThisWorkbook.Sheets("setup")

answer = MsgBox("Do you want to clear old data?", vbYesNo + vbQuestion, "Clear dat")

If answer = vbYes Then
    wsAS.Range("B10:EZ10000").ClearContents
Else
    'do nothing
End If

End Sub

Private Sub import_Click()
Dim sPath As String
Dim Count As Integer
Dim fileArray() As Variant
ReDim fileArray(0)
Dim wsAS As Worksheet
Set wsAS = Application.ThisWorkbook.ActiveSheet
Dim wsS As Worksheet
Set wsS = Application.ThisWorkbook.Sheets("setup")
     
Application.Calculation = xlAutomatic
    
sPath = Path.Value & "\"
        
Call clearData_Click
      
'Make array of selected files
Count = 0
For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) = True Then
        ReDim Preserve fileArray(Count)
        fileArray(Count) = sPath & ListBox1.List(i)
        wsAS.Range("B" & Count + 10) = ListBox1.List(i)
        Count = Count + 1
    End If
Next i

If UBound(fileArray) < 1 Then
    MsgBox ("Select files in list!")
    Exit Sub
End If

Unload Me

fileArray = SortArrayAtoZ(fileArray)

End Sub

Private Sub Path_Change()
Dim sPath As String
Dim sFilter As String
Dim filterArray As Variant
ReDim filterArray(0)
Dim maxFileLen As Integer
Dim fileArray As Variant
ReDim fileArray(0)
Dim prevArray As Variant
ReDim prevArray(0)
Dim oFSO As Object
Dim oFiles As Object
Dim oFolder As Object

Dim wsAS As Worksheet
Set wsAS = Application.ThisWorkbook.ActiveSheet
Dim wsS As Worksheet
Set wsS = Application.ThisWorkbook.Sheets("setup")

Set oFSO = CreateObject("Scripting.FileSystemObject")
sPath = Path.Value
Set oFolder = oFSO.GetFolder(sPath)
Set oFiles = oFolder.Files

'update in case someone copy pastes into text field
'check for last character being slash

    wsAS.Range("B2") = Path.Value
    
'Make sure path exists
    

    If wsAS.Range("B3") <> "" Then
        sFilter = LCase(wsAS.Range("B3"))
    Else
        sFilter = ""
    End If

    filterArray = Split(sFilter, "*")
    Debug.Print UBound(filterArray)
    
    maxFileLen = 50


'Make array of all files in folder
    k = 0
    For Each File In oFiles
        If Len(File.Name) > maxFileLen Then
            maxFileLen = Len(File.Name)
        End If
        fileArray(k) = File.Name
        ReDim Preserve fileArray(LBound(fileArray) To UBound(fileArray) + 1)
        k = k + 1
    Next
    
    If sFilter = "" Then
        GoTo allfiles
    End If
    
    prevArray = fileArray
    ReDim fileArray(0)

'Process filters
    
    For Each filt In filterArray
        k = 0
        ReDim fileArray(0)
        For Each File In prevArray
            If InStr(LCase(File), filt) > 0 Then
                fileArray(k) = File
                ReDim Preserve fileArray(LBound(fileArray) To UBound(fileArray) + 1)
                k = k + 1
            End If
        Next
        ReDim Preserve prevArray(LBound(fileArray) To UBound(fileArray))
        prevArray = fileArray
    Next

allfiles:

    If UBound(fileArray) = 0 Then
        fileArray(0) = "No files found"
    Else
        ReDim Preserve fileArray(LBound(fileArray) To UBound(fileArray) - 1)
    End If

    fileArray = SortArrayAtoZ(fileArray)
    
    ListBox1.List = fileArray
    ListBox1.ColumnWidths = maxFileLen * 4
    
End Sub

Private Sub UserForm_Initialize()
Dim wsAS As Worksheet
Set wsAS = Application.ThisWorkbook.ActiveSheet
Dim wsS As Worksheet
Set wsS = Application.ThisWorkbook.Sheets("setup")

'If user defined path for show is defined, use it. If not, use default
If wsAS.Range("B2") <> "" Then
    Path.Value = wsAS.Range("B2")
Else
    Path.Value = wsS.Range("B5")
End If

'If user defined path for show is defined, use it. If not, use default
If wsAS.Range("B3") <> "" Then
    Filter.Value = wsAS.Range("B3")
Else
    Filter.Value = ""
End If

End Sub
