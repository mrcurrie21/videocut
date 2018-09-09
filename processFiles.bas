Attribute VB_Name = "processFiles"
Function get_sec(time_str As String)
    timeArray = Split(time_str, ":")
    get_sec = CDbl(timeArray(0)) * 3600 + CDbl(timeArray(1)) * 60 + CDbl(timeArray(2))
End Function
Sub psCreateBatch()
Dim filePath As String
Dim ffPath As String
Dim fileName As String
Dim segName As String
Dim outString As String
Dim startTime As Double
Dim endTime As Double
Dim duration As Double
Dim segArray As Variant
ReDim segArray(0)

Dim wsAS As Worksheet
Set wsAS = Application.ThisWorkbook.ActiveSheet
Dim wsS As Worksheet
Set wsS = Application.ThisWorkbook.Sheets("setup")

ffPath = wsS.Range("B4")

If wsAS.Range("B2") <> "" Then
    filePath = wsAS.Range("B2")
ElseIf wsS.Range("B5") <> "" Then
    filePath = wsS.Range("B5")
Else
    filePath = Application.DefaultFilePath
End If

outfile = filePath & "\ffcut.bat"
Open outfile For Output As #1

'Loop files
    i = 10
    Do While wsAS.Range("B" & i) <> ""
        fileName = wsAS.Range("B" & i)
        
        'Loop seg once check seg count
        j = 6
        Do While wsAS.Cells(i, j) <> ""
            j = j + 1
        Loop
        
        If j Mod 2 Then
            GoTo segError
        End If
        
        'Add ffmpeg command to add keyframes every second
        keyName = CreateObject("Scripting.FilesystemObject").getBasename(fileName) _
            & "_keyed." & Split(fileName, ".")(UBound(Split(fileName, ".")))
        
        outString = Chr(34) & ffPath & Chr(34) & " -i " & Chr(34) & fileName & Chr(34)
        outString = outString & " -qscale 0 -g 1 " & Chr(34) & keyName & Chr(34)
        Print #1, outString
        
        
        'Loop Cuts
        j = 6
        s = 1
        Do While wsAS.Cells(i, j) <> ""
            
            
            If j Mod 2 = 0 Then
                startTime = get_sec(wsAS.Cells(i, j))
            Else
                endTime = get_sec(wsAS.Cells(i, j))
                duration = endTime - startTime
                                                
                'make sure seg duration is positive
                If duration < 0 Then
                    GoTo durError
                End If
                
                segName = CreateObject("Scripting.FilesystemObject").getBasename(fileName) _
                    & "_" & CStr(s) & "." & Split(fileName, ".")(UBound(Split(fileName, ".")))
                
                'add segment to segArray
                ReDim Preserve segArray(0 To UBound(segArray) + 1)
                s = s + 1
                
                outString = Chr(34) & ffPath & Chr(34) & " -i " & Chr(34) & keyName & Chr(34)
                outString = outString & " -acodec copy -vcodec copy -ss "
                outString = outString & CStr(startTime) & " -t " & duration
                outString = outString & " -y " & Chr(34) & segName & Chr(34)
                
                Print #1, outString
                
            End If
        
            j = j + 1
        Loop
        
        i = i + 1
    Loop

Close #1
Exit Sub

segError:
MsgBox (fileName & " has an odd number of time stamps." & _
    "Please check number of time stamps.")
Close #1
Exit Sub

durError:
MsgBox (fileName & " - Segment " & CStr(s) & _
    " has a negative duration. Check start and end time")
Close #1
Exit Sub


Close #1

End Sub
