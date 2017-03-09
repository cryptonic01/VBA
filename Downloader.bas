Attribute VB_Name = "Downloader"
Function urlExtract(input_data As Range) As String
    
    urlExtract = input_data.Hyperlinks(1).Address
    
End Function

Sub excelDLER()
    Dim i As Long
    i = 1
    Dim folderLoc As String
    
    Dim dlDIR As String
    dlDIR = "C:\Users\Leo.Cui\Downloads\KCentral\"
    
    Dim fso, creator As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Application.DisplayAlerts = False
    
    Do Until Cells(i, 1).Value = ""
        folderLoc = folderPather(Cells(i, 1).Value)
        Call folderPathBuilder(dlDIR, folderLoc, fso)
        Path = dlDIR + folderLoc + Cells(i, 2).Value
        
        ThisWorkbook.FollowHyperlink (urlExtract(Range(Cells(i, 2).Address)))
        Application.SendKeys "{LEFT}"
        Application.SendKeys "{RETURN}"
        ActiveWorkbook.SaveAs (Path)
        ActiveWorkbook.Close
        i = i + 1
    Loop
    Application.DisplayAlerts = True
End Sub

Sub downloader()
    
    Dim WinHTTP As Object
    Set WinHTTP = CreateObject("Microsoft.XMLHTTP")
    
    Dim outStream As Object
    Set outStream = CreateObject("ADODB.Stream")
    
    Dim fso, creator As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim i As Long
    i = Cells(3, 7).Value
    
    Dim urlz, dlDIR As String
    urlz = ""
    dlDIR = Cells(1, 1).Value
    
    Dim folderLoc As String
    folderLoc = ""
    
    outStream.Open
    outStream.Type = 1
    
    Do Until IsEmpty(Cells(6 + i, 8).Value) Or (i > Cells(2, 7).Value)
        urlz = urlExtract(Range(Cells(6 + i, 8).Address))
        WinHTTP.Open "Get", urlz, False
        WinHTTP.Send
        
        FileBuff = FreeFile()
        
        folderLoc = folderPather(Cells(6 + i, 7).Value)
        
        Call folderPathBuilder(dlDIR, folderLoc, fso)
        
        Open dlDIR & folderLoc & Cells(6 + i, 8).Value For Binary Access Write As #FileBuff
            Put #FileBuff, 1, WinHTTP.responseBody
        Close #FileBuff
        
        i = i + 1
        
    Loop
    
    outStream.Close
  
    
End Sub

Sub DLer_Alt()

    Dim i As Long
    i = 1
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim dlDIR As String
    
    Dim urlz As String
    urlz = ""
    dlDIR = "C:\Users\Leo.Cui\Downloads\KCentral\"
    
    Dim folderLoc As String
    folderLoc = ""
    
    
    Do Until IsEmpty(Cells(i, 1).Value)
        urlz = urlExtract(Range(Cells(i, 2).Address))
         
        folderLoc = folderPather(Cells(i, 1).Value)
        
        Call folderPathBuilder(dlDIR, folderLoc, fso)
        
        Cells(i, 3).Value = DownloadFile(urlz, dlDIR & folderLoc & Cells(i, 2).Value)
        
        i = i + 1
        
    Loop

End Sub


Sub file_builder()
    
    Dim fso, creator As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim i As Long
    i = Cells(3, 7).Value
    
    Dim dlDIR As String
    dlDIR = Cells(1, 1).Value
    
    Dim folderLoc As String
    folderLoc = ""
    
    Do Until IsEmpty(Cells(6 + i, 8).Value) Or (i > Cells(2, 7).Value)
        
        
        folderLoc = folderPather(Cells(6 + i, 7).Value)
        Call folderPathBuilder(dlDIR, folderLoc, fso)
        i = i + 1
        
    Loop

End Sub

Function folderPather(pbc As String) As String
    
    If Len(pbc) <= 4 Then
        folderPather = pbc & "\"
        Exit Function
    Else
        folderPather = Left(pbc, 4) & "\" & pbc & "\"
        Exit Function
    
    End If
    
End Function

Sub folderPathBuilder(home As String, newFolds As String, fso As Variant)
    
    If fso.FolderExists(home & newFolds) Then
        Exit Sub
    End If
    
    Dim newHome As String
    Dim newFolds2 As String
    Dim FoldIndex1 As Integer
    
    FoldIndex1 = InStr(newFolds, "\")
    newHome = home & Left(newFolds, FoldIndex1)
    newFolds2 = Right(newFolds, Len(newFolds) - FoldIndex1)
    
    If Not fso.FolderExists(newHome) Then
        fso.CreateFolder (newHome)
    End If
    
    If (Len(newFolds) - FoldIndex1) <> 0 Then
    
        Call folderPathBuilder(newHome, newFolds2, fso)
    End If
    
End Sub
