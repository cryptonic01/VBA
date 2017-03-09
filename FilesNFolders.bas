Attribute VB_Name = "FilesNFolders"
Public j As Long


Sub folder_recursion(fso As Object, fPath As String, fileQ As Boolean)
    
    Dim buff_fold, folderz, filez As Object
    
    Set buff_fold = fso.GetFolder(fPath)
    
    For Each folderz In buff_fold.SubFolders
        
        If fileQ = False Then
            Cells(6 + j, 1).Value = folderz.Path
            j = j + 1
        End If
        
        Call folder_recursion(fso, folderz.Path, fileQ)
        
    Next
    
    For Each filez In buff_fold.Files
        
        If fileQ = True Then
            Cells(6 + j, 1).Value = filez.Name
            'Cells(6 + j, 2).Value = filez.Path
            'Cells(6 + j, 3).Value = filez.DateLastModified
            j = j + 1
        End If
        
    Next

End Sub

Sub recursion_listFolders()
    
    Dim foldPath As String
    Dim fileSys As Object
    
    j = 0
    
    foldPath = Cells(1, 1).Value
    
    Set fileSys = CreateObject("Scripting.FileSystemObject")
    
    Call folder_recursion(fileSys, foldPath, False)
    
End Sub

Sub recursion_listFiles()
    
    Dim foldPath As String
    Dim fileSys As Object
    j = 0
    foldPath = Cells(1, 1).Value
    
    Set fileSys = CreateObject("Scripting.FileSystemObject")
    
    Call folder_recursion(fileSys, foldPath, True)
    
End Sub

Sub create_folderz()
    
    Dim input_fold, buff_fold As String
    
    input_fold = Cells(1, 4).Value
    
    Dim fso, destination_folder, creator As Object
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set destination_folder = fso.GetFolder(input_fold)
    
    i = 0
    
    Do Until IsEmpty(Cells(6 + i, 3).Value)
    
        buff_fold = Cells(6 + i, 3).Value
        Cells(6 + i, 4).Value = input_fold + "\" + buff_fold
        
        If fso.FolderExists(input_fold + "\" + buff_fold) Then
        
        Else
            Set creator = fso.CreateFolder(input_fold + "\" + buff_fold)
        End If
    
        i = i + 1
        
        If i > 50 Then
            Exit Sub
        End If
        
    Loop
    
End Sub

Sub compare_List()

    Dim i, j, k As Long
    
    i = 0
    j = 0
    k = 0
    
    Do Until IsEmpty(Cells(6 + i, 1).Value)
        
        j = 0
        
        Do Until IsEmpty(Cells(6 + j, 2).Value)
            
            If Cells(6 + i, 1).Value = Cells(6 + j, 2).Value Then
                
                Range(Cells(6 + j, 2).Address).Delete Shift:=xlUp
                GoTo Skipper
            
            End If
            
            If IsEmpty(Cells(6 + j + 1, 2).Value) Then
                Cells(6 + k, 3).Value = Cells(6 + i, 1).Value
                k = k + 1
            End If
            
            j = j + 1
        Loop
Skipper:
        i = i + 1
    Loop

End Sub
