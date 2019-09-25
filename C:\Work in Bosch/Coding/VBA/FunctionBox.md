# Folder
## CheckFolderExist
```vb
Public Function CheckFolderExist(FolderPath As String) As Boolean
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    CheckFolderExist = fs.FolderExists(FolderPath)
    Set fs = Nothing
End Function
```
## DeleteExistingFolder
```vb
Public Function DeleteExistingFolder(FolderPath As String) As Boolean
    On Error GoTo EH4DeleteExistingFolder
    Dim fs, f As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfolder(FolderPath)
    f.Delete
    Set fs = Nothing
    Set f = Nothing
    DeleteExistingFolder = True
    Exit Function
EH4DeleteExistingFolder:
    MsgBox prompt:="Unexcepted errors occurred during deleting the specified folder. Please try below solutions:" & _
        vbNewLine & "1. Check the folder path;" & _
        vbNewLine & "2. Exexute it again;" & _
        vbNewLine & "3. Ask for technical support.", Title:="Error Message"
    DeleteExistingFolder = False
End Function
```
## CheckFileFolderExist
```vb
Public Function CheckFileFolderExist(FullPath As String) As Boolean
    On Error GoTo EarlyExit
    If Not Dir(FullPath, vbDirectory) = vbNullString Then
        CheckFileFolderExist = True
    End If
EarlyExit:
    On Error GoTo 0
End Function
```
## CreateFolder
```vb
Public Function CreateFolder(FolderPath As String) As Boolean
    On Error GoTo EH4CreateFolder
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateFolder (FolderPath)
    Set fs = Nothing
    CreateFolder = True
    Exit Function
EH4CreateFolder:
    MsgBox prompt:="Unexcepted errors occurred during creating the specified folder. Please try below solutions:" & _
        vbNewLine & "1. Check the folder path;" & _
        vbNewLine & "2. Exexute it again;" & _
        vbNewLine & "3. Ask for technical support.", Title:="Error Message"
    CreateFolder = False
End Function
```

# Email
## 
```vb
Public Function ValidateEmailAddress(EmailAddress As String, Optional Connector As String, Optional IgnoreBlank As Integer = 0) As Boolean
    ''//IgnoreBlank = 0 :not ignore blanks
    ''//IgnoreBlank = 1 :ignore blanks between connector, but not allowed all blanks
    ''//IgnoreBlank = 2 :ignore all blanks even there is no any text between connectors
    Dim reg As New RegExp
    With reg
        .Global = True
        .Pattern = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
    End With
    
    Dim EmailText As String  ''//remove connectors from EmailAddress and only keep text
    Dim AllTextBlank As Boolean
    EmailText = Replace(EmailAddress, Connector, "")
    If EmailText = "" Then
        AllTextBlank = True
    Else
        AllTextBlank = False
    End If
    
    If Connector <> "" And EmailAddress Like "*" & Connector & "*" Then
        Dim rng As Variant
        Dim i As Integer

        rng = Split(EmailAddress, Connector)
        
        ValidateEmailAddress = True
        Select Case IgnoreBlank
        Case Is = 0
            If AllTextBlank = False Then
                For i = 0 To UBound(rng)
                    If reg.test(rng(i)) = False Then
                        ValidateEmailAddress = False
                        Exit For
                    End If
                Next
            Else
                ValidateEmailAddress = False
            End If
        Case Is = 1
            If AllTextBlank = False Then
                For i = 0 To UBound(rng)
                    If Not (rng(i) = "" Or reg.test(rng(i))) Then
                        ValidateEmailAddress = False
                        Exit For
                    End If
                Next
            Else
                ValidateEmailAddress = False
            End If
        Case Is = 2
             If AllTextBlank = False Then
                For i = 0 To UBound(rng)
                    If Not (rng(i) = "" Or reg.test(rng(i))) Then
                        ValidateEmailAddress = False
                        Exit For
                    End If
                Next
            Else
                ValidateEmailAddress = True
            End If
        End Select
    Else
        Select Case IgnoreBlank
        Case Is = 0
            ValidateEmailAddress = reg.test(EmailAddress)
        Case Is = 1
            ValidateEmailAddress = reg.test(EmailAddress)
        Case Is = 2
            If EmailAddress = "" Then
                ValidateEmailAddress = True
            Else
                ValidateEmailAddress = reg.test(EmailAddress)
            End If
        End Select
    End If
End Function
```
```vb
Function GetMail(TargetString As String) As String
    Dim arr, arr2 As Variant
    GetMail = ""
    arr = Split(TargetString, "<")
    For i = 1 To UBound(arr)
        arr2 = Split(arr(i), ">", -1, vbBinaryCompare)
        If GetMail = "" Then
            GetMail = Trim(arr2(0))
        Else
            GetMail = GetMail & ";" & Trim(arr2(0))
        End If
    Next
End Function
```
<!--stackedit_data:
eyJoaXN0b3J5IjpbLTE2MDEyNjcwMTZdfQ==
-->