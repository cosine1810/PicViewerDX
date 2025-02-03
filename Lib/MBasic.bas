Attribute VB_Name = "MBasic"
'#If VBA7 Then
'
'    Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'#Else
'    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'#End If
'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ÉùÃ÷GetTickCountº¯Êý
Declare Function GetTickCount Lib "kernel32" () As Long

Public dicsystem As New Dictionary
Public fso As New FileSystemObject

Public MousePos As POINTAPI
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Const PI = 3.1415926

Function GetColIndex(ByVal Target As Object, ByVal col As Collection) As Long
    Dim i As Long
    i = 1
    While Not Target Is col(i)
        Inc i
    Wend
    If i > col.Count Then i = -1
    GetColIndex = i
End Function


Sub SetCol(col, ByVal index As Long, ByVal value)
    If index = col.Count Then
        col.Remove index
        col.Add value
    Else
        col.Remove index
        col.Add value, , index
    End If
End Sub

Sub RemoveCol(ByVal obj, ByVal col As Collection)
    Dim v, i
    For Each v In col
        Inc i
        If IsObject(v) Then
            If obj Is v Then
                col.Remove i
                Exit Sub
            End If
        Else
            If obj = v Then
                col.Remove i
                Exit Sub
            End If
        
        End If
    Next
End Sub


Sub ClearCol(ByVal col As Collection)
    While col.Count > 0
        col.Remove 1
    Wend
End Sub

Function iC(ParamArray params())
    Dim col As New Collection
    Dim i
    For i = 0 To UBound(params)
        col.Add params(i)
    Next
    Set iC = col
End Function
'Public gdip As New KGDIP

Function max(ByVal x As Variant, ByVal y As Variant) As Variant
If x > y Then
    max = x
Else
    max = y
End If
End Function

Function min(ByVal x As Variant, ByVal y As Variant) As Variant
    min = -max(-x, -y)
End Function

Function ABSMax(ByVal x As Variant, ByVal y As Variant) As Variant
    If Abs(x) > Abs(y) Then
        ABSMax = x
    End If
End Function

Function ABSMin(ByVal x As Variant, ByVal y As Variant) As Variant
    ABSMin = -ABSMax(-x, -y)
End Function

Function Between(ByVal src As Variant, ByVal LowBound As Variant, ByVal UpperBound As Variant, Optional ByRef result As Integer) As Variant
    Dim r
    If src < LowBound Then
        result = -1
    ElseIf src > UpperBound Then
        result = 1
    Else
        result = 0
    End If
    
    Between = max(min(src, UpperBound), LowBound)
    
End Function


Function Ceiling(ByVal value As Double) As Double
    If value >= 0 Then
        Ceiling = Fix(value)
        If value > Fix(value) Then
            Ceiling = Fix(value) + 1
        End If
    Else
        Ceiling = Fix(value)
        If value < Fix(value) Then
            Ceiling = Fix(value) - 1
        End If
    End If
End Function

Function GetFiles(ByVal folderPath As String, Optional ByVal pattern As String)
    Dim fso As New FileSystemObject
    Dim fl As File
    Dim col As New Collection
    For Each fl In fso.GetFolder(folderPath).Files
        If pattern = "" Then
            col.Add fl.path
        Else
            If RegMatch(fl.Name, pattern, True) Then
                col.Add fl.path
            End If
        End If
        
    Next
    Set GetFiles = col
End Function

Function GetFolders(ByVal folderPath As String, Optional ByVal pattern As String, Optional ByVal reverse As Boolean = False) As Collection
    If Not fso.FolderExists(folderPath) Then Exit Function
    Dim col As New Collection, k As Folder
    For Each k In fso.GetFolder(folderPath).SubFolders
        If pattern = "" Then
            col.Add k.path
        Else
            If reverse Xor RegMatch(LCase(k.Name), pattern, True) Then
                col.Add k.path
            End If
        End If
        
    Next
    Set GetFolders = col
End Function

Sub writable(ByVal path As String)
    'ReadyFolder path
    If Not FileExists(path) = "" Then Exit Sub
    SetAttr path, vbNormal
End Sub

Sub ReadyFolder(path)
    Dim a
    a = Split(path, "\")
    Dim foldpath
    Dim i
    foldpath = a(0)
    If foldpath <> "" Then
        For i = 1 To UBound(a)
            If i = UBound(a) And InStr(1, a(i), ".") <> 0 Then
                Exit Sub
            End If
            
            foldpath = foldpath & "\" & a(i)
            
            If Not fso.FolderExists(foldpath) Then
                fso.CreateFolder foldpath
                        
            End If

        Next
    Else
        foldpath = "\" & a(1) & "\" & a(2)
        For i = 3 To UBound(a)
            If i = UBound(a) And InStr(1, a(i), ".") <> 0 Then
                Exit Sub
            End If
            foldpath = foldpath & "\" & a(i)
            
            
            If Not fso.FolderExists(foldpath) Then
                
                        fso.CreateFolder foldpath
            End If

        Next
    End If
End Sub

Sub DeleteFolder(path)
    If fso.FolderExists(path) Then
        Dim k, l
        For Each k In GetFiles(path)
            Kill k
        Next
        For Each k In GetFolders(path)
            DeleteFolder k
        Next
        RmDir path
    End If

End Sub

'Function GetFiles(ByVal folderPath As String, Optional ByVal pattern As String)
'    Dim fso As New FileSystemObject
'    Dim fl As file
'    Dim col As New Collection
'    For Each fl In fso.GetFolder(folderPath).files
'        If pattern = "" Then
'            col.Add fl.path
'        Else
'            If RegMatch(fl.Name, pattern, True) Then
'                col.Add fl.path
'            End If
'        End If
'
'    Next
'    Set GetFiles = col
'End Function

Function TrimRight(ByVal content As String, Optional ByVal num As Integer = 1) As String
    If content = "" Then Exit Function
    TrimRight = Left(content, Len(content) - num)
End Function

Function TrimLeft(ByVal content As String, Optional ByVal num As Integer = 1) As String
    If content = "" Then Exit Function
    TrimLeft = Right(content, Len(content) - num)
End Function

Function Prefix(ByVal Str As String, Optional ByVal Delimiter As String = ".") As String
    Prefix = Str
    On Error Resume Next
    Prefix = Left(Str, InStr(1, Str, Delimiter) - 1)
End Function



Function suffix(ByVal Str As String, Optional ByVal Delimiter As String = ".") As String
    suffix = Str
    On Error Resume Next
    'suffix = Right(str, Len(str) - InStrRev(str, Delimiter) - (Len(Delimiter) - 1))
    Dim a
    a = Split(Str, Delimiter)
    suffix = a(UBound(a))
End Function


Function back(ByVal Str As String) As String
    If Len(Str) <= 3 Then
        back = Str
        Exit Function
    End If
    back = Left(Str, InStrRev(Str, "\") - 1)
    
End Function


Function Inc(ByRef Target, Optional ByVal content = 1)
    
    Target = Target + content
    
    
    Inc = Target
End Function


Function RegMatch(ByVal Target As String, ByVal pattern As String, Optional ByVal ignorecase As Boolean = False) As Boolean
    Dim reg As New RegExp
    reg.Global = False
    reg.pattern = pattern
    reg.MultiLine = False
    reg.ignorecase = ignorecase
    On Error Resume Next
    RegMatch = reg.Test(Target)
    If Err.Number <> 0 Then
        RegMatch = False
    End If
End Function

Function FileExists(ByVal filepath As String)
    FileExists = fso.FileExists(filepath)
End Function

Function FolderExists(ByVal folderPath As String)
    FolderExists = fso.FolderExists(folderPath)
End Function


'Sub HyperText(ByVal Str As String, Optional ByVal param As String = vbNullString, Optional ByVal directory As Boolean = False, Optional ByVal selectfile As Boolean = False)
'    If Str <> "" Then
'        If selectfile Then
'            ShellExecute 0&, vbNullString, "explorer.exe", "/select," & Str, vbNullString, vbNormalFocus
'        Else
'
'            If directory Then
'                ShellExecute 0&, vbNullString, Str, param, left(Str, InStrRev(Str, "\")), vbNormalFocus
'            Else
'                ShellExecute 0&, vbNullString, Str, param, vbNullString, vbNormalFocus
'            End If
'        End If
'
'End If
'End Sub
