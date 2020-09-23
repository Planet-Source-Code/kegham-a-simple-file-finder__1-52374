Attribute VB_Name = "Module1"
Function FilesSearch(thepath As String, Ext As String)

    Dim unkndir() As String
    Dim tempdir As String
    Dim ff As String
    Dim DirCount As Integer
    Dim X As Integer
    DirCount = 0
    ReDim unkndir(0) As String
    unkndir(DirCount) = ""
    If Right(thepath, 1) <> "\" Then
    thepath = thepath & "\"
    End If
    DoEvents
    tempdir = Dir(thepath, vbDirectory)
    Do While tempdir <> ""
    If tempdir <> "." And tempdir <> ".." Then
    If (GetAttr(thepath & tempdir) And vbDirectory) = vbDirectory Then
    unkndir(DirCount) = thepath & tempdir & "\"
    DirCount = DirCount + 1
    ReDim Preserve unkndir(DirCount) As String
    End If
    End If
    tempdir = Dir
    Loop
    ff = Dir(thepath & Ext)
    Do Until ff = ""
    
    'In case file found
     Form1.List1.AddItem thepath & ff
     Form1.Label2.Caption = " File found in  " & thepath
     ff = Dir
     Loop

     'searches through all sub direictories

     For X = 0 To (UBound(unkndir) - 1)
     FilesSearch unkndir(X), Ext
     Next X
     End Function
