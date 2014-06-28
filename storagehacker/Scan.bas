Attribute VB_Name = "Scan"
Public Sub ShowFolderList(folderspec)
     Dim fs, f, f1, s, sf
     Dim hs, h, h1, hf
     Set fs = CreateObject("Scripting.FileSystemObject")
     Set f = fs.GetFolder(folderspec)
     Set sf = f.SubFolders
     For Each f1 In sf
        
     List1.AddItem folderspec & "\" & f1.Name
    
           Call ShowFolderList(folderspec & "\" & f1.Name)
     Next
End Sub


'遍历某文件夹下的文件
Private Sub Showfilelist(folderspec)
     Dim fs, f, f1, fc, s
     Set fs = CreateObject("Scripting.FileSystemObject")
     Set f = fs.GetFolder(folderspec)
     Set fc = f.Files
     For Each f1 In fc
     List1.AddItem f1.Name
     Next
End Sub


'遍历某文件夹及子文件夹中的所有文件
Sub SosuoFile(MyPath As String, TargetPath As String)
On Error GoTo err
Dim Myname As String
Dim a As String
Dim B() As String
Dim dir_i() As String
Dim i, idir As Long
If Right(MyPath, 1) <> "\" Then MyPath = MyPath + "\"
Myname = Dir(MyPath, vbDirectory Or vbHidden Or vbNormal Or vbReadOnly)
Do While Myname <> ""
If Myname <> "." And Myname <> ".." Then
If (GetAttr(MyPath & Myname) And vbDirectory) = vbDirectory Then '如果找到的是目录
idir = idir + 1
MkDir TargetPath & Mid(MyPath & Myname, 3)
ReDim Preserve dir_i(idir) As String
dir_i(idir - 1) = Myname
Else

FileCopy MyPath & Myname, TargetPath & Mid(MyPath, 3) & Myname  '复制
CopyLog = CopyLog & Now & Space(3) & MyPath & Myname & Space(3) & "复制成功" & vbCrLf
conti:
End If
End If
Myname = Dir '搜索下一项
Loop
For i = 0 To idir - 1
Call SosuoFile(MyPath + dir_i(i), Form1.TargetPath)
Next i
ReDim dir_i(0) As String
Exit Sub
err:
FileCopyEx MyPath & Myname, TargetPath & Mid(MyPath, 3) & Myname  '复制
CopyLog = CopyLog & Now & Space(3) & MyPath & Myname & Space(3) & "文件正被使用 试图使用LOF复制成功" & vbCrLf
GoTo conti
End Sub

Public Function FileCopyEx(ByVal SouFileName As String, ByVal DestFileName As String)
     Dim tmpArr() As Byte
     Open SouFileName For Binary Access Read As #1
         ReDim tmpArr(LOF(1))
         Get 1, , tmpArr
     Close #1
     Open DestFileName For Binary As #2
         Put 2, , tmpArr
     Close #2
     ReDim tmpArr(0)              '释放内存
End Function
