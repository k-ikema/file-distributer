Option Explicit

Dim FileSys
Dim Folder, File, OpenFile, TextStream, Text
Dim ScriptPath, ListPath, FilePath
Dim FileName, RenameFileName, MatchStr, ReplaceStr
Dim MoveFrom, MoveTo
Dim Args, Mode

Const ERROR=0
Const MOVE="m"
Const RENAME="r"

Set Args =  Wscript.Arguments

If checkArguments(Args) = ERROR Then
    Wscript.quit
End If

Mode = Lcase(Args(0))

Set FileSys = CreateObject("Scripting.FileSystemObject")

ScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
FilePath   = Args(2) '�Ώۃt�@�C���̃t�H���_�p�X

set Folder = FileSys.getFolder(FilePath)

for each File In Folder.Files

    FileName = File.Name
    Set TextStream = FileSys.OpenTextFile(ListPath, 1)

    Do Until TextStream.AtEndOfLine = True
        Dim ReplaceList
        Text = TextStream.ReadLine
        ReplaceList = split(Text, ",")
        
        If InStr(File.Name, ReplaceList(0)) <> 0 Then 
            
            Select Case Mode

                Case MOVE 
                    MoveFrom = FileSys.BuildPath( FilePath ,File.Name)
                    MoveTo   = FileSys.BuildPath( FilePath ,ReplaceList(1) & "\" & File.Name )
                    FileSys.MoveFile MoveFrom , MoveTo
                    Exit Do

                Case RENAME
                    MatchStr       = ReplaceList(0)
                    ReplaceStr     = ReplaceList(1)
                    RenameFileName = Replace(File.Name, MatchStr, ReplaceStr )
                    File.Name = RenameFileName

            End Select

        End If
    Loop

    TextStream.Close

next

Set TextStream = Nothing
Set FileSys = Nothing

Wscript.echo "Finish."


' �����`�F�b�N
Function checkArguments(Args)
    Dim chkFS
    Dim ErrMsg
    Dim chkStatus
    Dim DefaultListName
    chkStatus = 1
    DefaultListName = Year(Now) & "Q" & Fix((Month(Now)-1)/3)+1 & ".txt"
    set chkFS = CreateObject("Scripting.FileSystemObject")

    ErrMsg = ErrMsg & "Usage : " & vbcrlf & _
    "  file_distributer.vbs mode file_list work_dir" &  vbcrlf & vbcrlf & _ 
    "mode : m(ove) / r(ename) " &  vbcrlf  & _ 
    "list_file : full file path of rename (or move to) rule list" &  vbcrlf  & _ 
    "work_dir :  full path of directory of files to be renamed or moved" &  vbcrlf 

    '�����̐��`�F�b�N
    If Args.count < 3 Then 
        chkStatus = ERROR
    End If
    
    If chkStatus = ERROR Then 
        Wscript.echo ErrMsg
        checkArguments = ERROR
        Exit Function
    End If
    
    ErrMsg = ErrMsg & vbcrlf & "-- Arguments Check --" & vbcrlf
        
    '��1�����̃`�F�b�N
    If Not Lcase(Args(0)) = "m" and not Lcase(Args(0)) = "r" Then
        ErrMsg = ErrMsg & "mode : m(ove) / r(ename) " & vbcrlf
        chkStatus = ERROR
    End If
    
    ListPath=Args(1)
    
    '��2 �����Ƀt�@�C�����w�肳��Ă��Ȃ������ꍇ�̃f�t�H���g�t�@�C�����t�^
    If Right(ListPath,1)<>"\" and Right(ListPath,4)<>".txt" Then
        ListPath = ListPath & DefaultListName
    End If
    
    '�U�蕪���K�����X�g�t�@�C���̑��݃`�F�b�N
    If chkFS.FileExists(ListPath)=False Then
        ErrMsg = ErrMsg & "List File Path is Invalid. " & vbcrlf
        chkStatus = ERROR
    End If

    '��3 �����̃t�H���_���݃`�F�b�N
    If chkFS.FolderExists(Args(2))=False Then
        ErrMsg = ErrMsg & "Working Directory is not found. " & vbcrlf
        chkStatus = ERROR
    End If
    
    set chkFS = NOTHING
    
    If chkStatus = ERROR Then Wscript.echo ErrMsg
    
    checkArguments = chkStatus
    
end Function
