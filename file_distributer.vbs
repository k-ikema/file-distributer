Option Explicit

Dim FileSys
Dim Folder, File, OpenFile, TextStream, Text
Dim ScriptPath, ListPath, FilePath
Dim FileName, RenameFileName, MatchStr, ReplaceStr
Dim MoveFrom, MoveTo
Dim Args, Mode
Dim LogMessage
Dim c

Const ERROR=0
Const MOVE="m"
Const RENAME="r"
Const RENAMELOG="file_rename.log"
Const MOVELOG="file_move.log"

Set Args =  Wscript.Arguments

If checkArguments(Args) = ERROR Then
    Wscript.quit
End If

Mode = Lcase(Args(0))

Set FileSys = CreateObject("Scripting.FileSystemObject")

ScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
FilePath   = Args(2) 'リストファイルのパス

set Folder = FileSys.getFolder(FilePath)

    
for each File In Folder.Files

    FileName = File.Name
    Set TextStream = FileSys.OpenTextFile(ListPath, 1)
    
    Do Until TextStream.AtEndOfStream = True

        Dim ReplaceList
        Text = TextStream.ReadLine

        If Left(Trim(Text),1)<>"*" and Text<>"" Then
        
            ReplaceList = split(Text, ",")
            
            ' カラム内容の | をカンマに戻す
            ReplaceList(0) = Replace(ReplaceList(0),"|",",")
            ReplaceList(1) = Replace(ReplaceList(1),"|",",")
            
            If InStr(File.Name, ReplaceList(0)) <> 0 Then 
                
                Select Case Mode
    
                    Case MOVE 
                        MoveFrom = FileSys.BuildPath( FilePath ,File.Name)
                        MoveTo   = FileSys.BuildPath( FilePath ,ReplaceList(1) & "\" & File.Name )
                        LogMessage = MoveFrom & " ---> " & MoveTo
                        AddLog MOVELOG, LogMessage
                        FileSys.MoveFile MoveFrom , MoveTo
                        Exit Do
    
                    Case RENAME
                        MatchStr       = ReplaceList(0)
                        ReplaceStr     = ReplaceList(1)
                        RenameFileName = Replace(File.Name, MatchStr, ReplaceStr )
                        LogMessage = File.Name & " ---> " & RenameFileName
                        AddLog RENAMELOG, LogMessage
                        File.Name = RenameFileName
    
                End Select
    
            End If
          End If
      Loop

    TextStream.Close

next

Set TextStream = Nothing
Set FileSys = Nothing

Wscript.echo "Finish."


' 引数チェック
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

    '引数の個数チェック
    If Args.count < 3 Then 
        chkStatus = ERROR
    End If
    
    If chkStatus = ERROR Then 
        Wscript.echo ErrMsg
        checkArguments = ERROR
        Exit Function
    End If
    
    ErrMsg = ErrMsg & vbcrlf & "-- Arguments Check --" & vbcrlf
        
    '第1引数チェック
    If Not Lcase(Args(0)) = "m" and not Lcase(Args(0)) = "r" Then
        ErrMsg = ErrMsg & "mode : m(ove) / r(ename) " & vbcrlf
        chkStatus = ERROR
    End If
    
    ListPath=Args(1)
    
    '第2引数がディレクトリだった場合にデフォルトファイル名付与
    If Right(ListPath,1)<>"\" and Right(ListPath,4)<>".txt" Then
        ListPath = ListPath & DefaultListName
    End If
    
    '規則リストの存在チェック
    If chkFS.FileExists(ListPath)=False Then
        ErrMsg = ErrMsg & "List File Path is Invalid. : " & ListPath & vbcrlf
        chkStatus = ERROR
    End If

    '対象ファイル保存ディレクトリの存在チェック
    If chkFS.FolderExists(Args(2))=False Then
        ErrMsg = ErrMsg & "Working Directory is not found. : " & Args(2) & vbcrlf
        chkStatus = ERROR
    End If
    
    set chkFS = NOTHING
    
    If chkStatus = ERROR Then Wscript.echo ErrMsg
    
    checkArguments = chkStatus
    
end Function

Sub AddLog(LogFile, LogMessage)
    Const APPEND=8 
    Dim log
    Dim FileSys

    Set FileSys =  CreateObject("Scripting.FileSystemObject")
    
    Set log = FileSys.OpenTextFile(LogFile, APPEND, true)
    LogMessage = Date() & "-" & Time() & "-" & LogMessage
    log.WriteLine (LogMessage)
    
    Set log = NOTHING
    Set FileSys = NOTHING
    
End Sub