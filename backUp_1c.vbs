Dim FSO
Dim SrcPath, DstPath 
Dim CurYear, CurMnth 
Dim SrcFile

SrcPath = "D:\TEMP\BACKuP"
DstPath = "D:\TEMP\BACKuP"
SrcFile = "1Cv8.dt"

Set FSO = CreateObject("Scripting.FileSystemObject")

' проверка структуры BackUp
CheckAndCreateFolderStructure(DstPath)

' проверка файла для backUp'a
If FSO.FileExist(SrcPath &"\"&SrcFile )

    

else    
    wscript.echo "baclup file "& SrcPath &"\"&SrcFile &" not found!" 
End If

Sub CheckAndCreateFolderStructure(prmSrc)
    If FSO.FolderExists(prmSrc) then
        'wscript.echo "OK"
        CheckAndCreateFolder(prmSrc+"\1") 
        CheckAndCreateFolder(prmSrc+"\2") 
        CheckAndCreateFolder(prmSrc+"\3") 
        CheckAndCreateFolder(prmSrc+"\4") 
        CheckAndCreateFolder(prmSrc+"\5") 
        CheckAndCreateFolder(prmSrc+"\6") 
        CheckAndCreateFolder(prmSrc+"\7") 
        CheckAndCreateFolder(prmSrc+"\YEAR") 

        CurYear = Year(Now)
        'wscript.echo "current Year: "&CurYear
        CurMnth = Right("0" & Month(Now),2)
        'wscript.echo "current Month: "&CurMnth
        CheckAndCreateFolder(prmSrc+"\YEAR\" & CurYear) 
        CheckAndCreateFolder(prmSrc+"\YEAR\" & CurYear & "\" & CurMnth) 
    Else
        wscript.echo  prmSrc & " not found"
    End if
End Sub

Sub CheckAndCreateFolder(prmFldr)
    If FSO.FolderExists(prmFldr) then
        wscript.echo prmFldr &" - OK"
    Else
        wscript.echo  prmFldr & " not found"
        FSO.CreateFolder(prmFldr)
    End if
End Sub
