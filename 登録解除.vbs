On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin

'アドイン情報を設定 
addInName = "EXCEL楽っくん-version3.2" 
addInFileName = "EXCEL楽っくん-version3.2.xlam"

IF MsgBox(addInName & " アドインをアンインストールしますか？", vbYesNo + vbQuestion) = vbNo Then 
  WScript.Quit 
End IF

'Excel インスタンス化 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Workbooks.Add

'アドイン登録解除 
For i = 1 To objExcel.Addins.Count 
  Set objAddin = objExcel.Addins.item(i) 
  If objAddin.Name = addInFileName Then 
    objAddin.Installed = False 
  End If 
Next

'Excel 終了 
objExcel.Quit

Set objAddin = Nothing 
Set objExcel = Nothing

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'インストール先パスの作成 
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName] 
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'ファイル削除 
If objFileSys.FileExists(installPath) = True Then 
  objFileSys.DeleteFile installPath , True 
Else 
  MsgBox "アドインファイルが存在しません。", vbExclamation 
End If

Set objWshShell = Nothing 
Set objFileSys = Nothing

IF Err.Number = 0 THEN 
   MsgBox "アドインのアンインストールが終了しました。", vbInformation 
ELSE 
   MsgBox "エラーが発生しました。" & vbCrLF & "実行環境を確認してください。", vbExclamation 
End IF