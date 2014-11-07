Option Explicit
Const SearchUser = "taitioooo"
Dim baseDN, objRootDSE
Dim objConnection, objCommand, strCommandText
Dim objRecordSet, strUserDN
Dim objRec
baseDN = ""

' ベースDNの取得
On Error Resume Next
Set objRootDSE = GetObject("LDAP://dc=kure,dc=local")
If Err.Number <> 0 Then
  WScript.Echo "ドメイン接続に失敗しました。終了します。"
  WScript.Quit
Else
  baseDN = objRootDSE.Get("defaultNamingContext")
End If
On Error Goto 0

' DCに接続して検索

Set objCOnnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objCOnnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = 2
objCommand.Properties("Sort On") = "Name"

'SQLのWHERE句の条件には
objCommand.CommandText = _ 
  "SELECT distinguishedName FROM 'LDAP://dc=kure,dc=local' WHERE sAMAccountName='" & SearchUser & "'"
Set objRecordSet = objCommand.Execute


' 検索結果を表示
If objRecordSet.EOF Then
  strUserDN = "ログオンアカウント " & SearchUser & " は見つかりませんでした。"
Else
  strUserDN = objRecordSet.Fields("distinguishedName")
End If
WScript.Echo strUserDN
objConnection.Close
Set objCommand = Nothing