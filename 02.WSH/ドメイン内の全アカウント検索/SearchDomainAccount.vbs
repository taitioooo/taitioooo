Option Explicit
Const SearchUser = "taitioooo"
Dim baseDN, objRootDSE
Dim objConnection, objCommand, strCommandText
Dim objRecordSet, strUserDN
Dim objRec
baseDN = ""

' �x�[�XDN�̎擾
On Error Resume Next
Set objRootDSE = GetObject("LDAP://dc=kure,dc=local")
If Err.Number <> 0 Then
  WScript.Echo "�h���C���ڑ��Ɏ��s���܂����B�I�����܂��B"
  WScript.Quit
Else
  baseDN = objRootDSE.Get("defaultNamingContext")
End If
On Error Goto 0

' DC�ɐڑ����Č���

Set objCOnnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objCOnnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = 2
objCommand.Properties("Sort On") = "Name"

'SQL��WHERE��̏����ɂ�
objCommand.CommandText = _ 
  "SELECT distinguishedName FROM 'LDAP://dc=kure,dc=local' WHERE sAMAccountName='" & SearchUser & "'"
Set objRecordSet = objCommand.Execute


' �������ʂ�\��
If objRecordSet.EOF Then
  strUserDN = "���O�I���A�J�E���g " & SearchUser & " �͌�����܂���ł����B"
Else
  strUserDN = objRecordSet.Fields("distinguishedName")
End If
WScript.Echo strUserDN
objConnection.Close
Set objCommand = Nothing