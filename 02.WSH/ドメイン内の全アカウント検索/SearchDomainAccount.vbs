On Error Resume Next 
 
Const ADS_SCOPE_SUBTREE = 2 
 
Set objConnection = CreateObject("ADODB.Connection") 
Set objCommand =   CreateObject("ADODB.Command") 
objConnection.Provider = "ADsDSOObject" 
objConnection.Open "Active Directory Provider" 
Set objCommand.ActiveConnection = objConnection 
 
objCommand.Properties("Page Size") = 1000 
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE  
objCommandProperties("Sort On") = "Name" 
 
objCommand.CommandText = _ 
    "SELECT Name FROM 'LDAP://dc=fabrikam,dc=com' WHERE objectCategory='user'"   
Set objRecordSet = objCommand.Execute 
 
objRecordSet.MoveFirst 
Do Until objRecordSet.EOF 
    Wscript.Echo objRecordSet.Fields("Name").Value 
    objRecordSet.MoveNext 
Loop