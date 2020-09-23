Attribute VB_Name = "Common"
Option Explicit

'{ -------------------------------[  NiKroWare  ]-------------------------------
'$Archive:: /Visual Basic/NKW/NKWCreateMDB/Common.bas                          $
'$Author:: Enik                                                                $
'$Date:: 10-08-01 11:22                                                        $
'$Modtime:: 11-06-01 9:31                                                      $
'$Revision:: 4                                                                 $
'-------------------------------------------------------------------------------
'Purpose  : To find password in an Access database...
'-------------------------------------------------------------------------------}


'Where:
'
'
' Q3) How do I create a Autonumber column in a new JET database?
' The secret is to set the column's ParentCatalog property before appending
' the new column to the table.
' For example, the following VB code creates a new JET database (in 3.5 format),
' then creates a new table with an autonumber column.

' Create a new autonumber ID Column
'Set col = New ADOX.Column
'col.name = "WebSiteID"
'col.Type = adInteger
'Public' Must set before setting properties and append column!
'col.Properties("Autoincrement") = True
'cat.Tables("WebSite").Columns.Append col

'Q4) Does anyone know out there how to compact an Access database through VB
'    code?  I am using ADO to connect to it.
' Set a reference to the Jet Replication Objects. Then use code as follows

' JetEngine.CompactDatabase strSource, strDestination
'  or
' Dim obj As JRO.JetEngine
' Set obj = New JRO.JetEngine
' obj.CompactDatabase strSource, strDestination

Public Function GetDBPassword(ByVal FileName As String) As String
On Error GoTo errHandler
Dim REG As cRegistry
Dim HKey As HKEYS
Dim Section As String
Dim fPWD As frmPassword
  
  Set REG = New cRegistry
  
  If REG.IsWinNT Then
    HKey = HKEY_CURRENT_USER
  Else
    HKey = HKEY_LOCAL_MACHINE
  End If
  
  Set fPWD = New frmPassword
   
  Section = "SOFTWARE\NiKroWare\" & App.Title & "\Passwords"
  If Not REG.CheckRegistryKey(HKey, Section) Then REG.CreateRegistryKey HKey, Section
  
  fPWD.Password = REG.GetRegistryValue(HKey, Section, FileName)
  fPWD.SavePassword = REG.GetRegistryValue(HKey, Section, "Save", True)
  
  fPWD.Show vbModal
  
  If fPWD.OKPressed Then
    GetDBPassword = fPWD.Password
   
    If fPWD.SavePassword Then
      REG.SetRegistryValue HKey, Section, FileName, fPWD.Password
    Else
      REG.DeleteRegistryValue HKey, Section, FileName
    End If
    
    REG.SetRegistryValue HKey, Section, "Save", fPWD.SavePassword
    
  Else
    GetDBPassword = ""
  End If
  Set REG = Nothing
  Set fPWD = Nothing
  
Exit Function
errHandler:
  MsgBox "ERROR occcured:" & vbCrLf & Err.Number & ":  " & Err.Description, vbCritical, "ERROR"
  Exit Function
  Resume
End Function

Public Function GetAccess97Password(ByVal FileName As String) As String
On Error GoTo errHandler
Dim ch(18) As Byte
Dim x As Integer
Dim Sec

  GetAccess97Password = ""

  If Trim(FileName) = "" Then Exit Function
  
' Used integers instead of hex :-)  Easier to read
  Sec = Array(0, 134, 251, 236, 55, 93, 68, 156, 250, 198, 94, 40, 230, 19, 182, 138, 96, 84)
  
  Open FileName For Binary Access Read As #1 Len = 18
  Get #1, &H42, ch
  Close #1
  
  For x = 1 To 17
    GetAccess97Password = GetAccess97Password & Chr$(ch(x) Xor Sec(x))
  Next x
  GetAccess97Password = Replace(GetAccess97Password, Chr$(0), "")
Exit Function
errHandler:
  MsgBox "ERROR occcured:" & vbCrLf & Err.Number & ":  " & Err.Description, vbCritical, "ERROR"
  Exit Function
  Resume
End Function


Public Function GetFileName(ByVal FileName As String) As String

  GetFileName = Right$(FileName, InStr(1, StrReverse(FileName), "\") - 1)

End Function


Public Function cIndexNulls(ByVal Value As ADOX.AllowNullsEnum) As String
  Select Case Value
    Case adIndexNullsAllow: cIndexNulls = "adIndexNullsAllow"
    Case adIndexNullsDisallow: cIndexNulls = "adIndexNullsDisallow"
    Case adIndexNullsIgnore: cIndexNulls = "adIndexNullsIgnore"
    Case adIndexNullsIgnoreAny: cIndexNulls = "adIndexNullsIgnoreAny"
    Case Else: cIndexNulls = Value
  End Select
End Function

Public Function cKeyType(ByVal Value As ADOX.KeyTypeEnum) As String
  Select Case Value
    Case adKeyForeign: cKeyType = "adKeyForeign"
    Case adKeyPrimary: cKeyType = "adKeyPrimary"
    Case adKeyUnique: cKeyType = "adKeyUnique"
    Case Else: cKeyType = Value
  End Select
End Function

Public Function cUpdateRule(ByVal Value As ADOX.RuleEnum) As String
  Select Case Value
    Case adRINone: cUpdateRule = "adRINone"
    Case adRICascade: cUpdateRule = "adRICascade"
    Case adRISetNull: cUpdateRule = "adRISetNull"
    Case adRISetDefault: cUpdateRule = "adRISetDefault"
    Case Else: cUpdateRule = Value
  End Select
End Function

Public Function cColumnAttributes(ByVal Value As ADOX.ColumnAttributesEnum) As String
  Select Case Value
    Case adColFixed: cColumnAttributes = "adColFixed"
    Case adColNullable: cColumnAttributes = "adColNullable"
    Case adColFixed Or adColNullable: cColumnAttributes = "adColFixed or adColNullable"
    Case Else: cColumnAttributes = Value
  End Select
End Function
