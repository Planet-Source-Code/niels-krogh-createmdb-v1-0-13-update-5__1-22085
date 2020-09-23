Attribute VB_Name = "CreateBAS"
'{ -------------------------------[  NiKroWare  ]-------------------------------
'$Archive:: /Visual Basic/NKW/NKWCreateMDB/CreateBAS.bas                       $
'$Author:: Enik                                                                $
'$Date:: 10-08-01 11:22                                                        $
'$Modtime:: 8-08-01 12:55                                                      $
'$Revision:: 4                                                                 $
'-------------------------------------------------------------------------------

Option Explicit

Public mCon As ADODB.Connection
Public mCat As ADOX.Catalog

Public DB_Name As String
Public DB_Title As String

' Engine Type = 4 creates an Access database in 3.5 format
' Engine Type = 5 creates an Access database in 4.0 format  (default)

' Note, Access 97 will not be able to open up an Access 2000 database.
' However, Access 2000 will be able to open up an Access 97 or 2000 database.
' If 97 database, then Access 2000 will ask you if you want to convert to a 2000
' database format, or just  open it read-only.

' oCat.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                     "Data Source=c:\temp\new35.mdb;" & _
'                     "Jet OLEDB:Engine Type=5;"


Public Enum EngineTypeEnum
  adAccess35 = 4
  adAccess40 = 5
End Enum

Public Function SplitString(ByVal Str As String) As String
' vb has a 1024 char per line limit and it will not split lines nicely

Dim x As Integer, y As Integer
Const SP As Integer = 80 ' split pos

Dim sTemp() As String

  If Len(Str) > SP Then
    y = (Len(Str) \ SP)
    ReDim sTemp(y)
     
    For x = 0 To y
      sTemp(x) = Mid$(Str, (x * SP) + 1, SP)
    Next x
    
    For x = LBound(sTemp) To UBound(sTemp)
      SplitString = SplitString & sTemp(x)
      If x < UBound(sTemp) Then SplitString = SplitString & """ & _" & vbCrLf & vbTab & """"
    Next x

    Erase sTemp
      
  Else
    SplitString = Str
  End If


End Function

Public Sub CreateModule(ByVal FileName As String, Optional ByVal EngineType As EngineTypeEnum = adAccess35)
Dim fHandle As Integer
On Error GoTo ErrTrap
  
  fHandle = FreeFile
  
  Open FileName For Output As #fHandle
  
  WriteHeader fHandle
  
  WriteDB fHandle, EngineType
  CreateTables fHandle
  CreateViews fHandle
  CreateProcedures fHandle
  CreateIndexes fHandle
  CreateKeys fHandle
  
  Close fHandle

Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in CreateModule"
  Close fHandle
  Exit Sub
  Resume
End Sub

Private Sub WriteHeader(ByVal fHandle As Integer)
  
  Print #fHandle, "Attribute VB_Name = ""Create" & Replace(DB_Title, ".mdb", "") & """"
  Print #fHandle, "Option Explicit"
  Print #fHandle, ""
  Print #fHandle, "' ========================================================"
  Print #fHandle, "' === Generator       : CreateMDB v" & App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
  Print #fHandle, "' === Copyright©      : 2000-2002 NiKroWare"
  Print #fHandle, "' === Created         : " & Now
  Print #fHandle, "' === Access Database : " & DB_Title
  Print #fHandle, "' ========================================================"
  Print #fHandle, ""
  Print #fHandle, "Private CAT As ADOX.Catalog"
  Print #fHandle, ""

End Sub

Private Sub WriteDB(ByVal fHandle As Integer, ByVal EngineType As EngineTypeEnum)
  
  Print #fHandle, "Public Sub CreateMDB(Byval Path As String)"
  Print #fHandle, "On Error GoTo ErrTrap"
  Print #fHandle, ""
  Print #fHandle, "  Set CAT = New ADOX.Catalog"
  
  Print #fHandle, ""
  Print #fHandle, "  If Right$(Path, 1) = ""\"" Then Path = Left$(Path, Len(Path) - 1)"
  Print #fHandle, ""
  Print #fHandle, "' ===[Create Database]==="
  Print #fHandle, "  CAT.Create ""Provider=Microsoft.Jet.OLEDB.4.0;"" & _"
  Print #fHandle, "             ""Data Source="" & Path & ""\" & DB_Title & ";"" & _"
  Print #fHandle, "             ""Jet OLEDB:Database Password=" & mCon.Properties("Jet OLEDB:Database Password") & ";"" & _"
  Print #fHandle, "             ""Jet OLEDB:Engine Type=" & EngineType & ";"""
  Print #fHandle, ""
  Print #fHandle, "  CreateTables"
  Print #fHandle, "  CreateViews"
  Print #fHandle, "  CreateProcedures"
  Print #fHandle, "  CreateIndexes"
  Print #fHandle, "  CreateKeys"
  Print #fHandle, ""
  Print #fHandle, "  Set CAT = Nothing"
  Print #fHandle, ""
  Print #fHandle, "MsgBox ""Database created."", vbApplicationModal + vbInformation, App.Title"
  Print #fHandle, "Exit Sub"
  Print #fHandle, "ErrTrap:"
  Print #fHandle, "  MsgBox Err.Number & "" / "" & Err.Description"
  Print #fHandle, "  Exit Sub"
  Print #fHandle, "  Resume"
  Print #fHandle, "End Sub"
  Print #fHandle, ""
  
End Sub

' === [ TABLE Routines ] ===
Private Sub CreateTables(ByVal fHandle As Integer)
On Error GoTo ErrTrap
Dim F As Integer
    
  Print #fHandle, "Private Sub CreateTables()"
  Print #fHandle, "On Error GoTo ErrTrap"
  Print #fHandle, "Dim TBL As ADOX.Table"
  
  Print #fHandle, ""
  
  For F = 0 To mCat.Tables.Count - 1
    
    If mCat.Tables(F).Type = "TABLE" Or mCat.Tables(F).Type = "LINK" Then
      WriteTable fHandle, mCat.Tables(F)
      Print #fHandle, ""
    End If
  
  Next F
  
  Print #fHandle, "  Set TBL = Nothing"
  Print #fHandle, ""
  Print #fHandle, "Exit Sub"
  Print #fHandle, "ErrTrap:"
  Print #fHandle, "  MsgBox Err.Number & "" / "" & Err.Description,,""Error In CreateTables"""
  Print #fHandle, "  Exit Sub"
  Print #fHandle, "  Resume"
  Print #fHandle, "End Sub"
  Print #fHandle, ""
  
Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in CreateTables"
  Exit Sub
  Resume
End Sub

Private Sub WriteTable(ByVal fHandle As Integer, TBL As ADOX.Table)
On Error GoTo ErrTrap
Dim F As Integer
Dim RS As ADODB.Recordset
Dim FLD As ADODB.Field
Dim Col As ADOX.Column
  
  Set RS = New ADODB.Recordset
  RS.source = "SELECT * FROM [" & TBL.name & "] WHERE 0=1"
  RS.Open , mCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  Set RS.ActiveConnection = Nothing

  Print #fHandle, "' ===[Create Table '" & TBL.name & "']==="
  Print #fHandle, "  Set TBL = New ADOX.Table"
  Print #fHandle, "  Set TBL.ParentCatalog = CAT"
  Print #fHandle, "  TBL.Name = """ & TBL.name & """"
  
  For Each FLD In RS.Fields
    If Left$(FLD.name, 2) <> "s_" Then ' ignore the system columns...
      
      Set Col = TBL.Columns(FLD.name)

      Print #fHandle, "  TBL.Columns.Append """ & Col.name & """, " & cType(Col.Type) & ", " & Col.DefinedSize

      If Col.Properties("AutoIncrement") Then
        Print #fHandle, "  TBL.Columns(""" & Col.name & """).Properties(""AutoIncrement"") = -1 ' True"
        
        If Col.Properties("Seed") <> 1 Then
          Print #fHandle, "  TBL.Columns(""" & Col.name & """).Properties(""Seed"") = " & Col.Properties("Seed")
        End If
        
        If Col.Properties("Increment") <> 1 Then
          Print #fHandle, "  TBL.Columns(""" & Col.name & """).Properties(""Increment"") = " & Col.Properties("Increment")
        End If
      End If
      
' This is an odd one :
' The Access Column Property 'Required' is equal the ADOX property 'Nullable'
' but there must be some bug in ADOX because :
' to SET a column to be required you must set Nullable to False (make sense) but
' when to GET a column is required the Nullable must be True (make no sense)
' 04-July-2002
' Well, NullAble must be true, found by Morgan Haueisen.
 

    
      If Col.Properties("NullAble") Then
        Print #fHandle, "  TBL.Columns(""" & Col.name & """).Properties(""NullAble"") = True"
      End If
      
      If Col.Properties("Jet OLEDB:Allow Zero Length").Value Then
        Print #fHandle, "  TBL.Columns(""" & Col.name & """).Properties(""Jet OLEDB:Allow Zero Length"") = True"
      End If
      
      If Len(Col.Properties("Description")) > 0 Then
        Print #fHandle, "  TBL.Columns(""" & Col.name & """).Properties(""Description"") = """ & Replace(Col.Properties("Description").Value, """", "'") & """"
      End If
      
      If Col.Properties("Default") <> 0 Then
        If IsNumeric(Col.Properties("Default")) Then
          Print #fHandle, "  TBL.Columns(""" & Col.name & """).Properties(""Default"") = " & Col.Properties("Default")
        Else
          Print #fHandle, "  TBL.Columns(""" & Col.name & """).Properties(""Default"") = """ & Col.Properties("Default") & """"
        End If
      End If
      
    End If
  Next
  
  Print #fHandle, "  CAT.Tables.Append TBL"
  
  RS.Close
  Set RS = Nothing
  
Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in WriteTable"
  Exit Sub
  Resume
End Sub

' === [ View Routines ] ===
Private Sub CreateViews(ByVal fHandle As Integer)
On Error GoTo ErrTrap
Dim F As Integer
    
  Print #fHandle, "Private Sub CreateViews()"
  Print #fHandle, "On Error GoTo ErrTrap"
  Print #fHandle, "Dim CMD As ADODB.Command"
  Print #fHandle, ""
  
  For F = 0 To mCat.Views.Count - 1
    WriteView fHandle, mCat.Views(F)
   Print #fHandle, ""
  Next F
    
  Print #fHandle, "Exit Sub"
  Print #fHandle, "ErrTrap:"
  Print #fHandle, "  MsgBox Err.Number & "" / "" & Err.Description,,""Error In CreateViews"""
  Print #fHandle, "  Exit Sub"
  Print #fHandle, "  Resume"
  Print #fHandle, "End Sub"
  Print #fHandle, ""

Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in CreateViews"
  Exit Sub
  Resume
End Sub

Private Sub WriteView(ByVal fHandle As Integer, VIW As ADOX.View)
On Error GoTo ErrTrap
Dim CmdText As String
  
  
  CmdText = VIW.Command.CommandText
  CmdText = Replace(CmdText, vbCrLf, " ", , , vbTextCompare)  ' remove CRLF
  CmdText = Replace(CmdText, """", "'", , , vbTextCompare)    ' remove Quote
  CmdText = SplitString(CmdText)
  
  Print #fHandle, "' ===[Create View '" & VIW.name & "']==="
  Print #fHandle, "  Set CMD = New ADODB.Command"
  Print #fHandle, "  CMD.CommandText = """ & CmdText & """"
  
  Print #fHandle, "  CAT.Views.Append """ & VIW.name & """" & ",CMD"
  Print #fHandle, ""
  Print #fHandle, "  Set CMD = Nothing"

Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in WriteView"
  Exit Sub
  Resume
End Sub

' === [ Procedure Routines ] ===
Private Sub CreateProcedures(ByVal fHandle As Integer)
On Error GoTo ErrTrap
Dim F As Integer
    
  Print #fHandle, "Private Sub CreateProcedures()"
  Print #fHandle, "On Error GoTo ErrTrap"
  Print #fHandle, "Dim CMD As ADODB.Command"
  Print #fHandle, ""
  
  For F = 0 To mCat.Procedures.Count - 1
    WriteProcedure fHandle, mCat.Procedures(F)
    Print #fHandle, ""
  Next F
    
  Print #fHandle, "Exit Sub"
  Print #fHandle, "ErrTrap:"
  Print #fHandle, "  MsgBox Err.Number & "" / "" & Err.Description,,""Error In CreateProcedures"""
  Print #fHandle, "  Exit Sub"
  Print #fHandle, "  Resume"
  Print #fHandle, "End Sub"
  Print #fHandle, ""
  
Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in CreateProcedures"
  Exit Sub
  Resume
End Sub

Private Sub WriteProcedure(ByVal fHandle As Integer, PROC As ADOX.Procedure)
On Error GoTo ErrTrap
Dim CmdText As String
  
  CmdText = PROC.Command.CommandText
  CmdText = Replace(CmdText, vbCrLf, " ", , , vbTextCompare)  ' remove CRLF
  CmdText = Replace(CmdText, """", "'", , , vbTextCompare)    ' replace Quotes
  CmdText = SplitString(CmdText)
  
  Print #fHandle, "' ===[Create Procedure '" & PROC.name & "']==="
  Print #fHandle, "  Set CMD = New ADODB.Command"
  Print #fHandle, "  CMD.CommandText = """ & CmdText & """"
  
  Print #fHandle, "  CAT.Procedures.Append """ & PROC.name & """" & ",CMD"
  Print #fHandle, ""
  Print #fHandle, "  Set CMD = Nothing"

Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in WriteProcedure"
  Exit Sub
  Resume
End Sub

' === [ INDEX Routines ] ===
Private Sub CreateIndexes(ByVal fHandle As Integer)
On Error GoTo ErrTrap
Dim F As Integer, I As Integer
    
  Print #fHandle, "Private Sub CreateIndexes()"
  Print #fHandle, "On Error GoTo ErrTrap"
  Print #fHandle, "Dim IDX As ADOX.index"
  Print #fHandle, "  Set IDX = New ADOX.index"
  Print #fHandle, ""
  
  For F = 0 To mCat.Tables.Count - 1
    For I = 0 To mCat.Tables(F).Indexes.Count - 1
      If mCat.Tables(F).Type = "TABLE" Then
        If Left$(mCat.Tables(F).Indexes(I).name, 2) <> "s_" Then
          WriteIndex fHandle, mCat.Tables(F).name, mCat.Tables(F).Indexes(I)
        End If
      End If
    Next I
  Next F
  
  Print #fHandle, "  Set IDX = Nothing"
  Print #fHandle, ""
  Print #fHandle, "  Exit Sub"
  Print #fHandle, "ErrTrap:"
  Print #fHandle, "  MsgBox Err.Number & "" / "" & Err.Description,,""Error In CreateIndexes"""
  Print #fHandle, "  Exit Sub"
  Print #fHandle, "  Resume"
  Print #fHandle, "End Sub"
  Print #fHandle, ""
  
Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in CreateIndexes"
  Exit Sub
  Resume
End Sub

Private Sub WriteIndex(ByVal fHandle As Integer, ByVal TBL_Name As String, ByVal IDX As ADOX.index)
On Error GoTo ErrTrap
Dim F As Integer
  Print #fHandle, "' ===[Create Index '" & IDX.name & "']==="
  Print #fHandle, "  Set IDX = New ADOX.Index"
  Print #fHandle, "  IDX.Name = """ & IDX.name & """"
  
  For F = 0 To IDX.Columns.Count - 1
    Print #fHandle, "  IDX.Columns.Append """ & IDX.Columns(F).name & """"
  Next F
  
  Print #fHandle, "  IDX.PrimaryKey = " & IDX.PrimaryKey
  Print #fHandle, "  IDX.Unique = " & IDX.Unique
  Print #fHandle, "  IDX.Clustered = " & IDX.Clustered
  Print #fHandle, "  IDX.IndexNulls = " & cIndexNulls(IDX.IndexNulls)
  Print #fHandle, "  CAT.Tables(""" & TBL_Name & """).Indexes.Append IDX"
  Print #fHandle, ""
  
Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in WriteIndex"
  Exit Sub
  Resume
End Sub

' === [ KEY Routines ] ===
Private Sub CreateKeys(ByVal fHandle As Integer)
On Error GoTo ErrTrap
Dim F As Integer, I As Integer
    
  Print #fHandle, "Private Sub CreateKeys()"
  Print #fHandle, "On Error GoTo ErrTrap"
  Print #fHandle, "Dim KEY As ADOX.KEY"
  Print #fHandle, "Dim TBL As ADOX.Table"
  Print #fHandle, ""
  Print #fHandle, "  Set KEY = New ADOX.Key"
  Print #fHandle, "  Set TBL = New ADOX.Table"
  Print #fHandle, ""
  
  For F = 0 To mCat.Tables.Count - 1
    For I = 0 To mCat.Tables(F).Keys.Count - 1
      If Left$(mCat.Tables(F).Keys(I).name, 2) <> "s_" Then
        WriteKey fHandle, mCat.Tables(F).name, mCat.Tables(F).Keys(I)
      End If
    Next I
  Next F

  Print #fHandle, "  Set KEY = Nothing"
  Print #fHandle, "  Set TBL = Nothing"
  Print #fHandle, ""
  Print #fHandle, "  Exit Sub"
  Print #fHandle, "ErrTrap:"
  Print #fHandle, "  Select Case Err.Number"
  Print #fHandle, "  Case -2147467259  ' Index already exists - Remove it..."
  Print #fHandle, "    CAT.Tables(TBL.Name).Indexes.Delete Key.Name"
  Print #fHandle, "    Resume"
  Print #fHandle, "  Case Else"
  Print #fHandle, "    MsgBox Err.Number & "" / "" & Err.Description,,""Error In CreateKeys"""
  Print #fHandle, "    Exit Sub"
  Print #fHandle, "    Resume"
  Print #fHandle, "  End Select"
  Print #fHandle, "End Sub"
  Print #fHandle, ""

Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in CreateKeys"
  Exit Sub
  Resume
End Sub

Private Sub WriteKey(ByVal fHandle As Integer, ByVal TBL_Name As String, ByVal vKey As ADOX.KEY)
On Error GoTo ErrTrap
Dim F As Integer

  If vKey.Type = adKeyForeign Then
    Print #fHandle, "' ===[Create Key '" & vKey.name & "']==="
    
    Print #fHandle, "  Set KEY = New ADOX.Key"
    Print #fHandle, "  KEY.Name = """ & vKey.name & """"
    Print #fHandle, "  KEY.Type = " & cKeyType(vKey.Type)
    Print #fHandle, "  KEY.UpdateRule = " & cUpdateRule(vKey.UpdateRule)
    Print #fHandle, "  KEY.RelatedTable = """ & vKey.RelatedTable & """"
    
    For F = 0 To vKey.Columns.Count - 1
      Print #fHandle, "  KEY.Columns.Append """ & vKey.Columns(F).name & """"
      Print #fHandle, "  KEY.Columns(""" & vKey.Columns(F).name & """).RelatedColumn = """ & vKey.Columns(F).RelatedColumn & """"
    Next F
    Print #fHandle, "  TBL.Name = """ & TBL_Name; """"
    Print #fHandle, "  CAT.Tables(""" & TBL_Name & """).Keys.Append KEY"
    Print #fHandle, ""
  
  ElseIf vKey.Type = adKeyUnique Or vKey.Type = adKeyPrimary Then
  ' found by Rainer Leonardy / Brett Woodward
  ' ignore it - the keys create by the index routines...
    Debug.Print vKey.name & " is a " & cKeyType(vKey.Type)
  End If
   
Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in WriteKey"
  Exit Sub
  Resume
End Sub

Public Function cType(ByVal Value As ADOX.DataTypeEnum) As String
  Select Case Value
    Case adTinyInt: cType = "adTinyInt"
    Case adSmallInt: cType = "adSmallInt"
    Case adInteger: cType = "adInteger"
    Case adBigInt: cType = "adBigInt"
    Case adUnsignedTinyInt: cType = "adUnsignedTinyInt"
    Case adUnsignedSmallInt: cType = "adUnsignedSmallInt"
    Case adUnsignedInt: cType = "adUnsignedInt"
    Case adUnsignedBigInt: cType = "adUnsignedBigInt"
    Case adSingle: cType = "adSingle"
    Case adDouble: cType = "adDouble"
    Case adCurrency: cType = "adCurrency"
    Case adDecimal: cType = "adDecimal"
    Case adNumeric: cType = "adNumeric"
    Case adBoolean: cType = "adBoolean"
    Case adUserDefined: cType = "adUserDefined"
    Case adVariant: cType = "adVariant"
    Case adGUID: cType = "adGUID"
    Case adDate: cType = "adDate"
    Case adDBDate: cType = "adDBDate"
    Case adDBTime: cType = "adDBTime"
    Case adDBTimeStamp: cType = "adDBTimeStamp"
    Case adBSTR: cType = "adBSTR"
    Case adChar: cType = "adChar"
    Case adVarChar: cType = "adVarChar"
    Case adLongVarChar: cType = "adLongVarChar"
    Case adWChar: cType = "adWChar"
    Case adVarWChar: cType = "adVarWChar"
    Case adLongVarWChar: cType = "adLongVarWChar"
    Case adBinary: cType = "adBinary"
    Case adVarBinary: cType = "adVarBinary"
    Case adLongVarBinary: cType = "adLongVarBinary"
    Case Else: cType = Value
  End Select
End Function

Private Function cIndexNulls(ByVal Value As ADOX.AllowNullsEnum) As String
  Select Case Value
    Case adIndexNullsAllow: cIndexNulls = "adIndexNullsAllow"
    Case adIndexNullsDisallow: cIndexNulls = "adIndexNullsDisallow"
    Case adIndexNullsIgnore: cIndexNulls = "adIndexNullsIgnore"
    Case adIndexNullsIgnoreAny: cIndexNulls = "adIndexNullsIgnoreAny"
    Case Else: cIndexNulls = Value
  End Select
End Function

Private Function cKeyType(ByVal Value As ADOX.KeyTypeEnum) As String
  Select Case Value
    Case adKeyForeign: cKeyType = "adKeyForeign"
    Case adKeyPrimary: cKeyType = "adKeyPrimary"
    Case adKeyUnique: cKeyType = "adKeyUnique"
    Case Else: cKeyType = Value
  End Select
End Function

Private Function cUpdateRule(ByVal Value As ADOX.RuleEnum) As String
  Select Case Value
    Case adRINone: cUpdateRule = "adRINone"
    Case adRICascade: cUpdateRule = "adRICascade"
    Case adRISetNull: cUpdateRule = "adRISetNull"
    Case adRISetDefault: cUpdateRule = "adRISetDefault"
    Case Else: cUpdateRule = Value
  End Select
End Function


