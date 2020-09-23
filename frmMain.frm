VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3675
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   1695
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmMain.frx":0000
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   1665
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   2937
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgList"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   120
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Files"
      Begin VB.Menu mnuFileOpenDB 
         Caption         =   "&Open Database"
      End
      Begin VB.Menu mnuFileAnalyzeDB 
         Caption         =   "&Analyze Database"
      End
      Begin VB.Menu mnuFileCompressDB 
         Caption         =   "&Compress"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export"
         Begin VB.Menu mnuFileExportBAS 
            Caption         =   "BAS-module (Access 2000)"
            Index           =   0
         End
         Begin VB.Menu mnuFileExportBAS 
            Caption         =   "BAS-module (Access 97)"
            Index           =   1
         End
         Begin VB.Menu mnuFileExportLine1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileExportSQL 
            Caption         =   "SQL"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuMRUFiles 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'{ -------------------------------[  NiKroWare  ]-------------------------------
'$Archive:: /Visual Basic/NKW/NKWCreateMDB/frmMain.frm                         $
'$Author:: Enik                                                                $
'$Date:: 4-10-01 10:36                                                         $
'$Modtime:: 4-10-01 10:33                                                      $
'$Revision:: 6                                                                 $
'-------------------------------------------------------------------------------
'Purpose  : To generate a BAS module to be included into a VB project...
'-------------------------------------------------------------------------------}

' For use when we are dragging the splitter.
Private Const SPLITTER_WIDTH = 60

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private mMRU As cMRU

Private mJetPassword As String

  Private Percentage1 As Single
  Private mbDragging As Boolean

Private Sub Form_Load()
On Error GoTo ErrTrap

  Me.Icon = LoadResPicture("1ICON", vbResIcon)
  Me.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
  
  mnuFileAnalyzeDB.Enabled = 0   ' False
  mnuFileExport.Enabled = 0 ' False
   
  Me.Width = Screen.Width * 0.7
  Me.Height = Screen.Height * 0.7
  
  Percentage1 = 0.35
  mbDragging = 0 ' False
  
  ArrangeControls
  
  Load_ImgList
  TV_Setup
  LV_Setup
  SB_Setup
  
  Menu_Setup

Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in Form_Load"
  Exit Sub
  Resume
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   mbDragging = 1 ' True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Not mbDragging Then
    If x > TV.Width And x < LV.Left Then Me.MousePointer = vbSizeWE
    Exit Sub
  End If

  Percentage1 = x / Me.ScaleWidth   ' VSPLIT

  If Percentage1 < 0 Then Percentage1 = 0
  If Percentage1 > 1 Then Percentage1 = 1
  ArrangeControls
  
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  mbDragging = 0 ' False
  Me.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
  
  ArrangeControls

End Sub

Private Sub LV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   
  Me.MousePointer = vbDefault

End Sub

Private Sub mnuFileCompressDB_Click()
' Dim JRO2 As jro.JetEngine
' Set JRO2 = New jro.JetEngine
' JRO2.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\\nwind2.mdb", _
'                      "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\\abbc2.mdb;Jet OLEDB:Engine Type=4"

  MsgBox "Not implemented jet..."
  
End Sub

Private Sub mnuFileExportBAS_Click(index As Integer)
Dim EngineType As EngineTypeEnum
Dim DLG As clsOpenSave
On Error GoTo ErrTrap
  
  Set DLG = New clsOpenSave
  
  DLG.CancelError = 1 ' True
  DLG.flags = OFN_HIDEREADONLY + OFN_OVERWRITEPROMPT
  DLG.Filter = "All Files (*.*)|*.*|Visual Basic Module Files (*.bas)|*.bas"
  DLG.FilterIndex = 2
  DLG.hwnd = Me.hwnd
  DLG.DialogTitle = "Save BAS module"
  DLG.InitDir = ""
  
  DLG.FileName = "Create_" & Replace(DB_Title, ".mdb", "", , , vbTextCompare) & ".bas"
  DLG.ShowSave
  
  If index = 0 Then
    EngineType = adAccess40
  Else
    EngineType = adAccess35
  End If
  
  CreateBAS.CreateModule DLG.FileName, EngineType

  MsgBox "BAS-module created.", vbApplicationModal + vbInformation, App.Title

ErrTrap:
' User pressed cancel...
  Set DLG = Nothing
End Sub

Private Sub mnuFileExportSQL_Click()

' Not implemeted yet...
'Dim DLG As clsOpenSave
'On Error GoTo ErrTrap
'  Set DLG = New clsOpenSave
'  DLG.CancelError = 1 ' True
'  DLG.Flags = OFN_HIDEREADONLY + OFN_OVERWRITEPROMPT
'  DLG.Filter = "All Files (*.*)|*.*|SQL Files (*.sql)|*.sql"
'  DLG.FilterIndex = 2
'  DLG.hWnd = Me.hWnd
'  DLG.DialogTitle = "Save SQL file"
'  DLG.InitDir = ""
'  DLG.FileName = "Create_" & Replace(DB_Title, ".mdb", "", , , vbTextCompare) & ".sql"
'  DLG.ShowSave
'  CreateSQL.CreateSQL DLG.FileName
'  MsgBox "SQL-file created.", vbApplicationModal + vbInformation, App.Title
'ErrTrap:
'' User pressed cancel...
'  Set DLG = Nothing
End Sub

Private Sub mnuFileOpenDB_Click()
Dim DLG As New clsOpenSave
On Error GoTo ErrTrap:

  DLG.CancelError = 1 ' True
  DLG.FileName = "*.mdb"
  DLG.flags = OFN_HIDEREADONLY + OFN_FILEMUSTEXIST
  DLG.DialogTitle = "Open Access database"
  DLG.InitDir = ""
  DLG.hwnd = Me.hwnd
      
  DLG.Filter = "All Files (*.*)|*.*|Access Database Files (*.mdb)|*.mdb"
  DLG.FilterIndex = 2
  DLG.ShowOpen
  
  DB_Name = DLG.FileName
  Set DLG = Nothing
  OpenDB DB_Name

Exit Sub
ErrTrap:
  Set DLG = Nothing
  Select Case Err.Number
  Case 32755
  ' user pressed cancel...
  Case Else
    MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in mnuFileOpenDB_Click"
    Exit Sub
    Resume
  End Select
End Sub

Private Sub OpenDB(ByVal FileName As String)
Dim LoopTimes As Byte

  mJetPassword = ""
  LoopTimes = 3
  
  On Error Resume Next
  
  Do
    Err.Clear
    LoopTimes = LoopTimes - 1
    
    If Not mCon Is Nothing Then Set mCon = Nothing
    Set mCon = New ADODB.Connection
    
    mCon.Provider = "Microsoft.Jet.OLEDB.4.0"
    mCon.Mode = adModeRead
    mCon.CursorLocation = adUseClient
    mCon.Properties("Data Source") = FileName
    mCon.Properties("Jet OLEDB:Database Password") = mJetPassword
    mCon.Open
    
    If Err.Number = 0 Then ' success let's get out of this loop...
      LoopTimes = 0
    
    ElseIf (Err.Number = -2147217843) And (LoopTimes = 2) Then ' try Access 97 Password
       mJetPassword = Common.GetAccess97Password(FileName)
    ElseIf (Err.Number = -2147217843) And (LoopTimes = 1) Then  ' try the box...
       mJetPassword = Common.GetDBPassword(FileName)
    Else
      MsgBox "Can't open DB : " & FileName
      LoopTimes = 0
    End If
  Loop While (LoopTimes > 0)

  If Not mCon Is Nothing Then
    If mCon.State = adStateOpen Then
      SB.SimpleText = "File : " & FileName
      
      DB_Name = FileName
      DB_Title = GetFileName(FileName)
      
      Set mCat = Nothing
      Set mCat = New ADOX.Catalog
      mCat.ActiveConnection = mCon
  
    ' the the db to the MRU list...
      mMRU.Add FileName
      mMRU.Update Me
      
      AnalyzeDB
  
      mnuFileAnalyzeDB.Enabled = 1 ' True
      mnuFileExport.Enabled = 1 ' True
      
    End If
  End If
  
  
  
Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in OpenDB"
  Exit Sub
  Resume
End Sub

Private Sub AnalyzeDB()
On Error GoTo ErrTrap
Dim NodX As Node
Dim TBL As ADOX.Table
Dim Col As ADOX.Column
Dim IDX As ADOX.index
Dim VIW As ADOX.View

Dim PROC As ADOX.Procedure

  Screen.MousePointer = vbHourglass
  
  LV.ListItems.Clear
  
  TV.Nodes.Clear
  Set NodX = TV.Nodes.Add(, , "DATABASE", "Database", "DATABASE")
  NodX.Tag = "DATABASE"
  NodX.ForeColor = vbBlue
  NodX.Bold = True
   
  Set NodX = TV.Nodes.Add("DATABASE", tvwChild, "TABLES", "Tables", "TABLES")
  NodX.Tag = "TABLES"
  NodX.ForeColor = vbBlue
 
  
  Set NodX = TV.Nodes.Add("DATABASE", tvwChild, "QUERIES", "Queries", "TABLES")
  NodX.Tag = "QUERIES"
  NodX.ForeColor = vbBlue
    
  For Each TBL In mCat.Tables
    If TBL.Type = "TABLE" Then
      Set NodX = TV.Nodes.Add("TABLES", tvwChild, TBL.name, TBL.name, "TABLE")
    ElseIf TBL.Type = "LINK" Then
      Set NodX = TV.Nodes.Add("TABLES", tvwChild, TBL.name, TBL.name, "TABLELINKED")
    End If
    
    If TBL.Type = "TABLE" Or TBL.Type = "LINK" Then
      DoEvents
      NodX.Tag = "TABLE"
      NodX.EnsureVisible
      Set NodX = TV.Nodes.Add(TBL.name, tvwChild, TBL.name & "\" & "COLUMNS", "", "COLUMN")
      NodX.Tag = "COLUMNS"
      
      For Each Col In TBL.Columns
        If Left$(Col.name, 2) <> "s_" Then
          Set NodX = TV.Nodes.Add(TBL.name & "\" & "COLUMNS", tvwChild, TBL.name & "\" & Col.name, Col.name, "COLUMN")
          NodX.Tag = "COLUMN"
        End If
      Next
      TV.Nodes(TBL.name & "\" & "COLUMNS").Text = "Columns (" & TV.Nodes(TBL.name & "\" & "COLUMNS").Children & ")"
      
      Set NodX = TV.Nodes.Add(TBL.name, tvwChild, TBL.name & "\" & "INDEXES", "", "COLUMN")
      NodX.Tag = "INDEXES"
      
      For Each IDX In TBL.Indexes
        If Left$(IDX.name, 2) <> "s_" Then
          Set NodX = TV.Nodes.Add(TBL.name & "\" & "INDEXES", tvwChild, "IDX:" & TBL.name & "\" & IDX.name, IDX.name, "COLUMN")
          NodX.Tag = "INDEX"
        End If
      Next
      TV.Nodes(TBL.name & "\" & "INDEXES").Text = "Indexes (" & TV.Nodes(TBL.name & "\" & "INDEXES").Children & ")"
    ElseIf TBL.Type = "VIEW" Or TBL.Type = "PROC" Then
      Debug.Print TBL.name & " is a " & TBL.Type
    End If
  Next
  
  TV.Nodes("TABLES").Bold = (TV.Nodes("TABLES").Children > 0)
  TV.Nodes("TABLES").Text = "Tables (" & TV.Nodes("TABLES").Children & ")"
  
  Screen.MousePointer = vbDefault
  'Exit Sub
  
' View / procedures / query
  For Each VIW In mCat.Views
    Set NodX = TV.Nodes.Add("QUERIES", tvwChild, VIW.name, VIW.name, "QUERY")
    NodX.Tag = "VIEW"
  Next
  
  For Each PROC In mCat.Procedures
    Set NodX = TV.Nodes.Add("QUERIES", tvwChild, PROC.name, PROC.name, "QUERY")
    NodX.Tag = "PROC"
  Next
  
   TV.Nodes("QUERIES").Bold = (TV.Nodes("QUERIES").Children > 0)
  TV.Nodes("QUERIES").Text = "Queries (" & TV.Nodes("QUERIES").Children & ")"

  Screen.MousePointer = vbDefault

Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in AnalyzeDB"
  Exit Sub
  Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  mMRU.Save
  Set mMRU = Nothing
  
  Set mCat = Nothing
  Set mCon = Nothing

End Sub

Private Sub mnuFileAnalyzeDB_Click()
  
  AnalyzeDB

End Sub

Private Sub mnuFileExit_Click()
  
  Unload Me
  End

End Sub

Private Sub Load_ImgList()
On Error GoTo ErrTrap
' Load the icons from the resourcefile into the Imagelist...

  imgList.ListImages.Clear
  imgList.ImageHeight = 16
  imgList.ImageWidth = 16
  
  imgList.ListImages.Add , "DATABASE", LoadResPicture("DATABASE", vbResIcon)
  imgList.ListImages.Add , "TABLES", LoadResPicture("TABLES", vbResIcon)
  imgList.ListImages.Add , "TABLE", LoadResPicture("TABLE", vbResIcon)
  imgList.ListImages.Add , "TABLELINKED", LoadResPicture("TABLELINKED", vbResIcon)
  imgList.ListImages.Add , "COLUMNS", LoadResPicture("COLUMNS", vbResIcon)
  imgList.ListImages.Add , "COLUMN", LoadResPicture("COLUMN", vbResIcon)
  imgList.ListImages.Add , "VARIABLE", LoadResPicture("VARIABLE", vbResIcon)
  imgList.ListImages.Add , "QUERY", LoadResPicture("QUERY", vbResIcon)
  
Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, vbExclamation, "Error in Load_ImgList"
  Exit Sub
  Resume
End Sub

Private Sub TV_Setup()
' Setup of the TreeView...

  TV.LabelEdit = tvwManual
  TV.Indentation = 256
  TV.LineStyle = tvwTreeLines
  TV.Sorted = 1 ' True
  Set TV.ImageList = imgList

End Sub

Private Sub LV_Setup()

  LV.View = lvwReport
  
  LV.LabelEdit = lvwManual
  LV.GridLines = 1 ' True
  
  LV.ColumnHeaders.Add , "VARIABLE", "Variable", LV.Width * 0.35
  LV.ColumnHeaders.Add , "VALUE", "Value", LV.Width * 0.5
  
  Set LV.SmallIcons = imgList
  
  LV.PictureAlignment = lvwTile
  LV.Picture = LoadResPicture("LVBG", vbResBitmap)


End Sub

Private Sub SB_Setup()

  SB.Style = sbrSimple
  SB.SimpleText = "No MDB loaded."

End Sub

Private Sub Menu_Setup()
Dim I As Byte
  If mMRU Is Nothing Then Set mMRU = New cMRU
  
  mnuMRUFiles(0).Visible = False
  
  mMRU.Number = 4
  mMRU.Load
  
  
  For I = 1 To mMRU.Number
    Load mnuMRUFiles(I)
    mnuMRUFiles(I).Visible = False
  Next I
  
  mMRU.Update Me
  
End Sub

Private Sub LV_LoadDATABASE() '(ByVal Node As MSComctlLib.Node)
Dim ItmX As ListItem
  
  LV.ListItems.Clear
  Set ItmX = LV.ListItems.Add(, "FILENAME", "File Name", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", DB_Name
  
  Set ItmX = LV.ListItems.Add(, "PASSWORD", "Jet OleDB:Password", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", mJetPassword
  
  Dim F As Integer
  For F = 0 To mCon.Properties.Count - 1
    Set ItmX = LV.ListItems.Add(, F & "Key", mCon.Properties(F).name, , "VARIABLE")
    ItmX.ListSubItems.Add , "VALUE", mCon.Properties(F).Value
  Next F
  
End Sub

Private Sub LV_LoadTABLE(ByVal Node As MSComctlLib.Node)
Dim ItmX As ListItem
Dim KEY As ADOX.KEY
Dim IDX As ADOX.index
Dim Column As ADOX.Column
Dim F As Integer
Dim sPK As String, sFK As String
On Error Resume Next
  
  LV.ListItems.Clear
  Set ItmX = LV.ListItems.Add(, "TABLENAME", "Table Name", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", mCat.Tables(Node.Text).name

' find Primary Keys (from the indexes)
  For Each IDX In mCat.Tables(Node.Text).Indexes
    If IDX.PrimaryKey Then
      For Each Column In IDX.Columns
        sPK = sPK & Column.name & "; "
      Next
    End If
  Next
  
  If Len(sPK) > 2 Then
    Set ItmX = LV.ListItems.Add(, "PKEY", "Primary Key(s)", , "VARIABLE")
    ItmX.ListSubItems.Add , "VALUE", Left$(sPK, Len(sPK) - 2)
  End If
' Find Foreign Keys
  For Each KEY In mCat.Tables(Node.Text).Keys
   If KEY.Type = adKeyForeign Then
     For Each Column In KEY.Columns
       sFK = sFK & Column.name & "; "
     Next
   End If
  Next
  
  Set ItmX = LV.ListItems.Add(, "FKEYS", "Foreign Key(s)", , "VARIABLE")
  If Len(sFK) > 0 Then ItmX.ListSubItems.Add , "VALUE", Left$(sFK, Len(sFK) - 2)
  
  For F = 0 To mCat.Tables(Node.Text).Properties.Count - 1
    Set ItmX = LV.ListItems.Add(, F & "Key", mCat.Tables(Node.Text).Properties(F).name, , "VARIABLE")
    ItmX.ListSubItems.Add , "VALUE", mCat.Tables(Node.Text).Properties(F).Value
  Next F
  
End Sub

Private Sub LV_LoadCOLUMN(ByVal Node As MSComctlLib.Node)
Dim ItmX As ListItem
Dim TName As String
Dim CName As String
Dim Pos As Byte

  Pos = InStr(1, Node.KEY, "\")
  
  TName = Left$(Node.KEY, Pos - 1)
  CName = Mid$(Node.KEY, Pos + 1)
  
  LV.ListItems.Clear
  Set ItmX = LV.ListItems.Add(, "TABLENAME", "Table Name", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", TName
  
  Set ItmX = LV.ListItems.Add(, "COLUMNNAME", "Column Name", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", CName

  Set ItmX = LV.ListItems.Add(, "DATATYPE", "Data Type", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", cType(mCat.Tables(TName).Columns(CName).Type)


Dim F As Integer

  For F = 0 To mCat.Tables(TName).Columns(CName).Properties.Count - 1
    
    Set ItmX = LV.ListItems.Add(, F & "Key", mCat.Tables(TName).Columns(CName).Properties(F).name, , "VARIABLE")
    If ItmX.Text <> "Nullable" Then
      ItmX.ListSubItems.Add , "VALUE", mCat.Tables(TName).Columns(CName).Properties(F).Value
    Else
      ItmX.ListSubItems.Add , "VALUE", Not mCat.Tables(TName).Columns(CName).Properties(F).Value
    End If
    
  Next F

End Sub

Private Sub LV_LoadQuery(ByVal Node As MSComctlLib.Node)
Dim ItmX As ListItem
Dim QName As String
Dim DateCreated As Variant
Dim DateModified As Variant
Dim CMD As ADODB.Command

  If Node.Tag = "VIEW" Then
    QName = mCat.Views(Node.KEY).name
    DateCreated = mCat.Views(Node.KEY).DateCreated
    DateModified = mCat.Views(Node.KEY).DateModified
    Set CMD = mCat.Views(Node.KEY).Command
    
  ElseIf Node.Tag = "PROC" Then
    QName = mCat.Procedures(Node.KEY).name
    DateCreated = mCat.Procedures(Node.KEY).DateCreated
    DateModified = mCat.Procedures(Node.KEY).DateModified
    Set CMD = mCat.Procedures(Node.KEY).Command

  End If

  LV.ListItems.Clear
  
  If Len(QName) > 0 Then
    Set ItmX = LV.ListItems.Add(, "QNAME", "Query Name", , "VARIABLE")
    ItmX.ListSubItems.Add , "VALUE", QName
    
    Set ItmX = LV.ListItems.Add(, "TYPE", "Type", , "VARIABLE")
    ItmX.ListSubItems.Add , "VALUE", Node.Tag
    
    Set ItmX = LV.ListItems.Add(, "DC", "Date Created", , "VARIABLE")
    ItmX.ListSubItems.Add , "VALUE", DateCreated
    Set ItmX = LV.ListItems.Add(, "DM", "Date Modified", , "VARIABLE")
    ItmX.ListSubItems.Add , "VALUE", DateModified
    
    Set ItmX = LV.ListItems.Add(, "CMDTEXT", "Command Text", , "VARIABLE")
    ItmX.ListSubItems.Add , "VALUE", Replace(Replace(CMD.CommandText, vbCrLf, " "), """", "'")
  
  End If
  

End Sub

Private Sub LV_LoadINDEX(ByVal Node As MSComctlLib.Node)
Dim ItmX As ListItem
Dim TName As String
Dim IName As String
Dim CName As String
Dim Col As ADOX.Column
Dim Pos As Byte

  Pos = InStr(1, Node.KEY, "\")
  
  TName = Mid$(Node.KEY, 5, Pos - 5)
  IName = Mid$(Node.KEY, Pos + 1)
  
  LV.ListItems.Clear
  Set ItmX = LV.ListItems.Add(, "TABLENAME", "Table Name", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", TName
  
  Set ItmX = LV.ListItems.Add(, "INDEXNAME", "Index Name", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", IName
  
  Set ItmX = LV.ListItems.Add(, "COLUMNNAME", "Column Names", , "VARIABLE")
  For Each Col In mCat.Tables(TName).Indexes(IName).Columns
    CName = CName & Col.name & "; "
  Next
  ItmX.ListSubItems.Add , "VALUE", Left$(CName, Len(CName) - 2)
  
  Set ItmX = LV.ListItems.Add(, "UNIQUE", "Unique", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", IIf(mCat.Tables(TName).Indexes(IName).Unique, "True", "False")
  
  Set ItmX = LV.ListItems.Add(, "Clustered", "Clustered", , "VARIABLE")
  ItmX.ListSubItems.Add , "VALUE", IIf(mCat.Tables(TName).Indexes(IName).Clustered, "True", "False")
  
End Sub


Private Sub mnuHelpAbout_Click()
Dim S As String

  S = App.Title & " is a small database tool where you create a " & vbCrLf
  S = S & "BAS-module containing the structure to create an Access" & vbCrLf
  S = S & "database on the fly using ADO and ADOX." & vbCrLf & vbCrLf
  S = S & "Future features could be the data included in the BAS-module." & vbCrLf & vbCrLf
  S = S & "Any comments, please mail to: nikro@bigfoot.com"
  
  MsgBox S, vbApplicationModal + vbInformation, App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
  
End Sub

Private Sub mnuMRUFiles_Click(index As Integer)

  If index > 0 Then OpenDB mnuMRUFiles(index).Caption
  
End Sub

Private Sub SB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   Me.MousePointer = vbDefault

End Sub

Private Sub TV_Expand(ByVal Node As MSComctlLib.Node)
  Node.Sorted = 1 ' True
End Sub

Private Sub TV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     
  Me.MousePointer = vbDefault

End Sub

Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
  

  LockWindowUpdate Me.hwnd
  Select Case Node.Tag
    Case "COLUMN": LV_LoadCOLUMN Node
    Case "INDEX": LV_LoadINDEX Node
    Case "TABLE": LV_LoadTABLE Node
    Case "VIEW", "PROC": LV_LoadQuery Node
    Case "DATABASE": LV_LoadDATABASE 'Node
  End Select
    
  
  lvAutosizeControl LV
  LockWindowUpdate 0
End Sub

Private Sub ArrangeControls()
On Error Resume Next
Dim hgt1 As Single
Dim hgt2 As Single

' Don't bother if we're iconized.
  If WindowState = vbMinimized Then Exit Sub

  hgt1 = (Me.ScaleWidth - SPLITTER_WIDTH) * Percentage1
  TV.Move Me.ScaleLeft, Me.ScaleTop, hgt1, Me.ScaleHeight - SB.Height
    
  hgt2 = (Me.ScaleWidth - SPLITTER_WIDTH) - hgt1
  LV.Move hgt1 + SPLITTER_WIDTH, Me.ScaleTop, hgt2, Me.ScaleHeight - SB.Height
    
End Sub

