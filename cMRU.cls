VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************
' Name: cMRU
' Description:An object to maintain the last files opened opened by a program.
' cMRU gives you a quick and reliable way to provide a Most-Recently Used (MRU)
' file list in your application.
'**************************************
Private colMRUFiles As New Collection
Private Num As Integer

Public Sub Add(ByVal sValue As String)
  
  If colMRUFiles.Count = 0 Then
    colMRUFiles.Add sValue, UCase$(sValue)
  Else
    
    On Error Resume Next
    colMRUFiles.Remove UCase$(sValue)
    colMRUFiles.Add sValue, UCase$(sValue), 1
    
    If colMRUFiles.Count > Num Then colMRUFiles.Remove colMRUFiles.Count
    
  End If
  
End Sub

Public Sub Remove(ByVal sValue As String)
On Error Resume Next
    
  colMRUFiles.Remove UCase$(sValue)

End Sub


Public Sub Clear()
' *** Clears all files from the list.
  Do While colMRUFiles.Count > 0
    colMRUFiles.Remove 1
  Loop
End Sub

Public Property Get Count() As Long
' *** Returns the number of files in the list.
  
  Count = colMRUFiles.Count
  
End Property

Public Property Get Item(ByVal nValue As Integer) As String
' *** Returns the nth item from the list
On Error GoTo ItemError
  
  Item = colMRUFiles(nValue)

Exit Property
ItemError:
  Item = ""
End Property

Public Sub Load()
Dim Col As Collection
Dim I As Integer
Dim HKey As HKEYS
Dim Section As String
Dim REG As cRegistry

  Set REG = New cRegistry
  
  If Not REG.IsWinNT Then
    HKey = HKEY_LOCAL_MACHINE
  Else
    HKey = HKEY_CURRENT_USER
  End If
  
  Section = "SOFTWARE\NiKroWare\" & App.Title & "\MRUFiles"
  ' If Not REG.CheckRegistryKey(HKey, Section) Then REG.CreateRegistryKey HKey, Section
  
  Set Col = REG.EnumRegistryValues(HKey, Section)
  
  Set REG = Nothing
  
  If Col.Count > 0 Then
    Me.Clear
    For I = Col.Count To 1 Step -1
      colMRUFiles.Add Col(I)(1), UCase$(Col(I)(1))
    Next I
  End If
  
End Sub

Public Sub Save()
Dim I As Integer
Dim HKey As HKEYS
Dim Section As String
Dim REG As cRegistry
On Error Resume Next

  Set REG = New cRegistry
  
  If Not REG.IsWinNT Then
    HKey = HKEY_LOCAL_MACHINE
  Else
    HKey = HKEY_CURRENT_USER
  End If
 
  Section = "SOFTWARE\NiKroWare\" & App.Title & "\MRUFiles"
  ' If Not REG.CheckRegistryKey(HKey, Section) Then REG.CreateRegistryKey HKey, Section
 
  REG.DeleteRegistryKey HKey, Section

'  DeleteSetting AppN, "colMRUFiles"
  
  If Not REG.CheckRegistryKey(HKey, Section) Then REG.CreateRegistryKey HKey, Section
  
  For I = 1 To colMRUFiles.Count
    REG.SetRegistryValue HKey, Section, I, colMRUFiles(I)
  Next I

End Sub


Public Property Get Number() As Integer
' *** Gets the maximum size of the list.
    
  Number = Num
  
End Property

Public Property Let Number(ByVal nValue As Integer)
' *** Sets the maximum size of the list.

  Num = nValue
    
End Property

Public Sub Update(F As Form)
'// Note: The form must contain a menu control array named mnuMRUFiles
'// that is at least as big as Number.
Dim I As Long
On Error GoTo NextStep

  For I = 0 To Num
    F.mnuMRUFiles(I).Visible = False
  Next I
  
NextStep:
On Error GoTo MenuEnd

  If colMRUFiles.Count > 0 Then
    F.mnuMRUFiles(0).Visible = True

    For I = 1 To colMRUFiles.Count
      F.mnuMRUFiles(I).Caption = colMRUFiles(I)
      F.mnuMRUFiles(I).Visible = True
    Next I

    Do
      F.mnuMRUFiles(I).Visible = False
      I = I + 1
    Loop
  Else
    I = 0
    Do
      F.mnuMRUFiles(I).Visible = False
      I = I + 1
    Loop
  End If
  
MenuEnd:
End Sub

Private Sub Class_Initialize()
  Num = 5
End Sub
