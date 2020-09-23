VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Password"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3255
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSavePassword 
      Caption         =   "&Save Password."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "&Database Password :"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mOKPressed As Boolean

Public Property Get OKPressed() As Boolean
  OKPressed = mOKPressed
End Property

Public Property Get CancelPressed() As Boolean
  CancelPressed = Not mOKPressed
End Property

Public Property Let Password(ByVal sData As String)
  txtPassword.Text = sData
End Property

Public Property Get Password() As String
  Password = txtPassword.Text
End Property

Public Property Get SavePassword() As Boolean
  SavePassword = (chkSavePassword.Value = vbChecked)
End Property

Public Property Let SavePassword(ByVal bData As Boolean)
  chkSavePassword.Value = IIf(bData, vbChecked, vbUnchecked)
End Property

Private Sub cmdCancel_Click()
  mOKPressed = False
  Me.Hide
End Sub

Private Sub cmdOk_Click()

  mOKPressed = True
  Me.Hide
  
End Sub

Private Sub Form_GotFocus()
    txtPassword.SetFocus
End Sub

Private Sub txtPassword_GotFocus()
  If Len(txtPassword) > 0 Then
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
  End If
End Sub

