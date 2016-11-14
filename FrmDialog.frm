VERSION 5.00
Begin VB.Form FrmDialog 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   1440
   ClientLeft      =   4650
   ClientTop       =   6435
   ClientWidth     =   6030
   Icon            =   "FrmDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1440
   ScaleMode       =   0  'User
   ScaleWidth      =   5030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "@"
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Isi Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "FrmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Perintah As String

Private Sub CancelButton_Click()
'CancelBut = True'
    Perintah = ""
    Unload Me '
    frmMain.SetFocus
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Text1.Text = ""
End Sub

Private Sub OKButton_Click()
'If (Val(Text1.Text) <= 0 Or Val(Text1.Text) > 32767) Then'
'    MsgBox "Jumlah label tidak valid"'
'    Text1.Text = ""'
'    Text1.SetFocus'
'Else'
'    Qty_printed = Val(Text1.Text)'
'    CancelBut = False'
'    Unload Me'
'    frmMain.SetFocus'
    Select Case Perintah
    Case "Tanggal"
        If FrmDialog.Text1.Text <> PassOperator Then
            MsgBox ("Maaf, Anda tidak memiliki akses untuk ini")
            Perintah = ""
            Unload Me
        Else
            MsgBox ("Silahkan Masukan tanggal yang diinginkan")
            Perintah = ""
            Unload Me
            'frmMain.SetFocus'
            frmMain.Label5.Visible = True
            frmMain.Textdate.Visible = True
            frmMain.Textdate.Enabled = True
            frmMain.Textdate.Locked = True
            frmMain.MonthView1.Visible = True
            frmMain.MonthView1.Value = Now
        '    frmMain.TextDate.SetFocus'
        End If
    Case "Setup"
        If FrmDialog.Text1.Text <> PassAdmin Then
            MsgBox "Maaf, Anda Tidak memiliki akses untuk ini"
            Perintah = ""
            Unload Me
        Else
            MsgBox "Silahkan Melakukan Setup"
            Perintah = ""
            Unload Me
            frmOption.Show vbModal
        End If
    Case "Codesoft"
        If FrmDialog.Text1.Text <> PassAdmin Then
            MsgBox "Maaf, Anda Tidak memiliki akses untuk ini"
            Perintah = ""
            Unload Me
        Else
            MsgBox "Silahkan edit Codesoft"
            IsVisible = Not IsVisible
            ServerVisible IsVisible
        End If
    End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            OKButton_Click
        Case 27
            CancelButton_Click
    End Select
End Sub
