VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   1440
   End
   Begin VB.Frame fraMainFrame 
      Height          =   4600
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   7400
      Begin VB.TextBox txtBerita 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1000
         Left            =   1920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmSplash.frx":0000
         Top             =   3100
         Width           =   5295
      End
      Begin MSComctlLib.ProgressBar bar01 
         Height          =   255
         Left            =   100
         TabIndex        =   14
         Top             =   4200
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.PictureBox picLogo 
         Height          =   1515
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   3315
         TabIndex        =   3
         Top             =   1440
         Width           =   3375
         Begin VB.PictureBox Picture1 
            Height          =   1545
            Left            =   0
            Picture         =   "frmSplash.frx":0008
            ScaleHeight     =   1485
            ScaleWidth      =   3315
            TabIndex        =   10
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Silahkan Menggunakan"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   13
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label lblNama 
         Alignment       =   2  'Center
         Caption         =   "Halooo"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3480
         TabIndex        =   12
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label lblHalo 
         Alignment       =   2  'Center
         Caption         =   "Halooo"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   11
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Tag             =   "LicenseTo"
         Top             =   240
         Width           =   7095
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         Caption         =   "Products Electronic Lines (PEL) Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Tag             =   "Product"
         Top             =   840
         Width           =   7200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         Caption         =   "PT Schneider Electric Manufacturing Batam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   8
         Tag             =   "CompanyProduct"
         Top             =   480
         Width           =   7170
      End
      Begin VB.Label lblPlatform 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "PEL Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   50
         TabIndex        =   7
         Tag             =   "Platform"
         Top             =   3000
         Width           =   1800
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   50
         TabIndex        =   6
         Tag             =   "Version"
         Top             =   3360
         Width           =   1800
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company: SEMB"
         Height          =   255
         Left            =   50
         TabIndex        =   5
         Tag             =   "Company"
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright: (C) YK 2009"
         Height          =   255
         Left            =   50
         TabIndex        =   4
         Tag             =   "Copyright"
         Top             =   3720
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Progress As Integer
'Public TimerSet As Boolean

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblLicenseTo.Caption = "YK"
    
'    lblNama.Caption = NamaPengguna
    
    bar01.Value = 0
'    Progress = 0
    txtBerita.Text = "- Mulai menjalannkan program" & vbCrLf
    Label1.Caption = "Harap Tunggu!!!!"
    Label1.ForeColor = vbRed
'    Timer1.Interval = 100
'    TimerSet = True
'    If TimerSet Then
'        Timer1.Enabled = True
'    Else
'        Timer1.Enabled = False
'    End If
    
'    Me.Show
'    Call Main
    
End Sub

'Private Sub Timer1_Timer()
'Dim I As Integer
    
'    Timer1.Enabled = False
'    Progress = Progress + 1
'    If Progress = 100 Then Progress = 0
'    Tampilan (Progress)
'    bar01.Value = Progress
'    If TimerSet Then
'        Timer1.Enabled = True
'    Else
'        Timer1.Enabled = False
'    End If
        
'End Sub

Public Sub Tampilan(Teks As String)
    txtBerita.Text = txtBerita.Text & Teks & vbCrLf
    txtBerita.SelStart = Len(txtBerita.Text)
End Sub

Public Sub UbahProgress(Nilai As Integer)
    bar01.Value = Nilai
End Sub

Private Sub fraMainFrame_Click()
    If ProsesSts(0) Then
        End
    End If
End Sub

