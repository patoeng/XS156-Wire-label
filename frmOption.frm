VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOption 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Option"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd05 
      Caption         =   "Codesoft Option"
      Height          =   495
      Left            =   6240
      TabIndex        =   35
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Cmd04 
      Caption         =   "Seting Printer"
      Height          =   495
      Left            =   6240
      TabIndex        =   34
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Cmd03 
      Caption         =   "Pilih Printer"
      Height          =   495
      Left            =   6240
      TabIndex        =   33
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Cmd02 
      Caption         =   "Seting Print"
      Height          =   495
      Left            =   6240
      TabIndex        =   32
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Cmd01 
      Caption         =   "Seting Halaman"
      Height          =   495
      Left            =   6240
      TabIndex        =   31
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Printer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   7200
      Width           =   4455
      Begin VB.ComboBox cboPrinter 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Pilih Linenya"
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Software Setting"
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   6015
      Begin VB.CheckBox chkSatuKeluarga 
         Caption         =   "No"
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtSetup 
         Height          =   375
         Index           =   2
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtSetup 
         Height          =   375
         Index           =   1
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtSetup 
         Height          =   375
         Index           =   0
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Satu Keluarga"
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Posisi Kiri"
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Posisi Atas"
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Keluarga"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnHide 
      Caption         =   "Hide"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   6120
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Database Server Setting"
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   6015
      Begin VB.TextBox txtServer 
         Height          =   375
         Index           =   5
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtServer 
         Height          =   375
         Index           =   4
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox txtServer 
         Height          =   375
         Index           =   3
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtServer 
         Height          =   375
         Index           =   2
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtServer 
         Height          =   375
         Index           =   1
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtServer 
         Height          =   375
         Index           =   0
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Option"
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   41
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   39
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Directory Path"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtPath 
         Height          =   375
         Index           =   4
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Index           =   3
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Index           =   2
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Index           =   1
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Index           =   0
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Temporary Path"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gambar Path"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Database Path"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label Path"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Program Path"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   6720
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PanggilDialog(Nilai As Integer)
    Select Case Nilai
    Case 1
        CS7Server.Dialogs.Item(lppxPageSetupDialog).Show (Me.hWnd)
    Case 2
        CS7Server.Dialogs.Item(lppxPrinterSelectDialog).Show (Me.hWnd)
    Case 3
        CS7Server.Dialogs.Item(lppxPrintDialog).Show (Me.hWnd)
    Case 4
        CS7Server.Dialogs.Item(lppxPrinterSetupDialog).Show (Me.hWnd)
    Case 5
        CS7Server.Dialogs.Item(lppxOptionsDialog).Show (Me.hWnd)
    End Select
End Sub

Private Sub btnExit_Click()
    PathLabel = Me.txtPath(1).Text
    SaveSetting App.Title, "Settings", "Label", PathLabel
    PathData = Me.txtPath(2).Text
    SaveSetting App.Title, "Settings", "Data", PathData
    PathGambar = Me.txtPath(3).Text
    SaveSetting App.Title, "Settings", "Gambar", PathGambar
    PathTemp = Me.txtPath(4).Text
    SaveSetting App.Title, "Settings", "Sementara", PathTemp
    
    ServerAlamat = Me.txtServer(0).Text
    SaveSetting App.Title, "Settings", "Server", ServerAlamat
    ServerData = Me.txtServer(1).Text
    SaveSetting App.Title, "Settings", "ServerData", ServerData
    ServerUser = Me.txtServer(2).Text
    SaveSetting App.Title, "Settings", "ServerUser", ServerUser
    ServerPass = Me.txtServer(3).Text
    SaveSetting App.Title, "Settings", "ServerPass", ServerPass
    ServerDriver = Me.txtServer(4).Text
    SaveSetting App.Title, "Settings", "ServerDriver", ServerDriver
    ServerOption = Me.txtServer(5).Text
    SaveSetting App.Title, "Settings", "ServerOption", ServerOption

    Keluarga = Me.txtSetup(0).Text
    SaveSetting App.Title, "Settings", "Keluarga", Keluarga
    PosisiAtas = Me.txtSetup(1).Text
    SaveSetting App.Title, "Settings", "PosisiAtas", PosisiAtas
    PosisiKiri = Me.txtSetup(2).Text
    SaveSetting App.Title, "Settings", "PosisiKiri", PosisiKiri
    SatuKeluarga = chkSatuKeluarga.Value
    SaveSetting App.Title, "Settings", "SatuKeluarga", SatuKeluarga
    
    Unload Me
End Sub

Private Sub btnHide_Click()
    If (btnHide.Caption = "Hide") Then IsVisible = True
    IsVisible = Not IsVisible
    ServerVisible IsVisible
    If IsVisible Then btnHide.Caption = "Hide" Else btnHide.Caption = "unHide"
End Sub

Private Sub chkSatuKeluarga_Click()
    If chkSatuKeluarga.Value Then
        chkSatuKeluarga.Caption = "Yes"
    Else
        chkSatuKeluarga.Caption = "No"
    End If
End Sub

Private Sub Cmd01_Click()
    Call PanggilDialog(1)
End Sub

Private Sub Cmd02_Click()
    Call PanggilDialog(3)
End Sub

Private Sub Cmd03_Click()
    Call PanggilDialog(2)
End Sub

Private Sub Cmd04_Click()
    Call PanggilDialog(4)
End Sub

Private Sub Cmd05_Click()
    Call PanggilDialog(5)
End Sub

Private Sub Form_Load()
Dim Nilai As String

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.txtPath(0).Text = App.Path
    Me.txtPath(1).Text = GetSetting(App.Title, "Settings", "Label")
    Me.txtPath(2).Text = GetSetting(App.Title, "Settings", "Data")
    Me.txtPath(3).Text = GetSetting(App.Title, "Settings", "Gambar")
    Me.txtPath(4).Text = GetSetting(App.Title, "Settings", "Sementara")
    
    Me.txtServer(0).Text = GetSetting(App.Title, "Settings", "Server")
    Me.txtServer(1).Text = GetSetting(App.Title, "Settings", "ServerData")
    Me.txtServer(2).Text = GetSetting(App.Title, "Settings", "ServerUser")
    Me.txtServer(3).Text = GetSetting(App.Title, "Settings", "ServerPass")
    Me.txtServer(4).Text = GetSetting(App.Title, "Settings", "ServerDriver")
    Me.txtServer(5).Text = GetSetting(App.Title, "Settings", "ServerOption")
    
    Me.txtSetup(0).Text = GetSetting(App.Title, "Settings", "Keluarga")
    Me.txtSetup(1).Text = GetSetting(App.Title, "Settings", "PosisiAtas")
    Me.txtSetup(2).Text = GetSetting(App.Title, "Settings", "PosisiKiri")
    Nilai = GetSetting(App.Title, "Settings", "SatuKeluarga")
'    Me.chkSatuKeluarga.Value = CInt(Nilai)
    If Nilai = "True" Then
        Me.chkSatuKeluarga.Value = 1
        Me.chkSatuKeluarga.Caption = "Yes"
    Else
        Me.chkSatuKeluarga.Value = 0
        Me.chkSatuKeluarga.Caption = "No"
    End If
    
    If CS7Server.Visible Then btnHide.Caption = "Hide" Else btnHide.Caption = "unHide"
End Sub

Private Sub txtPath_Click(Index As Integer)
    CommonDialog1.ShowOpen
    txtPath(Index).Text = Left(CommonDialog1.FileName, InStr(1, CommonDialog1.FileName, CommonDialog1.FileTitle) - 2)
End Sub
