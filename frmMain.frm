VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packaging Software"
   ClientHeight    =   10710
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15240
   ControlBox      =   0   'False
   FillColor       =   &H80000002&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000002&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame8 
      BackColor       =   &H008080FF&
      Caption         =   "Posisi"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8160
      TabIndex        =   56
      Top             =   2040
      Width           =   3615
      Begin VB.TextBox TxtKiri 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   60
         Text            =   "Text4"
         ToolTipText     =   "Buat Geser Kiri Kanan"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TxtAtas 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   59
         Text            =   "Text3"
         ToolTipText     =   "Buat Geser Atas Bawah"
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptPosisi 
         BackColor       =   &H0000FFFF&
         Caption         =   "Kebalik"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Posisi Keluar Terbalik"
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton OptPosisi 
         BackColor       =   &H0000FFFF&
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Posisi keluar Normal"
         Top             =   600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C000&
         Caption         =   "Kiri"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   62
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "Atas"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   61
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   10800
      ScaleHeight     =   2835
      ScaleWidth      =   3915
      TabIndex        =   55
      Top             =   6120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H008080FF&
      Caption         =   "Label Type"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   12120
      TabIndex        =   52
      Top             =   960
      Width           =   3015
      Begin VB.OptionButton OptType 
         BackColor       =   &H0000FFFF&
         Caption         =   "Special"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Buat Label Packing Group"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.OptionButton OptType 
         BackColor       =   &H0000FFFF&
         Caption         =   "Date Code"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Buat Label Packing Group"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.OptionButton OptType 
         BackColor       =   &H0000FFFF&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Buat Label Packing Group"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton OptType 
         BackColor       =   &H0000FFFF&
         Caption         =   "Packing Group"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Buat Label Packing Group"
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton OptType 
         BackColor       =   &H0000FFFF&
         Caption         =   "Individual"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Buat Label Individual"
         Top             =   600
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H008080FF&
      Caption         =   "Reference"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4080
      TabIndex        =   50
      Top             =   2040
      Width           =   3735
      Begin VB.TextBox txtRef 
         BackColor       =   &H0080FF80&
         Height          =   465
         Left            =   240
         TabIndex        =   0
         Text            =   "Text3"
         ToolTipText     =   "Isi tipenya atau scan"
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox cboModel 
         BackColor       =   &H0080FF80&
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   51
         ToolTipText     =   "Pilih Modelnya"
         Top             =   720
         Width           =   3255
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H008080FF&
      Caption         =   "Family"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   48
      Top             =   2040
      Width           =   3615
      Begin VB.ComboBox cboFamily 
         BackColor       =   &H0080FF80&
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   49
         ToolTipText     =   "Pilih Familnya"
         Top             =   720
         Width           =   3135
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      Caption         =   "Line"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   46
      Top             =   2040
      Width           =   2295
      Begin VB.ComboBox cboLine 
         BackColor       =   &H0080FF80&
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
         Left            =   240
         TabIndex        =   47
         ToolTipText     =   "Pilih Linenya"
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   11640
      Top             =   1320
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   11640
      TabIndex        =   45
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C000&
      Caption         =   "Scanner"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H80000000&
      TabIndex        =   42
      ToolTipText     =   "Beri tanda centang bila ingin mengunakan scanner"
      Top             =   1440
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   12000
      TabIndex        =   41
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtchar9 
      Height          =   285
      Left            =   240
      TabIndex        =   40
      Text            =   "Char9"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   10680
      TabIndex        =   38
      ToolTipText     =   "Pilih Date code yang diinginkan"
      Top             =   3720
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   16711680
      BackColor       =   16744576
      Appearance      =   1
      MonthBackColor  =   8454143
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   65994754
      TrailingForeColor=   8421504
      CurrentDate     =   38856
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   10335
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10920
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777088
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C290
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D0E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D6C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DB19
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CSToggle"
            Object.ToolTipText     =   "Tampilkan CodeSoft"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CSHelp"
            Object.ToolTipText     =   "About this program"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CSExit"
            Object.ToolTipText     =   "Exit this program"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CSNyambung"
            Object.ToolTipText     =   "Buat nyambungkan Codesoft"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CSDatabase"
            Object.ToolTipText     =   "Buat nyambung ke Database"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   4815
      Left            =   12600
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox TxtIta 
         Height          =   285
         Left            =   240
         TabIndex        =   31
         Text            =   "Italiano"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox TxtSpain 
         Height          =   285
         Left            =   240
         TabIndex        =   25
         Text            =   "Spain"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox TxtBitmap 
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Text            =   "Bitmap"
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox TxtPower 
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Text            =   "Power"
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox TxtLoadpower 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Text            =   "LoadPower"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox TxtCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Text            =   "Current"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtType 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Text            =   "Type for Schile only"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtMaterialNumber 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Text            =   "MaterialNumber"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txtLabelSize 
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Text            =   "LabelSize"
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Text            =   "Quantity"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtBarcode 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Text            =   "Barcode"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox txtArticleNumber 
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Text            =   "ArticleNumber"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtReference 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Text            =   "Reference"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtVoltage 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Text            =   "Voltage"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtGerman 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Text            =   "German"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtFrance 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Text            =   "France"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtEnglish 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Text            =   "English"
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdStorePass 
      Caption         =   "Set Printer"
      Height          =   615
      Left            =   3960
      MaskColor       =   &H80000000&
      TabIndex        =   6
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   12120
      Top             =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   10575
      Begin VB.Label lblmessagebox 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   10335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Label information"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   10575
      Begin VB.TextBox TxtQty2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         TabIndex        =   44
         ToolTipText     =   "Jumlah yang akan diprint"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtmodel2 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   360
         TabIndex        =   43
         Top             =   960
         Width           =   5295
      End
      Begin VB.CommandButton Datechange 
         Caption         =   "Change Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         MaskColor       =   &H80000000&
         TabIndex        =   39
         ToolTipText     =   "Untuk Mengubah Date Code"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton CommandCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         MaskColor       =   &H80000000&
         TabIndex        =   37
         ToolTipText     =   "Untuk Batalin"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CommandOK 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Untuk Ngeprint"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Textdate 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   6480
         TabIndex        =   33
         ToolTipText     =   "Pilih dari tanggalan disamping"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox TextQty 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   6480
         TabIndex        =   32
         Text            =   "1"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox TxtDisplayRef 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   960
         HideSelection   =   0   'False
         Left            =   360
         TabIndex        =   26
         ToolTipText     =   "Reference yang siap diprint"
         Top             =   2400
         Width           =   5295
      End
      Begin VB.TextBox txtModel 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         HideSelection   =   0   'False
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Date (yymm)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6480
         TabIndex        =   35
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   34
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   28
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Article Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   27
         Top             =   480
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buat Keluar"
      Top             =   9240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "XUX Label Printing Station"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   15255
   End
   Begin VB.Label lblSoftwarever 
      BackColor       =   &H00C0C000&
      Caption         =   "Software Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   9720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Menu CSFile 
      Caption         =   "&File"
      Begin VB.Menu CSToggle 
         Caption         =   "&CS Server"
      End
      Begin VB.Menu CSExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnu_Option 
      Caption         =   "&Option"
   End
   Begin VB.Menu mnu_About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Loading As Boolean

Private Sub cboFamily_click()
Dim Kondisi As Boolean

    PesanSalah = ""
    frmMain.lblmessagebox.Caption = "Tunggu sedang ambil data Familynya...."
    OptPosisi(1).Value = True
    OptType(1).Value = True
    Picture1.Visible = False
    SetFrame Frame1, False
    SetFrame Frame6, False
    SetFrame Frame7, False
    SetFrame Frame8, False
    Check3.Enabled = False
    Check3.Visible = False
    Keluarga = cboFamily.Text
    If Keluarga = "" Then
        frmMain.lblmessagebox.Caption = "Silahkan Pilih Family...."
    Else
        IsiCombo "Ref", ServerData, "Label", "Group`, `Ref` ASC", cboModel, Kondisi, "Group` = '" & Keluarga & "' and `Aktif`='1'"
        If Not Kondisi Then
            PesanSalah = "Gagal ambil data model"
            GoTo Salah
        End If
        Kondisi = Bukalab(Keluarga)
        If Not Kondisi Then
            PesanSalah = "Problem saat membuka template file"
            GoTo Salah
        Else
            SetFrame Frame6, True
            SetFrame Frame7, True
            SetFrame Frame8, True
            Check3.Visible = True
            Check3.Enabled = True
            Check3.Value = 0
            Call OptType_Click(1)
            Call OptPosisi_Click(1)
        End If
        Call Check3_Click
        frmMain.lblmessagebox.Caption = "Silahkan Pilih Modelnya...."
    End If
    StatusBar1.Panels.Item(2).Text = "OK"
    Exit Sub
    
Salah:
    If PesanSalah = "" Then
        PesanSalah = Err.Description
    End If
    lblmessagebox.Caption = PesanSalah
    StatusBar1.Panels.Item(2).Text = "Problem"
    Timer2.Enabled = True
    
End Sub

Private Sub Form_Initialize()
Dim I As Byte
Dim Kondisi As Boolean
Dim Nama As String

On Error GoTo Salah
    Loading = True
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    txtRef.Text = ""
    TxtAtas.Text = PosisiAtas
    TxtKiri.Text = PosisiKiri
    For I = 1 To 4
        StatusBar1.Panels.Add ' Add 2 panels.
    Next I

    With StatusBar1.Panels
        .Item(1).AutoSize = sbrSpring
        .Item(1).Text = "Tanggal : " & " <" & Format(Date, "dd mmmm yyyy") & ">, Minggu ke : " & Format(Date, "yyww")
        .Item(2).AutoSize = sbrSpring
        .Item(3).AutoSize = sbrSpring
        .Item(4).Style = sbrNum                             ' NumLock
        .Item(5).Style = sbrCaps
    End With
    lblSoftwarever.Caption = "Software Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblSoftwarever.Visible = True
    cmdStorePass.Visible = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(5).Visible = False
    
    If ProsesSts(0) Then
        StatusBar1.Panels.Item(3).Text = "Database: OK, CodeSoft: " & CSStatus
        SetFrame Frame1, False
        SetFrame Frame2, True
        SetFrame Frame4, False
'        Frame4.Visible = False
'        Frame5.Visible = True
        
        If SatuKeluarga Then
            Label1.Caption = Keluarga & " Label Printing Station"
            SetFrame Frame5, False
            Frame5.Visible = False
            IsiCombo "Model", _
                ServerData, _
                "Label` LEFT JOIN `" & ServerData & "`.`family` on `" & ServerData & "`.`label`.`group` = `" & ServerData & "`.`family`.`nama", _
                ServerData & "`.`label`.`ref` ASC", _
                cboModel, _
                Kondisi, _
                ServerData & "`.`family`.`mesin`='3' and `" & ServerData & "`.`label`.`aktif`='1'"
            If Not Kondisi Then
                PesanSalah = "Gagal ambil data model"
                GoTo Salah
            End If
'            Call BuatModel(Keluarga)
'            Kondisi = Bukalab(Keluarga)
            'If Not Kondisi Then
            '    SetFrame Frame6, False
            '    SetFrame Frame7, False
            '    SetFrame Frame8, False
            '    Check3.Enabled = False
            '    Check3.Visible = False
            '    PesanSalah = "Problem saat membuka template file"
            '    GoTo Salah
            'Else
            '    SetFrame Frame6, True
            '    SetFrame Frame7, True
            '    SetFrame Frame8, True
            '    Check3.Visible = True
            '    Check3.Enabled = False
            '    Check3.Value = 1
            '    Call OptType_Click(1)
            '    Call OptPosisi_Click(1)
            'End If
            'Call Check3_Click
        Else
            Label1.Caption = "XS156 Label Printing Station"
            IsiCombo "Nama", ServerData, "Family", "Nama` ASC", cboFamily, Kondisi, "Mesin3`='1'"
            If Not Kondisi Then
                PesanSalah = "Gagal ambil data family"
                GoTo Salah
            End If
            SetFrame Frame5, True
            Frame5.Visible = True
        End If
    Else
        If Not ProsesSts(2) Then
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(6).Enabled = True
        End If
        If Not ProsesSts(3) Then
            Toolbar1.Buttons(5).Visible = True
            Toolbar1.Buttons(5).Visible = True
        End If
        PesanSalah = "Seting Parameter bermasalah"
        StatusBar1.Panels.Item(3).Text = "Database: Problem, CodeSoft: " & CSStatus
        GoTo Salah
    End If

'Textdate.Text = Format(Now, "yyww", vbMonday, vbFirstJan1)
'MonthView1.Value = Now
    StatusBar1.Panels.Item(2).Text = "OK"
    MonthView1_DateClick Now
    Loading = False
    ModelPilih = ""
    Exit Sub

Salah:
    lblmessagebox.Caption = PesanSalah
    StatusBar1.Panels.Item(2).Text = "Problem"
    Timer2.Enabled = True
'    MsgBox "Problem: (Inisialisasi Main)" & vbCrLf & PesanSalah & vbCrLf & Err.Number & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub Form_Load()

On Error Resume Next
    txtModel.Text = ""
    lblmessagebox.Caption = "Silahkan Pilih Modelnya"
    lblmessagebox.BackColor = &HC0C000
    OptType(1).Value = True
    If IsVisible Then frmOption.btnHide.Caption = "Hide" Else frmOption.btnHide.Caption = "unHide"
    scan = 1
    If Loading Then
        Loading = True
    Else
        Loading = False
    End If
    Exit Sub
Error_msg:
    MsgBox "Problem: (Main Load) dan tidak dapat melanjutkan.", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CS7Server_Stop
    End
End Sub

Private Sub Form_Terminate()
    cmdExit_Click
End Sub

Private Sub Printing()
Dim HasilPrint
Dim response As String
Dim Lokasi$, Pesan$
On Error GoTo ErrorHandle

    If CS7Server.Documents.Count < 1 Then
        response = "Error!! No open document"
        GoTo ErrorHandle
    End If
    
    frmMain.lblmessagebox.Caption = "Printing sedang dilakukan..."
    Lokasi$ = App.Path & "\DataLok\" & Format(Date, "yyww", vbSunday, vbFirstFullWeek) & ".log"
    Pesan = Day(Date) & "/" & Month(Date) & "/" & Year(Date) & ";" & Time & ";" & cboLine.List(cboLine.ListIndex) & ";" & cboFamily.List(cboFamily.ListIndex) & ";" & cboModel.List(cboModel.ListIndex) & ";" & TextQty.Text & ";" & Textdate.Text
    
    Set Doc = CS7Server.ActiveDocument
    Doc.VertPrintOffset = Val(TxtAtas.Text)
    Doc.HorzPrintOffset = Val(TxtKiri.Text)

    If CancelBut Then Exit Sub
    Qty_printed = Int(TextQty.Text)
'    HasilPrint = Doc.PrintLabel(Qty_printed, 1)
    HasilPrint = Doc.PrintDocument(Qty_printed)
    If HasilPrint = 1 Then HasilPrint = True Else HasilPrint = False
    If Not HasilPrint Then
        response = "Problem: tak dapat nge-print."
        GoTo ErrorHandle
    End If


'    Delay (1.75)

    frmMain.lblmessagebox.Caption = "Printing Selesai !! "
    Logdata Lokasi$, Pesan$
    Exit Sub 'necessary if not it will always go into Errorhandle routine

ErrorHandle:
    MsgBox response & vbCrLf & "Printing Problem", vbExclamation, "Warning"
'       MsgBox "Missing file .Lab, .mdb file or tables or unable to find the reference", vbExclamation, "Warning"
End Sub


Private Sub cmdStorePass_Click()
Dim PrintName, I, PrintPort
    I = InStr(1, cboFamily.List(cboFamily.ListIndex), ",", vbTextCompare)
    PrintName = Mid(cboFamily.List(cboFamily.ListIndex), 1, I - 1)
    PrintPort = Mid(cboFamily.List(cboFamily.ListIndex), I + 1, 100)
    CS7Server.ActiveDocument.Printer.SwitchTo PrintName, PrintPort
End Sub

Sub SetFrame(kotak As Frame, Aktif As Boolean)
On Error GoTo Salah
    If Aktif Then
        kotak.BackColor = &HC0C000      'untuk warna biru FF8080
    Else
        kotak.BackColor = &H8080FF       'untuk warna abu-abu &HF0F0F0
    End If
    kotak.Enabled = Aktif
    Exit Sub
    
Salah:
    MsgBox "Problem: (Set Frame)" & vbCrLf & PesanSalah & vbCrLf & Err.Number & vbCrLf & Err.Description, vbExclamation
End Sub

Private Function Penyegaran() As Boolean

On Error GoTo Salah
    PesanSalah = ""
'    Set Doc = CS7Server.ActiveDocument
'    Set CSVar = Doc.Variables.FormVariables("Model")
'    CSVar.Value = ModelPilih
'    Set CSVar = Doc.Variables.FormVariables("DateCode")
'    CSVar.Value = Textdate.Text
    Doc.Variables.FormVariables("Model").Value = ModelPilih
    Doc.Variables.FormVariables("DateCode").Value = Textdate.Text
    Set CSVar = Nothing
    If AmbilGambar Then
        SetFrame Frame1, True
        frmMain.lblmessagebox.Caption = "Silahkan isi jumlah yang akan diprint..."
    Else
        SetFrame Frame1, False
        PesanSalah = "Ada problem dalam ambil data...!!!"
        GoTo Salah
    End If
    
    Penyegaran = True
    Exit Function
    
Salah:
    If PesanSalah = "" Then PesanSalah = Err.Description
'    PesanSalah = "Problem : Penyegaran gagal" & vbCrLf & Err.Description
    Penyegaran = False
End Function

Private Function AmbilGambar() As Boolean
On Error GoTo Salah
    If CS7Server.Documents.Count < 1 Then
        Exit Function
    End If
     
    CS7Server.ActiveDocument.ViewMode = lppxViewModeValue
    CS7Server.ActiveDocument.CopyToClipboard
    Picture1.Visible = False
    Picture1.AutoSize = True
    Picture1.Picture = Clipboard.GetData(vbCFMetafile)
    DoEvents
    Picture1.Visible = True
    AmbilGambar = True
    
    Exit Function

Salah:
    PesanSalah = "Problem saat ambil gambar"
'    MsgBox "Problem: (Ambil Preview) hubungi test enginner.", vbExclamation
    AmbilGambar = False
End Function

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        cboModel.Visible = False
        txtRef.Visible = True
        If Not Loading Then
            txtRef.SetFocus
        End If
    Else
        cboModel.Visible = True
        txtRef.Visible = False
    End If
End Sub

Private Sub HapusModel()
'Dim A
'    On Error GoTo Salah
    cboModel.Clear
'    A = cboModel.ListCount
'    Do While A > 0
'        cboModel.RemoveItem (A - 1)
'        A = cboModel.ListCount
'    Loop
    SetFrame Frame6, False
    Exit Sub

'Salah:
'    MsgBox "Ada kesalahan (Hapus Model tak jalan) hubungi test enginner.", vbExclamation

End Sub

'Private Sub BuatModel(Family As String)
''Dim DBHis
''Dim RS As ADODB.Recordset
'Dim Kondisi As Boolean
'Dim nilai As Integer
'Dim A As String
'Dim Pesan As String
    
'    On Error GoTo Salah
'        Kondisi = BukaKoneksi
    
'        Set DataRS = New ADODB.Recordset
''        SQL = "SELECT `model` FROM Label where `Group`='" & Family & "' AND `Aktif`='1' ORDER BY `model`"
'        SQL = "SELECT `model` FROM Label where `Group`='" & Family & "' ORDER BY `model`"
'        DataRS.CursorLocation = adUseClient
'        DataRS.Open SQL, Conn, adOpenStatic, adLockReadOnly
'        nilai = DataRS.RecordCount
'        If nilai < 1 Then
'            Pesan = "Ada kesalahan (Data tak bisa diambil)"
'            GoTo Salah
'        Else
'            DataRS.MoveFirst
'            While Not DataRS.EOF
'                A = DataRS.Fields("Model").Value
'                cboModel.AddItem A
'                DataRS.MoveNext
'            Wend
'        End If
'        DataRS.Close
'        TutupKoneksi
''    Conn.Close
'    Set DataRS = Nothing
''    Set Conn = Nothing
'    Exit Sub

'Salah:
'    HapusModel
'    MsgBox Pesan & vbCrLf & Err.Number & " : " & Err.Description, vbExclamation
'    If Not (TypeName(DataRS) = "Nothing") Then
'        Set DataRS = Nothing
'        DataRS.Close
'    End If
'    If Not (TypeName(Conn) = "Nothing") Then
'        TutupKoneksi
'    End If
'End Sub

Private Sub CommandCancel_Click()
    CancelBut = True
    'Unload Me'
    scan = 1
    Text2.Text = "2"

    frmMain.SetFocus
    frmMain.Show
    frmMain.Label4.Visible = False
    frmMain.TextQty.Visible = False
    frmMain.Label5.Visible = False
    frmMain.Textdate.Visible = False
    frmMain.CommandOK.Visible = False
    frmMain.CommandCancel.Visible = False
    frmMain.Datechange.Visible = False
    frmMain.MonthView1.Visible = False
    frmMain.TxtQty2.Visible = False

    lblmessagebox.Caption = "Scan-lah barcode untuk print label..."
    lblmessagebox.BackColor = &H8080FF
    frmMain.txtModel = ""
    frmMain.txtRef = ""
    frmMain.Textdate = ""
    frmMain.TextQty = ""
    frmMain.TxtQty2 = ""
    frmMain.txtmodel2 = ""
End Sub

Private Sub cmdExit_Click()
    SaveSetting App.Title, "Settings", "PosisiAtas", TxtAtas.Text
    SaveSetting App.Title, "Settings", "PosisiKiri", TxtKiri.Text
    Dialog.Show vbModal
End Sub


Private Sub mnu_About_Click()
    frmAbout.Show vbModal
End Sub

Private Sub CSExit_Click()
    cmdExit_Click
End Sub

Private Sub CSToggle_Click()
    FrmDialog.Perintah = "Codesoft"
    FrmDialog.Show vbModal
'    IsVisible = Not IsVisible
'    ServerVisible IsVisible
End Sub

Private Sub Datechange_Click()
    FrmDialog.Perintah = "Tanggal"
    FrmDialog.Show vbModal
End Sub

Private Sub mnu_Option_Click()
    FrmDialog.Perintah = "Setup"
    FrmDialog.Show vbModal
'    frmOption.Show vbModal
End Sub

Private Sub CommandOK_Click()
'If Textdate <> "" Then
On Error GoTo Salah
    If (Val(TextQty.Text) <= 0 Or Val(TextQty.Text) > 32767) Then
        MsgBox "Jumlah label tidak valid"
        TextQty.Text = ""
        TextQty.SetFocus
    Else
        Qty_printed = Val(TextQty.Text)
        CancelBut = False
        'Unload Me
        frmMain.SetFocus
        Call Printing
        MonthView1_DateClick Now
    End If
    If Check3.Value = 1 Then
        txtRef.SetFocus
    End If
    Exit Sub

Salah:
    lblmessagebox.Caption = "Problem : tombol print bermasalah"
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Dim Kondisi As Boolean
Dim Minggu As String
Dim Tahun As String
Dim Bulan As String

    Minggu = Format(DateClicked, "ww", vbMonday, vbFirstFourDays)
    Tahun = Format(DateClicked, "yy", vbMonday, vbFirstFourDays)
    Bulan = Format(DateClicked, "mm", vbMonday, vbFirstFourDays)
    'If tahun <> Right(MonthView1.Year, 2) Then
    '    If Minggu = 1 Then
    '        tahun = tahun
    '    End If
    'End If
    If Minggu = 1 And Bulan = "12" Then
        Tahun = Tahun + 1
    End If
    If Minggu = 52 Or Minggu = 53 Then
        If Bulan = "01" Then
            Tahun = Tahun - 1
        End If
    End If
    If Minggu <> MonthView1.Week Then
        If Minggu = 53 And MonthView1.Week = 1 Then
            Minggu = "01"
            Tahun = Tahun + 1
        End If
    End If
    If Len(Tahun) = 1 Then Tahun = "0" & Tahun
    If Len(Minggu) = 1 Then Minggu = "0" & Minggu
    'Minggu = MonthView1.Week
    'If Len(Minggu) = 1 Then
    '    Minggu = "0" & Minggu
    'End If
    'Textdate.Text = Right(MonthView1.Year, 2) & Minggu
    Textdate.Text = Tahun & Minggu
    'Textdate.Locked = True
    Textdate.Enabled = False
    MonthView1.Visible = False
    Kondisi = Penyegaran
    If Not Kondisi Then
    '            PesanSalah = "Reload tampilan gagal"
    '    GoTo Salah
    End If
    
    'Set Textdate.Text = MonthView1.Week
    'Set Text1.Text = MonthView1.Week
End Sub

Private Sub MonthView1_GotFocus()
    Textdate.Text = Right(MonthView1.Year, 2) & MonthView1.Week
End Sub

Private Sub OptPosisi_Click(Index As Integer)
'Dim Parameter180 As Variant
Dim Kondisi As Boolean
Dim Nilai As Integer
Dim NamaFile As String

'    Kondisi = CS7Server.ActiveDocument.Printer.GetParameter(3, Parameter180)
'    If Not Kondisi Then
'        MsgBox "salah ambil" & vbCrLf & Parameter180
'    Else
'        MsgBox Parameter180
'    End If
    If OptType(1).Value Then
        Nilai = 1
    ElseIf OptType(2).Value Then
        Nilai = 2
    ElseIf OptType(3).Value Then
        Nilai = 3
    ElseIf OptType(4).Value Then
        Nilai = 4
    ElseIf OptType(5).Value Then
        Nilai = 5
    Else
        Nilai = 0
    End If
    NamaFile = ""

    If Index = 1 Then
        NamaFile = Lab(Nilai) & ".lab"
    ElseIf Index = 2 Then
        NamaFile = Lab(Nilai) & "RR.lab"
    End If
    If Not (NamaFile = "") Then
        CS7Server.Documents(NamaFile).Activate
        Set Doc = CS7Server.ActiveDocument
    End If
    If Check3.Value = 0 Then
        Nilai = cboModel.ListIndex
    Else
        Nilai = -1
    End If
    If Nilai > -1 Or Not ModelPilih = "" Then
        Kondisi = Penyegaran
        If Not Kondisi Then
            lblmessagebox.Caption = "Problem dalam penyegaran...."
            Timer2.Enabled = True
            StatusBar1.Panels.Item(2).Text = "Problem"
            Exit Sub
        End If
    End If
    If Check3.Value = 1 Then
        If Not Loading Then
            txtRef.SetFocus
        End If
    End If
    StatusBar1.Panels.Item(2).Text = "OK"
End Sub

Private Sub OptType_Click(Index As Integer)
Dim Nilai As Integer
Dim NamaFile As String
Dim Kondisi As Boolean

On Error GoTo Salah
    PesanSalah = ""
    If OptPosisi(1).Value Then
        Nilai = 1
    ElseIf OptPosisi(2).Value Then
        Nilai = 2
    Else
        Nilai = 0
    End If
    NamaFile = ""
                
    If Nilai = 1 Then
        NamaFile = Lab(Index) & ".Lab"
    ElseIf Nilai = 2 Then
        NamaFile = Lab(Index) & "RR.Lab"
    End If
    If Not (NamaFile = "") Then
        CS7Server.Documents(NamaFile).Activate
        Set Doc = CS7Server.ActiveDocument
    End If
    'Nilai = cboModel.ListIndex
    If Check3.Value = 0 Then
        Nilai = cboModel.ListIndex
    Else
        Nilai = -1
    End If
    If Nilai > -1 Or Not ModelPilih = "" Then
        Kondisi = Penyegaran
        If Not Kondisi Then
'            lblmessagebox.Caption = "Problem dalam penyegaran...."
            PesanSalah = "Problem dalam penyegaran...."
        End If
    End If
    If Check3.Value = 1 Then
        If Not Loading Then
            txtRef.SetFocus
        End If
    End If
    Exit Sub
    
Salah:
    If PesanSalah = "" Then PesanSalah = Err.Description
    lblmessagebox.Caption = PesanSalah
    StatusBar1.Panels.Item(2).Text = "Problem"
    
End Sub

Private Sub Timer1_Timer()
    StatusBar1.Panels.Item(2).Text = Time
If Text2.Text = "1" Then
    CommandCancel_Click
End If
End Sub

Private Sub Timer2_Timer()
Dim Nilai As Long
Dim Kondisi As String
    
    Kondisi = StatusBar1.Panels.Item(2).Text
    If Kondisi = "OK" Then
        Timer2.Enabled = False
        Frame2.BackColor = &HC0C000
    Else
        Nilai = Frame2.BackColor
        If Nilai = &HC0C000 Then
            Frame2.BackColor = &H8080FF
        Else
            Frame2.BackColor = &HC0C000
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Buttonclicked As MSComctlLib.Button)
Select Case Buttonclicked.Key
    Case "CSToggle"
        CSToggle_Click
    Case "CSHelp"
        mnu_About_Click
    Case "CSExit"
        cmdExit_Click
End Select

End Sub

Private Sub cboModel_Click()
Dim Kondisi As Boolean
Dim I As Integer
Dim A As String

On Error GoTo Salah
    frmMain.lblmessagebox.Caption = "Tunggu sedang ambil data modelnya...."
    Picture1.Visible = False
    MonthView1_DateClick Now
    ModelPilih = cboModel.Text
    If ModelPilih = "" Then
        frmMain.lblmessagebox.Caption = "Silahkan Pilih Modelnya...."
    Else
        AmbilLabAktif ServerData, "Label", "Nama` ASC", "Model`='" & ModelPilih & "'", Kondisi
        If Kondisi Then
            For I = 1 To 5
                If Lab(I) <> "" Then
                    OptType(I).Visible = LabAktif(I)
                End If
            Next I
            If ((cboModel.ListIndex + 1) > 0) Then
                TxtDisplayRef.Text = ModelPilih
                Kondisi = Penyegaran
                If Not Kondisi Then
        '            PesanSalah = "Reload tampilan gagal"
                    GoTo Salah
                End If
            End If
        Else
            PesanSalah = "Gagal ambil DataLab Aktif"
            GoTo Salah
        End If
    End If
    StatusBar1.Panels.Item(2).Text = "OK"
    Exit Sub

Salah:
    lblmessagebox.Caption = PesanSalah
    StatusBar1.Panels.Item(2).Text = "Problem"
    Timer2.Enabled = True
'    MsgBox "Problem: (Pilih Model tak jalan) hubungi test enginner.", vbExclamation

End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
Dim Kondisi As Boolean
Dim I As Integer
Dim KodeRef As String
Dim KodeGroup As String

On Error GoTo Salah
    If KeyAscii = 13 Then
        MonthView1_DateClick Now
        ModelPilih = txtRef.Text
        If Len(ModelPilih) > 16 Then
            ModelPilih = Trim(Left(ModelPilih, 15))
        End If
        TxtDisplayRef = ModelPilih
'Begin v5.0.5 ----------------------------------------------------------------------------------------------
' Oleh  : Yohanes Khan
' Tgl   : 20130923
' Alasan    : Problem untuk TQ product dan pembatasan pemilihan agar label hanya muncul yang diperlukan saja
        If Right(ModelPilih, 2) = "TQ" Then
            ModelPilih = Left(ModelPilih, Len(ModelPilih) - 2)
        ElseIf Right(ModelPilih, 3) = "SAX" Then
            ModelPilih = Left(ModelPilih, Len(ModelPilih) - 3)
        End If
'End v5.0.5 ----------------------------------------------------------------------------------------------
        AmbilKolom "Group", ServerData, "Label", "group` asc", "ref`='" & ModelPilih & "'", Kondisi, KodeGroup
        If Kondisi Then
            Kondisi = Bukalab(KodeGroup)
            If Not Kondisi Then
                PesanSalah = "Gagal Buka template"
                GoTo Salah
            End If
        Else
            PesanSalah = "Gagal ambil kode group"
            GoTo Salah
        End If
        AmbilLabAktif ServerData, "Label", "Nama` ASC", "Ref`='" & ModelPilih & "'", Kondisi
        If Kondisi Then
'Begin v5.0.5 ----------------------------------------------------------------------------------------------
' Oleh  : Yohanes Khan
' Tgl   : 20130923
' Alasan    : Problem untuk TQ product dan pembatasan pemilihan agar label hanya muncul yang diperlukan saja
            'For I = 1 To 5
            '    If Lab(I) <> "" Then
            '        OptType(I).Visible = LabAktif(I)
            '    End If
            'Next I
            For I = 1 To 5
                If I = 3 Then
                    OptType(I).Visible = True
                Else
                    OptType(I).Visible = False
                End If
            Next I
            SetFrame Frame8, True
            SetFrame Frame7, True
            'OptType(3).Visible = True
            'OptType_Click (3)
            OptType(3).Value = True
'End v5.0.5 ----------------------------------------------------------------------------------------------
            'KodeRef = Left(ModelPilih, 2)
            'If KodeRef = "XU" Then
            '    OptType_Click (3)
            'ElseIf KodeRef = "SA" Then
            '    OptType_Click (4)
            'Else
            '    OptType_Click (1)
            'End If
            Kondisi = Penyegaran
            If Not Kondisi Then
                GoTo Salah
            End If
        Else
            PesanSalah = "Gagal ambil DataLab Aktif"
            GoTo Salah
        End If
        txtRef.Text = ""
    End If
    Exit Sub
Salah:
    SetFrame Frame8, False
    SetFrame Frame7, False
    lblmessagebox.Caption = PesanSalah
    StatusBar1.Panels.Item(2).Text = "Problem"
    Timer2.Enabled = True
    'MsgBox "Kesalahan dalam mengambil data" & vbCrLf & PesanSalah, vbExclamation, "Warning"

End Sub

Public Function Logdata(FileLok$, Pesan As String)
Dim MyMsg, ID$, TestTime
    
On Error GoTo ErrorWrite
Open FileLok For Append As #1
Close #1
    
MyMsg = Day(Date) & "/" & Month(Date) & "/" & Year(Date)

Open FileLok$ For Append As #1
Print #1, Pesan
Close #1

Exit Function

ErrorWrite:
    If Err.Number = 53 Then
        On Error Resume Next
        Open FileLok$ For Append As #1
        Print #1, Pesan
        Close #1
        Logdata FileLok$, Pesan
    Else
        MsgBox "Please Close Datalog file!!" & vbCrLf & Err.Number & ":" & Err.Description, vbCritical, "Peringatan"
    End If
End Function


