VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Driving my labeling software with ActiveX"
   ClientHeight    =   6120
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Opening Document"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   240
      TabIndex        =   26
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Dialog.frx":0000
         Left            =   1800
         List            =   "Dialog.frx":0002
         TabIndex        =   39
         ToolTipText     =   "Pilih Tipe Label yang akan di Print"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   38
         ToolTipText     =   "Isi Dengan DateCode (YYWW)"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   36
         ToolTipText     =   "Isi dengan Jumlah dalam Box"
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Preview"
         Height          =   375
         Left            =   4080
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox Daftar 
         Height          =   315
         ItemData        =   "Dialog.frx":0004
         Left            =   1800
         List            =   "Dialog.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   27
         ToolTipText     =   "Pilih Tipe Produknya"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Artikel No:"
         Height          =   255
         Left            =   1920
         TabIndex        =   42
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Label Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "DateCode:"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Qty:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1800
         TabIndex        =   34
         Top             =   720
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Tipe Produk:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Artikel No:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.PictureBox imgPreview 
      BackColor       =   &H80000009&
      Height          =   3375
      Left            =   5760
      ScaleHeight     =   3315
      ScaleWidth      =   3435
      TabIndex        =   25
      Top             =   120
      Width           =   3495
   End
   Begin VB.Data DataLabel 
      Connect         =   "Access"
      DatabaseName    =   "C:\Label\datalabel.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LABEL"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7200
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   7440
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   "Database browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdPreviousRecord 
         Height          =   375
         Left            =   960
         Picture         =   "Dialog.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Preview2 
         Caption         =   "Preview"
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdLastRecord 
         Default         =   -1  'True
         Height          =   375
         Left            =   2640
         Picture         =   "Dialog.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdNextRecord 
         Height          =   375
         Left            =   1800
         Picture         =   "Dialog.frx":0A34
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdFirstRecord 
         Height          =   375
         Left            =   120
         Picture         =   "Dialog.frx":0F12
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Access to the designer module"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   4935
      Begin VB.OptionButton OptLppx 
         Caption         =   "Visible"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptNoControl 
         Caption         =   "No visible"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "&Merge"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      ToolTipText     =   "Tekan Tombol ini untuk mencetak label"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Printing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   3375
      Begin VB.CommandButton cmdAddPrinters 
         Caption         =   "&Add Printer "
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtQuantity 
         Height          =   375
         Left            =   960
         TabIndex        =   10
         ToolTipText     =   "Isi dengan jumlah label yang akan diprint"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Jumlah:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Displaying Variables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3840
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
      Begin VB.CommandButton cmdFiller 
         Caption         =   "&Form"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ListBox lstVariables 
         Height          =   2400
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdVariables 
         Caption         =   "&Variables "
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opening Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8760
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmdCloseQuery 
         Caption         =   "&Close Query "
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtQueryname 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
      Begin VB.CommandButton cmdOpenQuery 
         Caption         =   "Open &Query"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyApp As LabelManager2.Application
Dim MyDoc As LabelManager2.Document
Dim MyVars As LabelManager2.Variables
Dim DB As Database
Dim RS As Recordset
Dim DateCode As String

Private Sub HitungWeek()
    Dim Week As Integer
    Week = Format(Date, "ww", vbMonday, vbFirstFullWeek)
    DateCode = Format(Date, "YY", vbMonday, vbFirstFullWeek) & Week
End Sub
Private Sub Combo1_Click()
    Daftar_Click
End Sub

Private Sub Command1_Click()
    Dim var As LabelManager2.Variable
    lstVariables.Clear
    Set MyVars = MyDoc.Variables
        For Each var In MyVars
            lstVariables.AddItem var.Name
        Next
    Label1.Caption = MyDoc.Variables.Item(2).Value
  
End Sub

Private Sub Daftar_Click()
    Dim StringBuf
'    On Error Resume Next
    If Daftar.Text <> "" Then
    StringBuf = "Select * FROM LABEL WHERE REF = '" & Daftar.Text & "'"
'    Set RS = DataLabel.OpenRecordset(StringBuf, dbOpenDynaset)
    DataLabel.Recordset.FindFirst ("REF = '" & Daftar.Text & "'")
    Set RS = DataLabel.Recordset
    Label6.Caption = RS!Article
    Text1.Text = RS!Qty
    End If
'    Text2.Text = Mid(CStr(Year(Date)), 3, 2) & CStr(Day(Date))
' Ini untuk mengisi datecode (Perlu dicari cara agar menampilkan week skr yang benar)
    HitungWeek
    Text2.Text = DateCode
    
'Kode dibawah untuk membuka label antara individual dan packing
    If StringBuf <> "" Then
    If Combo1.ListIndex = 1 Then
'    Set MyDoc = MyApp.Documents.Open(StringBuf) Pemanggilan Manual
        StringBuf = "C:\Label\" & RS!Lab_Template
        BukaFile (StringBuf) 'Untuk Ambil Nilai Variablesnya
' Ini untuk mengubah nilai di Label (Reference, Article, Barcode, Qty)
'    MyDoc.DocObjects.Texts.Item(4).Value = "1"
'    MyDoc.DocObjects.Texts.Item(5).Value = "123456"
'    MyDoc.DocObjects.Texts.Item(1).AppendString "123345", 1
'    MyDoc.DocObjects.Texts.Item(1).Value = Label6.Caption
'    MyDoc.DocObjects.Texts.Item(1).Value = RS!Qty
'    MyDoc.DocObjects.Texts.Item(6).Value = RS!Qty
'    MyDoc.DocObjects.Texts.Item(7).Value = "0551"
        cmdVariables_Click
        MyDoc.Variables.FormVariables.Item(1).Value = RS!Article
        StringBuf = RS!REF
        MyDoc.Variables.FormVariables.Item(3).Value = StringBuf
        MyDoc.Variables.FormVariables.Item(4).Value = Text2.Text
        MyDoc.Variables.FormVariables.Item(2).Value = "Sn=1.5mm" 'Perlu ditambah di Databasenya
        MyDoc.DocObjects.Barcodes.Item(1).Value = "338911" & RS!Article
    End If
    If Combo1.ListIndex = 0 Then
        StringBuf = "C:\Label\" & RS!Pack_Template
        BukaFile (StringBuf)
        cmdVariables_Click  'Untuk Ambil Nilai Variablesnya
        MyDoc.Variables.FormVariables.Item(1).Value = RS!Article
        MyDoc.Variables.FormVariables.Item(2).Value = Text1.Text
        StringBuf = RS!REF
        MyDoc.Variables.FormVariables.Item(3).Value = StringBuf
        MyDoc.Variables.FormVariables.Item(4).Value = Text2.Text
        MyDoc.DocObjects.Barcodes.Item(1).Value = "338911" & RS!Article
    End If
        Label10.Caption = MyDoc.Variables.FormVariables.Item(1).FormOrder
    Tampil 'Ini untuk menampilkan preview labelnya
    End If
End Sub

Private Sub Daftar_KeyPress(KeyAscii As Integer)
    Label1.Caption = CStr(KeyAscii)
End Sub

Private Sub Form_Activate()
DataLabel.Recordset.MoveFirst
On Error GoTo AdaKesalahan
    DataLabel.Caption = "Data ke " + Str(DataLabel.Recordset.RecordCount * DataLabel.Recordset.PercentPosition / 100 + 1)
    MsgBox "Data berhasil terhubung", vbOKOnly, "Pesan"
    Exit Sub
AdaKesalahan:
MsgBox "Data belum terhubung. Panggil Prod Spec!", vbOKOnly, "Pesan"
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Set MyApp = New LabelManager2.Application   'Memanggil Label Manager CodeSoft
    MyApp.Visible = False
'    DATABASEDir = "C:\Yo\"
'    Set DB = OpenDatabase("C:\Yo\DataLabel.MDB")
    
    AmbilTipe   'Ini untuk memasukkan tipe produk ke dalam combobox
    Label6.Caption = ""
    Combo1.AddItem "Packing"
    Combo1.AddItem "Individual"
    Combo1.ListIndex = 1
    
    Me.MousePointer = vbDefault
End Sub
Private Sub AmbilTipe()
    Dim StringBuf
    DataLabel.Refresh
    While Not DataLabel.Recordset.EOF
        StringBuf = DataLabel.Recordset.Fields("REF")
        Daftar.AddItem StringBuf
        DataLabel.Recordset.MoveNext
    Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadLppx
    Unload frmPrvw
End Sub
Private Sub DataLabel_Reposition()
Dim angka As Integer
'    angka = 1
'    DataLabel.Caption = "Data ke " + Str(DataLabel.Recordset.RecordCount * DataLabel.Recordset.PercentPosition / 100 + 1)
'    TxtRef(0).Text = DataLabel.Recordset.Fields("Ref")
'    Label1.Caption = Str(DataLabel.Recordset.PercentPosition)
End Sub

' *********** BUTTONS MANAGMENT *********
Private Sub BukaFile(Nama As String)
        Me.MousePointer = vbHourglass
        MyApp.Documents.CloseAll (False)
        Set MyDoc = MyApp.Documents.Open(Nama)
        Me.MousePointer = vbDefault
End Sub
'Private Sub cmdOpen_Click()
'    If SelectFile("lab (*.lab)|*.lab", 1) Then
'        txtFilename = CommonDialog1.FileName
'        Me.MousePointer = vbHourglass
'        Set MyDoc = MyApp.Documents.Open(txtFilename)
'        Me.MousePointer = vbDefault
'    End If
'End Sub

'Private Sub cmdSave_Click()
'    If SelectFile("lab (*.lab)|*.lab", 2) Then
'        txtFilename(1) = CommonDialog1.FileName
'        MyDoc.SaveAs txtFilename
'    End If
'End Sub

Private Sub cmdOpenQuery_Click()
    If SelectFile("csq (*.csq)|*.csq", 1) Then
        txtQueryname = CommonDialog1.FileName
        Me.MousePointer = vbHourglass
        MyDoc.Database.OpenQuery txtQueryname
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdCloseQuery_Click()
    MyDoc.Database.Close
    txtQueryname.Text = ""
    lstVariables.Clear
End Sub

Private Sub cmdPreview_Click()
    Preview
End Sub

Private Sub cmdVariables_Click()
    Dim var As LabelManager2.Variable
    lstVariables.Clear
    Set MyVars = MyDoc.Variables
        For Each var In MyVars
            lstVariables.AddItem var.Name
        Next
End Sub

Private Sub cmdFiller_Click()
    MyApp.Dialogs(lppxFormDialog).Show
End Sub

Private Sub cmdAddPrinters_Click()
    MyApp.Dialogs(lppxPrinterSelectDialog).Show
End Sub

Private Sub cmdPrint_Click()
    MyDoc.PrintLabel txtQuantity
    MyDoc.FormFeed
End Sub

Private Sub cmdMerge_Click()
    MyDoc.Merge 1
End Sub

Private Sub OptNoControl_Click()
    MyApp.Visible = False
End Sub

Private Sub OptLppx_Click()
    MyApp.Visible = True
    'sample of ActiveX error management:
    Dim Errornum As Long, ErrorMsg As String
    Errornum = MyApp.GetLastError
    ErrorMsg = MyApp.ErrorMessage(Errornum)
        If Errornum <> 0 Then
            MsgBox ErrorMsg, vbCritical, "Error #" & Errornum
        End If
    'sample end here
End Sub

Private Sub cmdFirstRecord_Click()
    MyDoc.Database.MoveFirst
    If frmPrvw.Visible Then Preview
End Sub

Private Sub cmdPreviousRecord_Click()
    MyDoc.Database.MovePrevious
    If frmPrvw.Visible Then Preview
End Sub

Private Sub cmdNextRecord_Click()
    MyDoc.Database.MoveNext
    If frmPrvw.Visible Then Preview
End Sub

Private Sub cmdLastRecord_Click()
    MyDoc.Database.MoveLast
    If frmPrvw.Visible Then Preview
End Sub

Private Sub Preview2_Click()
    Preview
End Sub

'**** Procedure & Function ****
Sub Preview()
    MyDoc.CopyToClipboard
    Load frmPrvw
    frmPrvw.imgPreview.Picture = Clipboard.GetData(vbCFMetafile)
    frmPrvw.Show
End Sub
Sub Tampil()
    MyDoc.CopyToClipboard
'    Load frmPrvw
    imgPreview.Picture = Clipboard.GetData(vbCFMetafile)
'    frmPrvw.Show
End Sub

Private Function SelectFile(ByVal filters As String, ByVal actioncode As Long) As Boolean
    Dim s As String
    
    s = Mid(filters, InStrRev(filters, "|") + 1)
    
    On Error Resume Next
    With CommonDialog1
        .Filter = filters
        .FileName = s
        .Action = actioncode
        'Dialog box is displaying
        '...
        'Dialog box is closed
        If Err = 0 Then
            SelectFile = True
        End If
    End With
End Function
        
Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub UnloadLppx()
    MyApp.Documents.CloseAll False
    MyApp.Quit
    Set MyApp = Nothing
End Sub

