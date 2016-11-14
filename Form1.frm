VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FF8080&
   ForeColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
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
      Height          =   1215
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8880
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Label information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3975
      Left            =   2400
      TabIndex        =   28
      Top             =   1680
      Width           =   10575
      Begin VB.CommandButton Cancel 
         Caption         =   "Command2"
         Height          =   375
         Left            =   9000
         TabIndex        =   41
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton OK 
         BackColor       =   &H008080FF&
         Caption         =   "OK"
         Height          =   375
         Left            =   9000
         TabIndex        =   40
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox TextDate 
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H0000FF00&
         Height          =   735
         Left            =   6120
         TabIndex        =   38
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox TextQty 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   6000
         TabIndex        =   36
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtModel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   945
         HideSelection   =   0   'False
         Left            =   480
         TabIndex        =   30
         Text            =   "1001000100011"
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox TxtDisplaymodel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   960
         HideSelection   =   0   'False
         Left            =   240
         TabIndex        =   29
         Text            =   "011167"
         Top             =   2640
         Width           =   5295
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Left            =   6480
         TabIndex        =   39
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
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
         Left            =   6240
         TabIndex        =   37
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Left            =   1080
         TabIndex        =   32
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number"
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
         TabIndex        =   31
         Top             =   2160
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00C0C0FF&
      Height          =   2535
      Left            =   2400
      TabIndex        =   26
      Top             =   5880
      Width           =   10575
      Begin VB.Label lblmessagebox 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Please scan product to print label..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1455
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   10215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3960
      Top             =   8880
   End
   Begin VB.CommandButton cmdStorePass 
      Caption         =   "StorePass"
      Height          =   615
      Left            =   4440
      TabIndex        =   25
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdShowLabelPage 
      Caption         =   "Show Label Page"
      Height          =   735
      Left            =   480
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdGetLabel 
      Caption         =   "Get Label "
      Height          =   735
      Left            =   480
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CS toggle"
      Height          =   615
      Left            =   480
      TabIndex        =   22
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   4815
      Left            =   0
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtEnglish 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Text            =   "English"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtFrance 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Text            =   "France"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtGerman 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Text            =   "German"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtVoltage 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Text            =   "Voltage"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtReference 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Text            =   "Reference"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtArticleNumber 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Text            =   "ArticleNumber"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtBarcode 
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Text            =   "Barcode"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Text            =   "Quantity"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtLabelSize 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Text            =   "LabelSize"
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtMaterialNumber 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Text            =   "MaterialNumber"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txtType 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Text            =   "Type for Schile only"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox TxtCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Text            =   "Current"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TxtLoadpower 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Text            =   "LoadPower"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox TxtPower 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Text            =   "Power"
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox TxtBitmap 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Text            =   "Bitmap"
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox TxtSpain 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "Spain"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox TxtIta 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Text            =   "Italiano"
         Top             =   2040
         Width           =   2055
      End
   End
   Begin VB.TextBox Date 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Date"
      Top             =   720
      Width           =   1335
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   375
      Left            =   11160
      TabIndex        =   0
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16744576
      FullWidth       =   33
      FullHeight      =   25
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2715
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BB34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CSToggle"
            Object.ToolTipText     =   "CS Server"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CSHelp"
            Object.ToolTipText     =   "About this program"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CSExit"
            Object.ToolTipText     =   "Exit this program"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSoftwarever 
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10560
      TabIndex        =   35
      Top             =   10080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Maximo Name Plate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   3000
      TabIndex        =   34
      Top             =   720
      Width           =   9375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstRef As Recordset
Option Explicit

Private Sub cmdExit_Click()
If CS7Server_Stop Then
    End
Else
    MsgBox "Code Soft server did not return the value requested.", vbExclamation
End If

End Sub

Private Sub GetLabel()
Dim DBHis
Dim i As Integer
Dim rs As ADODB.Recordset
Dim a As String

    If (txtModel.Text <> "") Then
        Set DBHis = New ADODB.Connection
        DBHis.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & _
        "data source=" & "Maximo.mdb" & ";Jet OLEDB:Database Password = plutonium;"
        DBHis.Open
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM Etiquette", DBHis, adOpenKeyset, adLockOptimistic
        With rs
        rs.MoveFirst
        If .EOF = True Then 'if nothing on first row of database
            lblmessagebox.Caption = "No data found in the database"
            lblmessagebox.BackColor = vbRed
            .Close
            Exit Sub
        Else
            While (.EOF = False)
                If (.Fields("REFERENCE_COMMERCIALE") = ReadScan) Or (.Fields("REFERENCE_COMMERCIALE_IMPRIMER") = ReadScan) Then
                    If IsNull(.Fields("Numero_d_article")) Then
                        frmMain.txtArticleNumber.Text = Format(.Fields("Numero_d_article"), "XXXXXXX")
                    Else
                        frmMain.txtArticleNumber.Text = .Fields("Numero_d_article")
                        TxtDisplaymodel.Text = .Fields("Numero_d_article")
                    End If
                    
                    If IsNull(.Fields("PCX9")) Then
                        frmMain.TxtBitmap.Text = ""
                    Else
                        frmMain.TxtBitmap.Text = .Fields("PCX9")
                    End If

                    If IsNull(.Fields("Made_in")) Then
                        frmMain.TxtCurrent.Text = ""
                    Else
                        frmMain.TxtCurrent.Text = .Fields("Made_in")
                    End If

                    If IsNull(.Fields("Caracteristique_5")) Then
                        frmMain.txtEnglish.Text = ""
                    Else
                        frmMain.txtEnglish.Text = .Fields("Caracteristique_5")
                    End If

                    If IsNull(.Fields("Caracteristique_4")) Then
                        frmMain.txtFrance.Text = ""
                    Else
                        frmMain.txtFrance.Text = .Fields("Caracteristique_4")
                    End If

                    If IsNull(.Fields("Caracteristique_7")) Then
                        frmMain.txtGerman.Text = ""
                    Else
                        frmMain.txtGerman.Text = .Fields("Caracteristique_7")
                    End If

                    If IsNull(.Fields("Caracteristique_8")) Then
                        frmMain.TxtSpain.Text = ""
                    Else
                        frmMain.TxtSpain.Text = .Fields("Caracteristique_8")
                    End If
                    .Close
                    lblmessagebox.Caption = "Label found."
'MsgBox "Label fould"
                    Call cmdPrint
                    Exit Sub
                Else
                    .MoveNext
                    If .EOF = True Then 'if nothing on first row of database
                        lblmessagebox.Caption = "Article number is not valid for Maximo products"
                        lblmessagebox.BackColor = vbRed
                        .Close
                        Exit Sub
                    End If
                End If
            Wend
        End If
        End With
    End If
Exit Sub 'necessary if not it will always go into Errorhandle routine
        
ErrorHandle:
        MsgBox "Missing file PrtLabels.mdb", vbExclamation, "Warning"
        End

End Sub

Private Sub cmdPrint_Click()
    Call frmPackaging.cmdPrint_Click

End Sub

Private Sub cmdShowLabelPage_Click()
    frmPackaging.Show

End Sub

Private Sub Command1_Click()
    IsVisible = Not IsVisible
    ServerVisible IsVisible

End Sub

Private Sub CSExit_Click()
    cmdExit_Click
    
End Sub

Private Sub CSToggle_Click()
    Command1_Click
    
End Sub

Private Sub Form_Activate()
Dim i, pnlx
    txtModel.SetFocus

End Sub

Private Sub Form_Load()
    On Error GoTo Error_msg
    txtModel.Text = ""
    lblmessagebox.Caption = "Scan-lah barcode untuk print label..."
    lblmessagebox.BackColor = &HFF8080
    
Dim ErrAnzeige As String
Dim LastErr&, i, pnlx, a
For i = 1 To 3
        Set pnlx = StatusBar1.Panels.Add() ' Add 2 panels.
Next i
With StatusBar1.Panels
    .Item(1).AutoSize = sbrSpring
    .Item(1).Text = "AP Inc." & " <" & Format(Date, "dd mmmm yyyy") & "> "
    .Item(3).Style = sbrNum                             ' NumLock
    .Item(4).Style = sbrCaps
End With
LastErr = OpenLab(CurDir & "\Maximo_Name_Plate.lab")
Exit Sub

Error_msg:
a = MsgBox("Fatal error dan tidak dapat melanjutkan.", vbExclamation)
End

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call CS7Server_Stop
End

End Sub

Private Sub Mnu_Help_Click()
frmAbout.Show 1

End Sub

Private Sub mnuTester_ver_update_Click()
On Error Resume Next
Update_soft_flag = True
frmSecurity.txtpasswd.Text = ""
frmSecurity.Height = 3135
frmSecurity.Show 1
frmSecurity.txtpasswd.SetFocus

End Sub

Private Sub Text1_Change()
te
End Sub

Private Sub TextDate_Change()
TextDate.Text = Date
End Sub

Private Sub Timer1_Timer()
    StatusBar1.Panels.Item(2).Text = Time

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Buttonclicked As MSComctlLib.Button)
Select Case Buttonclicked.Key
    Case "CSToggle"
        Command1_Click
    Case "CSHelp"
        Mnu_Help_Click
    Case "CSExit"
        cmdExit_Click
End Select

End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
Dim j As Integer
Dim hello As String
Dim Print_Label_flag As Boolean
Print_Label_flag = False

    If (KeyAscii = 13) Then
        SendKeys "{Home}", True
        SendKeys "+{End}", True
        If txtModel.Text = "" Then Exit Sub
        lblmessagebox.ForeColor = &H80FFFF
        ReadScan = UCase(Left(txtModel.Text, 12)) 'convert to upper case
        'MsgBox ReadScan
        Call GetLabel
        Exit Sub
    End If

End Sub


