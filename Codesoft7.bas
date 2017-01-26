Attribute VB_Name = "Codesoft7"
'*****************************************************
'*
'* OLE CODESOFT APPLICATION
'* AP Inc. 2006
'*
'*****************************************************
Option Explicit
'Option Base 1

' Variabel untuk Parameter software
Public ServerAlamat As String
Public ServerData As String
Public ServerUser As String
Public ServerPass As String
Public ServerDriver As String
Public ServerOption As String

Public PathData As String
Public PathLabel As String
Public PathGambar As String
Public PathTemp As String
Public NamaLabel As String

Public KodePowerSupply As String
Public KodeMultiMeter As String
Public AlamatMultiMeter As String
Public AlamatPowerSupply As String
Public Keluarga As String
Public PosisiAtas As String
Public PosisiKiri As String
Public SatuKeluarga As Boolean

Public PassAdmin As String
Public PassOperator As String

Public LabAktif(5) As Boolean
Public Lab(5) As String

' Variable untuk Error Trapping
Public PesanSalah As String
Public SalahOption As Boolean
Public ProsesSts(4) As Boolean

' Variable untuk Codesoft
Public IsVisible As Boolean
Public OrgXachse As Double
Public OrgYachse As Double
Public CSStatus As String
Public CS7Server As LabelManager2.Application
'Public CS7Server As Lppx2.Application
Public Doc As LabelManager2.Document
Public Vars As LabelManager2.Variables
Public CSVar As LabelManager2.Variable
Public CStexts As LabelManager2.Texts
Public CSText As LabelManager2.Text
Public CSImages As LabelManager2.Images
Public CSImage As LabelManager2.Image

' Variable untuk Database
'Public DB As Database
Public SQL As String
Public StringKoneksi As String
Public Conn As ADODB.Connection
Public DataRS As ADODB.Recordset
'Public Fld As ADODB.Field
'Public RS As Recordset

'Variable untuk Printing
Public Qty_printed As Integer
Public ModelPilih As String
Public scan As Integer

'Variable lainnya
Public CancelBut As Boolean
'Public HOME As String
'Public DscFile As String
'Public DscStandard As String
'Public DscOption As Boolean
'Public Errtext As Boolean
'Public DscFeldAnz As Integer
'Public DatenFeldAnz As Integer
'Public DscDelimiter As String
'Public LabFelder() As String
'Public ValFelder() As String
'Public LenFelder() As Integer
'Public DscFelder() As String
'Public Dateifelder() As String
'Public LabFeldAnz As Integer
'Public Delimiter As String
'Public AktLabel As String
'Public ole_server
'Public EtikettDir As String
'Public DatenDir As String
'Public DateiDir As String
'Public Typendir As String
'Public VerpackDir As String
'Public DatenOldDir As String
'Public XAchse As Double
'Public YAchse As Double
'Public CsErr As Variant

'Public Sti_flag As Boolean
'Public OKbut As Boolean
'Public dbsSchneider As Database, dbsBatam As Database
'Public Update_soft_flag As Boolean
'Public Add_Reference_Flag As Boolean
'Public Manual_Print_Flag As Boolean
'Public Previous_Temp_Label As String
'Public ReadScan As String
'Public Wrong_Label As Boolean

'Const DscLabMistmatch = "Die Anzahl der Felder in der DSC-Datei stimmt nicht mit der Anzahl im Label überein!"
'Const DscDateiMistmatch = "Die Anzahl der Felder in der DSC-Datei stimmt nicht mit der Anzahl in der Datei überein!"
'Const LabDateiMistmatch = "Die Anzahl der Felder im Label stimmt nicht mit der Anzahl in der Datei überein!"


Public ActiveLablePath As String

Sub Main()
Dim Kondisi As Boolean

On Error GoTo Salah
    ProsesSts(0) = True
    frmSplash.Show
    frmSplash.Refresh

    frmSplash.Tampilan "-- Cek program"
    Kondisi = App.PrevInstance
    If Kondisi Then
        ProsesSts(1) = False
        PesanSalah = "Program sudah pernah jalan"
        frmSplash.Tampilan "-> NOK"
        frmSplash.Tampilan "xx " & PesanSalah
    Else
        ProsesSts(1) = True
        frmSplash.Tampilan "-> OK"
    End If
    frmSplash.UbahProgress 20
    
    frmSplash.Tampilan "-- Ambil parameter software"
    Kondisi = SetingParameter
    If Not Kondisi Then
        ProsesSts(2) = False
        frmSplash.Tampilan "-> NOK"
        frmSplash.Tampilan "xx " & PesanSalah
    Else
        ProsesSts(2) = True
        frmSplash.Tampilan "-> OK"
    End If
    frmSplash.UbahProgress 40
    
    frmSplash.Tampilan "-- Seting & Cek koneksi ke database"
    Kondisi = BukaKoneksi
    If Not Kondisi Then
        ProsesSts(3) = False
        frmSplash.Tampilan "-> NOK"
        frmSplash.Tampilan "xx " & PesanSalah
        'GoTo Salah
    Else
        ProsesSts(3) = True
        frmSplash.Tampilan "-> OK"
    End If
    frmSplash.UbahProgress 60
    
    frmSplash.Tampilan "-- Seting & Check Codesoft"
    Kondisi = CS7Server_Start
    If Not Kondisi Then
        ProsesSts(4) = False
        frmSplash.Tampilan "-> NOK"
        frmSplash.Tampilan "xx " & PesanSalah
        'GoTo Salah ' try to start OLE Server
    Else
        ProsesSts(4) = True
        frmSplash.Tampilan "-> OK"
    End If
    frmSplash.UbahProgress 80
    
    
    frmSplash.Tampilan "-- Persiapan akhir"
    ProsesSts(0) = ProsesSts(1) And ProsesSts(2) And ProsesSts(3) And ProsesSts(4)
    Load frmMain
    frmSplash.UbahProgress 100
    If ProsesSts(0) Then
        Unload frmSplash
        frmMain.Show
        If frmMain.Check3.Value = 0 Then
            frmMain.txtRef.SetFocus
        End If
    Else
        frmSplash.Tampilan "xx Proses ada masalah, lihat bagian atas"
        Delay (5)
    End If
    Exit Sub
    
Salah:
    MsgBox "Problem: (Sub Main)" & vbCrLf & vbCrLf & Err.Number & " : " & Err.Description, vbExclamation
    Unload frmSplash
    Call TutupKoneksi
    If Not CSStatus = "Tidak Jalan" Then
        Call CS7Server_Stop
    End If
    End
End Sub

Function Delay(num As Single)  '1 = 1 s
Dim start As Double
    start = Timer
    Do While Timer < start + num
        DoEvents
    Loop
End Function

Public Function CS7Server_Start() As Boolean
'Dim ErrAnzeige As String
Dim LastErr&
'Dim a As Integer

On Error Resume Next ' catch errors
    Set CS7Server = New LabelManager2.Application     'implements object
'    ServerVisible True
    LastErr = Err ' store resulting error code
    On Error GoTo 0 ' returns to normal error trapping
    Select Case LastErr ' depending on error code...
        Case 0 ' no error, return true
'            PesanSalah = "OKE - CS7 Server Starting"
            CS7Server.Documents.CloseAll False
            If CS7Server.IsEval Then
                CSStatus = "Jalan Demo"
            Else
                CSStatus = "Jalan Full"
            End If
            CS7Server_Start = True
            Exit Function
        Case 429 ' OLE common error, display special message
            CSStatus = "Tidak Jalan"
            PesanSalah = "CodeSoft OLE server tidak ditemukan, periksa registrasi."
        Case Else ' for other errors, use VB error processing
'            Err.Raise LastErr
            CSStatus = "Tidak Jalan"
            PesanSalah = "CodeSoft Problem No : " & LastErr
    End Select
    CS7Server_Start = False
End Function

Public Function CS7Server_Stop() As Boolean

On Error Resume Next
    CS7Server.Documents.CloseAll False ' Tutup semua document yang terbuka
    CS7Server.Quit ' Panggil perintah keluar
    Set Doc = Nothing ' Bebaskan memory yang terpakai
    Set CS7Server = Nothing ' free automation interface and close OLE server
    CS7Server_Stop = True
End Function

Public Sub ServerVisible(NewState As Boolean)

On Error Resume Next
    CS7Server.Visible = NewState ' display/hide OLE server user interface
End Sub

Public Function GetPreview(MaxDisplay As Control, Display As Control) As Boolean

On Error Resume Next
    Dim Dw&, Dh& ' width and height of control used for display
    Dim Mw&, Mh& ' internal width and height of bounding control
    Dim FitFactor! ' zoom value to fit display in its contener
    Dim Foto As String
    
    ' ask server to draw the label as a metafile and copy it to the clipboard
    
    If CS7Server.Documents.Count < 1 Then
        Exit Function
    End If
     
    CS7Server.ActiveDocument.ViewMode = lppxViewModeValue
    CS7Server.ActiveDocument.CopyToClipboard
    'Foto = CS7Server.ActiveDocument.CopyImageToFile(8, "BMP", 0, 100, App.Path & "\aktlabel.bmp")
    Display.Visible = False ' hide picture during update
    Display.AutoSize = True ' picture control will resize itself to the image
    Display.Picture = Clipboard.GetData(vbCFMetafile) ' get picture from clipboard
    DoEvents ' process any pending messages ( like resize,paint, ... )
    Display.AutoSize = False ' swith off auto resizing
    'Display.Picture = LoadPicture(App.Path & "\aktlabel.bmp")
    
    If Display.Picture.Height > Display.Picture.Width Then
        FitFactor = Display.Picture.Height / Display.Picture.Width
        'Display.Height = MaxDisplay.Height
        Display.Width = Display.Height / FitFactor
    Else
        FitFactor = Display.Picture.Width / Display.Picture.Height
        'Display.Width = MaxDisplay.Width
        Display.Height = Display.Width / FitFactor
    End If
    
    Dw = Display.Width
    Dh = Display.Height
    Mw = MaxDisplay.ScaleWidth
    Mh = MaxDisplay.ScaleHeight
    FitFactor = Dw / Mw ' get width ratio
    If Dh / Mh > FitFactor Then FitFactor = Dh / Mh ' use height ratio if it's greater
    'If FitFactor < 1 Then FitFactor = 1 ' don't use ratio smaller than 1 ( disable enlarging )
    If Not FitFactor > 0 Then FitFactor = 1 ' don't use ratio smaller than 1 ( disable enlarging )
    
    Dw = Dw / FitFactor ' recalculate size of picture to make sure it fits its contener
    Dh = Dh / FitFactor
    
    ' stretch and center picture
    Display.Move (Mw - Dw) / 2, (Mh - Dh) / 2, Dw, Dh

    Display.Visible = True ' show picture
    GetPreview = True
End Function

Public Function Bukalab(Group As String) As Boolean
Dim Kondisi As Boolean
Dim I As Integer
Dim NamaFile As String
Dim Pesan As String

On Error GoTo Salah
    AmbilLab ServerData, "family", "Nama` ASC", "Nama`='" & Group & "'", Kondisi
    If Not Kondisi Then
        Pesan = "Gagal ambil Lab" & vbCrLf & PesanSalah
        GoTo Salah
    End If
    For I = 1 To 5
        If Lab(I) <> "" Then
            NamaFile = PathLabel & Lab(I) & ".lab"
            Kondisi = True
            If I = 3 Then
                Kondisi = OpenLab(NamaFile)
                ActiveLablePath = NamaFile
            End If
            If Not Kondisi Then
                Pesan = "Buka Label 0" & I & " gagal"
                GoTo Salah
            End If
            ''NamaFile = PathLabel & Lab(I) & "RR.lab"
            ''Kondisi = OpenLab(NamaFile)
            ''If Not Kondisi Then
            ''    Pesan = "Buka Label 0" & I & "RR gagal"
            ''    GoTo Salah
            ''End If
            frmMain.OptType(I).Visible = True
        Else
            frmMain.OptType(I).Visible = False
        End If
    Next I

    Bukalab = True
    Exit Function
Salah:
    MsgBox "Salah dalam membuka template" & vbCrLf & Pesan, vbInformation, "PROBLEM"
    Bukalab = False
End Function

Public Function OpenLab(ByVal LabFile$) As Boolean
'    On Error Resume Next
On Error GoTo Salah
    Dim LastErr&
    Dim ErrAnzeige As String
    
    'OpenLab = CS7Server.Documents.Open(LabFile)
    Set Doc = CS7Server.Documents.Open(LabFile, True)
    'Open_doc = True
    OpenLab = True
    OrgXachse = CS7Server.ActiveDocument.Format.MarginLeft
    OrgYachse = CS7Server.ActiveDocument.Format.MarginTop
    Exit Function

Salah:
    LastErr = CS7Server.Application.GetLastError
    If LastErr > 0 Then
        ErrAnzeige = CS7Server.Application.ErrorMessage(LastErr)
        MsgBox ErrAnzeige & vbCrLf & "(Buka file template gagal)", vbCritical
        OpenLab = False
    Else
        OpenLab = True
    End If
    
End Function

Public Function PrintLab(ByVal num%, ByVal Serie%) As Boolean
    On Error Resume Next
    Dim I As Integer
    Dim Cut As Long
   
    'XAchse = GetSetting(App.Title, "Settings", Trim(CS7Server.ActiveDocument.Name) & "X", 0)
    'YAchse = GetSetting(App.Title, "Settings", Trim(CS7Server.ActiveDocument.Name) & "Y", 0)
    'CS7Server.ActiveDocument.Format.MarginTop = OrgYachse + (YAchse * 100)
    'CS7Server.ActiveDocument.Format.MarginLeft = OrgXachse + (XAchse * 100)
    
  If PrinterSelect("Zebra 105SL (300 dpi) (Copy 2)", "COM1:") = True Then
'    If PrinterSelect("Zebra TLP2844-Z", "LPT1:") = True Then
    PrintLab = CS7Server.ActiveDocument.PrintLabel(num, Serie)
    CS7Server.ActiveDocument.FormFeed
   Else
    MsgBox "System printer error", vbExclamation
    '    Exit Function
  End If

End Function
Private Function PrinterSelect(ByVal Printer As String, ByVal IdPrinter As String) As Boolean
    On Error Resume Next
    PrinterSelect = Doc.Printer.SwitchTo(Printer, IdPrinter, True)
End Function

Public Function ShowFiller() As Boolean
    On Error Resume Next
    ShowFiller = CS7Server.ActiveDocument.Variables.ShowFiller
End Function

Public Function PrinterSettings() As Boolean
    On Error Resume Next
    PrinterSettings = CS7Server.ActivePrinterSetup
End Function

Public Function Labelhome(ByVal X As Double, Y As Double) As Boolean
On Error Resume Next
    CS7Server.ActiveDocument.Format.MarginLeft = X
    CS7Server.ActiveDocument.Format.MarginTop = Y
End Function

Public Function InitString(ByVal initstr As String) As Boolean
On Error Resume Next
    CS7Server.ActiveDocument.Printer.Send (initstr)
End Function

Public Function CloseLab() As Boolean
Dim LastErr As Integer
Dim ErrAnzeige As String
'On Error Resume Next
On Error GoTo Salah
'    CS7Server.ActiveDocument.Database.Close
'    CS7Server.ActiveDocument.Format.MarginLeft = OrgXachse
'    CS7Server.ActiveDocument.Format.MarginTop = OrgYachse
    CS7Server.ActiveDocument.Close True
    Set Doc = Nothing
    Set Vars = Nothing
    Exit Function
        
Salah:
    LastErr = CS7Server.Application.GetLastError
    If LastErr > 0 Then
        ErrAnzeige = CS7Server.Application.ErrorMessage(LastErr)
        MsgBox ErrAnzeige, vbCritical
        CloseLab = False
    Else
        CloseLab = True
    End If
End Function

'Public Function GetEtikett(ByVal Etikett$, ByRef Errtext) As Boolean
'    On Error GoTo ErrHandler
'
'    If Len(Etikett) > 0 Then
'
'         OpenLab (Etikett)
'         Exit Function
'    End If
'ErrHandler:
'     Errtext = True
'     Resume Next
'End Function

Public Function SetPrinters(ByVal Drucker$) As Boolean
On Error Resume Next
Dim I As Integer
Dim pos As Integer
Dim eintrag As String
Dim jumlah As Integer
    
    SetPrinters = False
    jumlah = CS7Server.PrinterSystem.Printers(lppxAllPrinters).Count
    For I = 0 To jumlah
        eintrag = CS7Server.PrinterSystem.Printers(lppxAllPrinters).Item(I)
        pos = InStr(1, eintrag, ",", vbTextCompare)
        If Trim(eintrag) = Trim(Drucker) Then
                SetPrinters = True
                Exit For
        End If
    Next
    If SetPrinters = True Then
        CS7Server.ActiveDocument.Printer.SwitchTo (Drucker)
    End If
End Function

Public Sub SetLabel(Dok As LabelManager2.Document, PageH As Integer, PageW As Integer, LabelH As Integer, LabelW As Integer)
    CS7Server.Options.MeasureSystem = lppxMillimeter
    Dok.Format.PageWidth = PageW * 100
    Dok.Format.PageHeight = PageH * 100
    Dok.Format.LabelHeight = LabelH * 100
    Dok.Format.LabelWidth = LabelW * 100
End Sub

Public Sub IsiText(CSTeks As LabelManager2.Text, Tulisan As String, Posisi As Long)
Dim A, b
    On Error GoTo ErrorHandle
    CSTeks.SelText.Select
    CSTeks.SelText.Cut
'    a = CSTeks.Font.Name
'    b = CSTeks.Font.Size
'    CSTeks.AppendString Tulisan, CSTeks.Font
    CSTeks.InsertString Tulisan, Posisi, CSTeks.Font
'    CSTeks.Font.Name = a
'    CSTeks.Font.Size = b
'    CSTeks.Value = Tulisan
    Exit Sub
ErrorHandle:
    MsgBox "Fatal Error(IsiText).", vbExclamation
End Sub
 
Public Sub SetFont(CSTeks As LabelManager2.Text, Nama As String, Tebal As Boolean, Ukuran As Long, Point As Boolean)
    On Error GoTo ErrorHandle
    With CSTeks.Font
        .Name = Nama
        .Bold = Tebal
        If Point Then
            .Size = Ukuran
        Else
            .Size = Ukuran * 2.8346
        End If
'        .Weight = 100
    End With
    Exit Sub
ErrorHandle:
    MsgBox "Fatal Error(SetFont).", vbExclamation
End Sub

Public Sub PrinterList()
Dim jumlah, I As Integer
Dim Printer As String

On Error GoTo ErrorHandle
    jumlah = CS7Server.PrinterSystem.Printers(lppxAllPrinters).Count
    For I = 0 To jumlah
        Printer = CS7Server.PrinterSystem.Printers(lppxAllPrinters).Item(I)
        frmOption.cboPrinter.AddItem Printer
    Next
    Exit Sub
ErrorHandle:
    MsgBox "Kesalahan dalam mengambil data printer.", vbExclamation
End Sub

Public Sub BuatLabel() ' Latihan untuk mengisi data
Dim A As Variant
    Set Doc = CS7Server.Documents.Add("Label")
'    Doc = CS7Server.ActiveDocument
'    Doc.Format.PageHeight = 44 * 100
'    Doc.Format.PageWidth = 44 * 100
'    Doc.Format.LabelHeight = 40 * 100
'    Doc.Format.LabelWidth = 40 * 100
    SetLabel Doc, 50, 50, 50, 50
    Set CSText = Doc.DocObjects.Add(lppxObjectText, "Teks01")
'    CSText = Doc.DocObjects.Texts.Item("Teks01")
    SetFont CSText, "Kartika", True, 10, True
    CSText.Value = "Percobaan"
    CSText.AnchorPoint = lppxTopLeft
    CSText.SelText.Select
    CSText.Left = 20 * 100
    CSText.Top = 20 * 100
    IsiText CSText, " Udah berhasil ", 0
'    CSText.InsertString "Udah Berhasil ", 1, CSText.Font
    SetFont CSText, "Arial", True, 20, False
'    DoEvent
'    a = CSText.Font
    Delay (0.1)
    Set CSText = Nothing
End Sub

