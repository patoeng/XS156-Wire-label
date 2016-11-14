Attribute VB_Name = "Database"
Option Explicit

Public Function BukaKoneksi() As Boolean
Dim Kondisi As Boolean

On Error GoTo Salah
    If TypeName(Conn) = "Nothing" Then
        Set Conn = New ADODB.Connection
        Conn.ConnectionString = StringKoneksi
        Conn.Open
    End If
    Kondisi = Conn.State
    If Not Kondisi Then
        BukaKoneksi = False
'        Conn.Close
    Else
'        PesanSalah = "OKE - Buka Koneksi"
        BukaKoneksi = True
    End If
    Exit Function
    
Salah:
    PesanSalah = "Problem saat buka koneksi" & vbCrLf & Err.Number & vbCrLf & Err.Description
    BukaKoneksi = False
End Function

Public Sub TutupKoneksi()
'Dim Kondisi As String

On Error Resume Next
    If Not (TypeName(Conn) = "Nothing") Then
        Conn.Close
    End If
    Set Conn = Nothing
End Sub

'Public Function CekKoneksi() As Boolean
'Dim Koneksi As Integer

'On Error GoTo Salah
'    Koneksi = Conn.State
''    Conn.Close
''    Set Conn = Nothing
'    If Koneksi = 1 Then
'        PesanSalah = "OKE - Cek Koneksi"
'        CekKoneksi = True
'    Else
'        PesanSalah = "NOK - Cek Koneksi tak berhasil"
'        CekKoneksi = False
'    End If
'    Exit Function

'Salah:
'    PesanSalah = "Problem saat Cek Koneksi" & vbCrLf & Err.Number & vbCrLf & Err.Description
'End Function

Public Sub IsiCombo(ByVal Kolom As String, ByVal database As String, ByVal tabel As String, ByVal urut As String, Combo As ComboBox, ByRef Status As Boolean, Optional ByVal Kriteria As String)
Dim Nilai As Integer
Dim Kondisi As Boolean
Dim Pesan As String
Dim Temp As String

On Error GoTo Salah
    Combo.Clear
    Kondisi = BukaKoneksi
    If Kondisi Then
        Set DataRS = New ADODB.Recordset
        If IsMissing(Kriteria) Or Trim(Kriteria) = "" Then
            SQL = "SELECT `" & Kolom & "` FROM `" & database & "`.`" & tabel & "` ORDER BY `" & urut
        Else
            SQL = "SELECT `" & Kolom & "` FROM `" & database & "`.`" & tabel & "` WHERE `" & Kriteria & " ORDER BY `" & urut
        End If
        DataRS.CursorLocation = adUseClient
        DataRS.Open SQL, Conn, adOpenStatic, adLockReadOnly
        Nilai = DataRS.RecordCount
        If Nilai < 1 Then
            Pesan = "Ada kesalahan (Data tak bisa diambil)"
            GoTo Salah
        Else
            DataRS.MoveFirst
            Combo.AddItem ""
            Do While Not DataRS.EOF
                Temp = DataRS.Fields(Kolom)
                Combo.AddItem Temp
                DataRS.MoveNext
            Loop
        End If
        DataRS.Close
        Set DataRS = Nothing
    Else
        Pesan = "Database tidak aktif. Harap periksa koneksi LAN"
'        Kondisi = MesinLog("RUN  FAIL", "IsiCombo - Koneksi ke Database Terputus")
        GoTo Salah
    End If
    TutupKoneksi
    Status = True
    Exit Sub

Salah:
    TutupKoneksi
    If Pesan = "" Then Pesan = Err.Number & " - " & Err.Description
    PesanSalah = "PROBLEM IsiCombo" & vbCrLf & Pesan
    Status = False
End Sub

Public Sub AmbilLab(ByVal database As String, ByVal tabel As String, ByVal urut As String, ByVal Kriteria As String, ByRef Status As Boolean)
Dim Kondisi As Boolean
Dim Nilai As Integer
Dim I As Integer
Dim Pesan As String
Dim Temp As String

On Error GoTo Salah
    Kondisi = BukaKoneksi
    If Kondisi Then
        Set DataRS = New ADODB.Recordset
        If IsMissing(Kriteria) Or Trim(Kriteria) = "" Then
            SQL = "SELECT * FROM `" & database & "`.`" & tabel & "` ORDER BY `" & urut
        Else
            SQL = "SELECT * FROM `" & database & "`.`" & tabel & "` WHERE `" & Kriteria & " ORDER BY `" & urut
        End If
        DataRS.CursorLocation = adUseClient
        DataRS.Open SQL, Conn, adOpenStatic, adLockReadOnly
        Nilai = DataRS.RecordCount
        If Nilai < 1 Then
            Pesan = "Ada kesalahan (Data tak bisa diambil)"
            GoTo Salah
        Else
            DataRS.MoveFirst
            For I = 1 To 5
                Temp = "" & DataRS.Fields("Template0" & I)
                Lab(I) = Temp
            Next I
        End If
        DataRS.Close
        Set DataRS = Nothing
    Else
        Pesan = "Database tidak aktif. Harap periksa koneksi LAN"
'        Kondisi = MesinLog("RUN  FAIL", "IsiCombo - Koneksi ke Database Terputus")
        GoTo Salah
    End If
    TutupKoneksi
    Status = True
    Exit Sub

Salah:
    If Pesan = "" Then Pesan = Err.Number & " - " & Err.Description
    TutupKoneksi
    PesanSalah = "PROBLEM AmbilLab" & vbCrLf & Pesan
    Status = False
End Sub

Public Sub AmbilKolom(ByVal Kolom As String, ByVal database As String, ByVal tabel As String, ByVal urut As String, ByVal Kriteria As String, ByRef Status As Boolean, ByRef NilaiKolom As String)
Dim Kondisi As Boolean
Dim Nilai As Integer
Dim I As Integer
Dim Pesan As String
Dim Temp As String

On Error GoTo Salah
    Kondisi = BukaKoneksi
    If Kondisi Then
        Set DataRS = New ADODB.Recordset
        If IsMissing(Kriteria) Or Trim(Kriteria) = "" Then
            SQL = "SELECT `" & Kolom _
                & "` FROM `" & database & "`.`" & tabel & "` ORDER BY `" & urut
        Else
            SQL = "SELECT `" & Kolom _
                & "` FROM `" & database & "`.`" & tabel _
                & "` WHERE `" & Kriteria & " ORDER BY `" & urut
        End If
        DataRS.CursorLocation = adUseClient
        DataRS.Open SQL, Conn, adOpenStatic, adLockReadOnly
        Nilai = DataRS.RecordCount
        If Nilai < 1 Then
            Pesan = "Ada kesalahan (Data tak bisa diambil)"
            GoTo Salah
        Else
            DataRS.MoveFirst
            Temp = "" & DataRS.Fields(Kolom)
        End If
        NilaiKolom = Temp
        DataRS.Close
        Set DataRS = Nothing
    Else
        Pesan = "Database tidak aktif. Harap periksa koneksi LAN"
'        Kondisi = MesinLog("RUN  FAIL", "IsiCombo - Koneksi ke Database Terputus")
        GoTo Salah
    End If
    TutupKoneksi
    Status = True
    Exit Sub

Salah:
    If Pesan = "" Then Pesan = Err.Number & " - " & Err.Description
    TutupKoneksi
    PesanSalah = "PROBLEM AmbilKolom" & vbCrLf & Pesan
    Status = False
End Sub

Public Sub AmbilLabAktif(ByVal database As String, ByVal tabel As String, ByVal urut As String, ByVal Kriteria As String, ByRef Status As Boolean)
Dim Kondisi As Boolean
Dim Nilai As Integer
Dim I As Integer
Dim Pesan As String
Dim Temp As String

On Error GoTo Salah
    Kondisi = BukaKoneksi
    If Kondisi Then
        Set DataRS = New ADODB.Recordset
        If IsMissing(Kriteria) Or Trim(Kriteria) = "" Then
            SQL = "SELECT * FROM `" & database & "`.`" & tabel & "` ORDER BY `" & urut
        Else
            SQL = "SELECT * FROM `" & database & "`.`" & tabel & "` WHERE `" & Kriteria & " ORDER BY `" & urut
        End If
        DataRS.CursorLocation = adUseClient
        DataRS.Open SQL, Conn, adOpenStatic, adLockReadOnly
        Nilai = DataRS.RecordCount
        If Nilai < 1 Then
            Pesan = "Ada kesalahan (Data tak bisa diambil)"
            GoTo Salah
        Else
            DataRS.MoveFirst
            For I = 1 To 5
                Temp = "" & DataRS.Fields("Template0" & I)
                If Temp = "1" Then
                    LabAktif(I) = True
                Else
                    LabAktif(I) = False
                End If
            Next I
        End If
        DataRS.Close
        Set DataRS = Nothing
    Else
        Pesan = "Database tidak aktif. Harap periksa koneksi LAN"
'        Kondisi = MesinLog("RUN  FAIL", "IsiCombo - Koneksi ke Database Terputus")
        GoTo Salah
    End If
    TutupKoneksi
    Status = True
    Exit Sub

Salah:
    TutupKoneksi
    If Pesan = "" Then Pesan = Err.Number & " - " & Err.Description
    PesanSalah = "PROBLEM IsiCombo" & vbCrLf & Pesan
    Status = False
End Sub

Public Function SetingParameter() As Boolean
On Error GoTo Salah
    If GetSetting(App.Title, "Settings", "Server") = "" Then
        ServerAlamat = "10.184.65.211"
        SaveSetting App.Title, "Settings", "Server", ServerAlamat
    Else
        ServerAlamat = GetSetting(App.Title, "Settings", "Server")
    End If
    If GetSetting(App.Title, "Settings", "ServerData") = "" Then
        ServerData = "XS"
        SaveSetting App.Title, "Settings", "ServerData", ServerData
    Else
        ServerData = GetSetting(App.Title, "Settings", "ServerData")
    End If
    If GetSetting(App.Title, "Settings", "ServerUser") = "" Then
        ServerUser = "OperatorXS"
        SaveSetting App.Title, "Settings", "ServerUser", ServerUser
    Else
        ServerUser = GetSetting(App.Title, "Settings", "ServerUser")
    End If
    If GetSetting(App.Title, "Settings", "ServerPass") = "" Then
        ServerPass = "Inductive"
        SaveSetting App.Title, "Settings", "ServerPass", ServerPass
    Else
        ServerPass = GetSetting(App.Title, "Settings", "ServerPass")
    End If
    If GetSetting(App.Title, "Settings", "ServerDriver") = "" Then
        ServerDriver = "{MySQL ODBC 5.1 Driver}"   '{MySQL ODBC 3.51 Driver}
        SaveSetting App.Title, "Settings", "ServerDriver", ServerDriver
    Else
        ServerDriver = GetSetting(App.Title, "Settings", "ServerDriver")
    End If
    If GetSetting(App.Title, "Settings", "ServerOption") = "" Then
        ServerOption = "35"
        SaveSetting App.Title, "Settings", "ServerOption", ServerOption
    Else
        ServerOption = GetSetting(App.Title, "Settings", "ServerOption")
    End If
    If GetSetting(App.Title, "Settings", "Label") = "" Then
        PathLabel = App.Path & "\Label\"
        SaveSetting App.Title, "Settings", "Label", PathLabel
    Else
        PathLabel = GetSetting(App.Title, "Settings", "Label")
    End If
    If GetSetting(App.Title, "Settings", "Data") = "" Then
        PathData = App.Path & "\Database\"
        SaveSetting App.Title, "Settings", "Data", PathData
    Else
        PathData = GetSetting(App.Title, "Settings", "Data")
    End If
    If GetSetting(App.Title, "Settings", "Gambar") = "" Then
        PathGambar = App.Path & "\Gambar\"
        SaveSetting App.Title, "Settings", "Gambar", PathGambar
    Else
        PathGambar = GetSetting(App.Title, "Settings", "Gambar")
    End If
    If GetSetting(App.Title, "Settings", "Sementara") = "" Then
        PathTemp = App.Path & "\Temp\"
        SaveSetting App.Title, "Settings", "Sementara", PathTemp
    Else
        PathTemp = GetSetting(App.Title, "Settings", "Sementara")
    End If
    If GetSetting(App.Title, "Settings", "NamaLabel") = "" Then
        NamaLabel = App.Path & "\PCBA.Lab"
        SaveSetting App.Title, "Settings", "NamaLabel", NamaLabel
    Else
        NamaLabel = GetSetting(App.Title, "Settings", "NamaLabel")
    End If
    If GetSetting(App.Title, "Settings", "Keluarga") = "" Then
        Keluarga = "XS156"
        SaveSetting App.Title, "Settings", "Keluarga", Keluarga
    Else
        Keluarga = GetSetting(App.Title, "Settings", "Keluarga")
    End If
    If GetSetting(App.Title, "Settings", "PosisiAtas") = "" Then
        PosisiAtas = "0"
        SaveSetting App.Title, "Settings", "PosisiAtas", PosisiAtas
    Else
        PosisiAtas = GetSetting(App.Title, "Settings", "PosisiAtas")
    End If
    If GetSetting(App.Title, "Settings", "PosisiKiri") = "" Then
        PosisiKiri = "0"
        SaveSetting App.Title, "Settings", "PosisiKiri", PosisiKiri
    Else
        PosisiKiri = GetSetting(App.Title, "Settings", "PosisiKiri")
    End If
    If GetSetting(App.Title, "Settings", "SatuKeluarga") = "" Then
        SatuKeluarga = "True"
        SaveSetting App.Title, "Settings", "SatuKeluarga", SatuKeluarga
    Else
        SatuKeluarga = GetSetting(App.Title, "Settings", "SatuKeluarga")
    End If
    
    If GetSetting(App.Title, "Settings", "GPIBMultiMeter") = "" Then
        AlamatMultiMeter = "2"
        SaveSetting App.Title, "Settings", "GPIBMultiMeter", AlamatMultiMeter
    Else
        AlamatMultiMeter = GetSetting(App.Title, "Settings", "GPIBMultiMeter")
    End If
    If GetSetting(App.Title, "Settings", "IDMultiMeter") = "" Then
        KodeMultiMeter = "FLUKE,8845A,9499033,04/02/07-08:10"
        SaveSetting App.Title, "Settings", "IDMultiMeter", KodeMultiMeter
    Else
        KodeMultiMeter = GetSetting(App.Title, "Settings", "IDMultiMeter")
    End If
    If GetSetting(App.Title, "Settings", "GPIBPowerSupply") = "" Then
        AlamatPowerSupply = "3"
        SaveSetting App.Title, "Settings", "GPIBPowerSupply", AlamatPowerSupply
    Else
        AlamatPowerSupply = GetSetting(App.Title, "Settings", "GPIBPowerSupply")
    End If
    If GetSetting(App.Title, "Settings", "IDPowerSupply") = "" Then
        KodePowerSupply = "EXTECH ELECTRONICS. LTD.,6605,1991016,Version1.02"
        SaveSetting App.Title, "Settings", "IDPowerSupply", KodePowerSupply
    Else
        KodePowerSupply = GetSetting(App.Title, "Settings", "IDPowerSupply")
    End If
    If GetSetting(App.Title, "Settings", "PassAdmin") = "" Then
        PassAdmin = "Call FBI"
        SaveSetting App.Title, "Settings", "PassAdmin", PassAdmin
    Else
        PassAdmin = GetSetting(App.Title, "Settings", "PassAdmin")
    End If
    If GetSetting(App.Title, "Settings", "PassOperator") = "" Then
        PassOperator = "Call 911"
        SaveSetting App.Title, "Settings", "PassOperator", PassOperator
    Else
        PassOperator = GetSetting(App.Title, "Settings", "PassOperator")
    End If

    StringKoneksi = "DRIVER=" & ServerDriver & "; SERVER=" & ServerAlamat & "; DATABASE=" & ServerData & "; UID=" & ServerUser & ";PWD=" & ServerPass & "; OPTION=" & ServerOption
'    PesanSalah = "OKE - Setting Parameter"
    SetingParameter = True
    Exit Function
Salah:
    PesanSalah = "Problem dengan setting parameter" & vbCrLf & Err.Description
    SetingParameter = False
End Function

