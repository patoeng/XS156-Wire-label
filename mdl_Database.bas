Attribute VB_Name = "ModDatabase"
Option Explicit

'Database
'Public DB As Database
'Public RS As Recordset
Public Conn As ADODB.Connection
Public DataRS As ADODB.Recordset
Public Fld As ADODB.Field
Public Sql As String
Public StringKoneksi As String

Public Function BukaKoneksi() As Boolean
Dim Kondisi As Boolean

On Error GoTo Salah
    Set Conn = New ADODB.Connection
    Conn.ConnectionString = StringKoneksi
    Conn.Open
    Kondisi = Conn.State
    If Not Kondisi Then
        PesanSalah = "Problem BukaKoneksi" & vbCrLf & "Koneksi tidak terjadi, Cek LAN dan Komputer Sinba Packing"
        BukaKoneksi = False
        Conn.Close
    Else
        BukaKoneksi = True
    End If
    Exit Function
    
Salah:
    PesanSalah = "Problem BukaKoneksi" & vbCrLf & Err.Number & vbCrLf & Err.Description
    BukaKoneksi = False
End Function

Public Sub CloseKoneksi()

On Error Resume Next
    Conn.Close
    Set Conn = Nothing

End Sub

'Public Function CekKoneksi() As Boolean
'Dim Koneksi As Integer

'On Error GoTo Salah
'    Koneksi = Conn.State
'    If Koneksi = 1 Then
'        CekKoneksi = True
'    Else
'        PesanSalah = "Problem CekKoneksi" & vbCrLf & "Koneksi tak Berhasil, Cek Network"
'        CekKoneksi = False
'    End If
'    Exit Function

'Salah:
'    PesanSalah = "Problem CekKoneksi" & vbCrLf & Err.Number & vbCrLf & Err.Description
'    CekKoneksi = False
'End Function

Public Function SpecLog(ByVal Status As String, ByVal Keterangan As String, Optional ByVal User As String = "Tanpa User") As Boolean
Dim MyMsg As String
Dim ID As String
Dim FileLoc As String
Dim Kondisi As Boolean
    
    On Error Resume Next
    FileLoc = Format(Date, "yyww", vbSunday, vbFirstFullWeek) & ".txt"
    FileLoc = PathLog & "SpecLog" & FileLoc
        
'    MyMsg = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
    Open FileLoc$ For Append As #1
        
    Print #1, Now; ","; Status; ","; Keterangan; ","; User
        
    Close #1
    
On Error GoTo Salah
    Kondisi = BukaKoneksi
    If Kondisi Then
        MyMsg = Replace(Status, "'", "")
        Status = MyMsg
        Sql = "INSERT INTO `xs`.`spechistory` (`SQL`,`Alasan`,`User`,`TS`) VALUES ('"
        Sql = Sql & Status & "','" & Keterangan & "','" & User & "',null)"
        Conn.Execute Sql
        If Conn.Errors.Count > 0 Then
            PesanSalah = "Problem SpecLog" & vbCrLf & "Ada kesalahan (Data tak bisa disimpan)" & vbCrLf & Conn.Errors.Item(1).Description
            SpecLog = False
        Else
            SpecLog = True
        End If
    End If
    CloseKoneksi
    Exit Function

Salah:
    CloseKoneksi
    PesanSalah = "Problem SpecLog" & vbCrLf & Err.Number & vbCrLf & Err.Description
    SpecLog = False
End Function

Public Function MesinLog(ByVal Status As String, ByVal Keterangan As String, Optional ByVal User As String = "Tanpa User") As Boolean
Dim MyMsg As String
Dim ID As String
Dim FileLoc As String
Dim Kondisi As Boolean
    
    On Error Resume Next
    FileLoc = Format(Date, "yyww", vbSunday, vbFirstFullWeek) & ".txt"
    FileLoc = PathLog & "MesinLog" & FileLoc
        
'    MyMsg = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
    Open FileLoc$ For Append As #1
        
    Print #1, Now; ","; "Laser01Meas,"; Status; ","; Keterangan; ","; User
        
    Close #1
    
On Error GoTo Salah
    Kondisi = BukaKoneksi
    If Kondisi Then
        Sql = "INSERT INTO `xs`.`StartStop` (`Mesin`,`Status`,`Keterangan`,`User`,`TS`) VALUES ('"
        Sql = Sql & "Laser01Meas','" & Status & "','" & Keterangan & "','" & User & "',null)"
        Conn.Execute Sql
        If Conn.Errors.Count > 0 Then
            PesanSalah = "Problem MesinLog" & vbCrLf & "Ada kesalahan (Data tak bisa disimpan)" & vbCrLf & Conn.Errors.Item(1).Description
            MesinLog = False
        Else
            MesinLog = True
        End If
    End If
    CloseKoneksi
    Exit Function

Salah:
    CloseKoneksi
    PesanSalah = "Problem MesinLog" & vbCrLf & Err.Number & vbCrLf & Err.Description
    MesinLog = False
End Function

Public Sub HapusData(ByVal database As String, ByVal tabel As String, ByRef Status As Boolean, Optional ByVal Kriteria As String = "Kosong")
Dim Pesan As String
Dim Nilai As Integer
Dim Kondisi As Boolean

On Error GoTo Salah
    Kondisi = BukaKoneksi
    If Kondisi Then
        If Kriteria = "Kosong" Then
            Sql = "DELETE FROM `" & database & "`.`" & tabel & "`"
        Else
            Sql = "DELETE FROM `" & database & "`.`" & tabel & "` WHERE " & Kriteria
        End If
        Conn.Execute Sql
        Nilai = Conn.Errors.Count
        If Nilai > 0 Then
            Pesan = "Tak bisa hapus ke Database"
            GoTo Salah
        End If
    Else
        Pesan = "Database tidak aktif. Harap periksa koneksi LAN"
        GoTo Salah
    End If
    CloseKoneksi
    Status = True
    Exit Sub

Salah:
    CloseKoneksi
    If Pesan = "" Then Pesan = Err.Number & " - " & Err.Description
    PesanSalah = "PROBLEM TulisData" & vbCrLf & Pesan
    Status = False

End Sub

Public Sub TulisData(ByVal database As String, ByVal tabel As String, ByVal Kolom As String, ByVal Isi As String, ByRef Status As Boolean)
Dim Pesan As String
Dim Nilai As Integer
Dim Kondisi As Boolean

On Error GoTo Salah
    Kondisi = BukaKoneksi
    If Kondisi Then
        Sql = "INSERT INTO `" & database & "`.`" & tabel & "` (`" & Kolom & "`) VALUES ('" & Isi & "')"
        Conn.Execute Sql
        Nilai = Conn.Errors.Count
        If Nilai > 0 Then
            Pesan = "Tak bisa tulis ke Database"
            GoTo Salah
        End If
    Else
        Pesan = "Database tidak aktif. Harap periksa koneksi LAN"
        GoTo Salah
    End If
    CloseKoneksi
    Status = True
    Exit Sub

Salah:
    CloseKoneksi
    If Pesan = "" Then Pesan = Err.Number & " - " & Err.Description
    PesanSalah = "PROBLEM TulisData" & vbCrLf & Pesan
    Status = False
End Sub

Public Sub UbahPosisi(ByVal Posisi As Integer, ByVal Cavity As Integer, ByRef Status As Boolean)
Dim Pesan As String
Dim Nilai As Integer
Dim Kondisi As Boolean

On Error GoTo Salah
    Kondisi = BukaKoneksi
    If Kondisi Then
        Select Case Posisi
            Case 1  'Atas
                If Cavity = 1 Then
                    Sql = "UPDATE `xs`.`trim` SET YPos1=YPos1+0.1"
                ElseIf Cavity = 2 Then
                    Sql = "UPDATE `xs`.`trim` SET YPos2=YPos2+0.1"
                End If
            Case 2  'Bawah
                If Cavity = 1 Then
                    Sql = "UPDATE `xs`.`trim` SET YPos1=YPos1-0.1"
                ElseIf Cavity = 2 Then
                    Sql = "UPDATE `xs`.`trim` SET YPos2=YPos2-0.1"
                End If
            Case 3  'Kanan
                If Cavity = 1 Then
                    Sql = "UPDATE `xs`.`trim` SET XPos1=XPos1+0.1"
                ElseIf Cavity = 2 Then
                    Sql = "UPDATE `xs`.`trim` SET XPos2=XPos2+0.1"
                End If
            Case 4  'Kiri
                If Cavity = 1 Then
                    Sql = "UPDATE `xs`.`trim` SET XPos1=XPos1-0.1"
                ElseIf Cavity = 2 Then
                    Sql = "UPDATE `xs`.`trim` SET XPos2=XPos2-0.1"
                End If
        End Select
        Conn.Execute Sql
        Nilai = Conn.Errors.Count
        If Nilai > 0 Then
            Pesan = "Tak bisa tulis ke Database"
            GoTo Salah
        End If
    Else
        Pesan = "Database tidak aktif. Harap periksa koneksi LAN"
        GoTo Salah
    End If
    CloseKoneksi
    Status = True
    Exit Sub

Salah:
    CloseKoneksi
    If Pesan = "" Then Pesan = Err.Number & " - " & Err.Description
    PesanSalah = "PROBLEM TulisData" & vbCrLf & Pesan
    Status = False
End Sub

'Public Sub CariData(ByVal Kolom As String, ByVal database As String, ByVal tabel As String, ByVal Pilihan As String, ByRef Status As Boolean, Optional ByVal urut As String)
'Dim Kondisi As Boolean
'Dim Nilai As Integer
'Dim Pesan As String
'Dim Temp As String

'On Error GoTo Salah
'    Kondisi = BukaKoneksi
'    If Kondisi Then
'        If IsMissing(urut) Then
'            Sql = "SELECT `" & Kolom & "` FROM `" & database & "`.`" & tabel & "` WHERE `" & Pilihan & "'"
'        Else
'            Sql = "SELECT `" & Kolom & "` FROM `" & database & "`.`" & tabel & "` WHERE `" & Pilihan & "' " & "ORDER BY `" & urut
'        End If
'        Set DataRS = New ADODB.Recordset
'        DataRS.CursorLocation = adUseClient
'        DataRS.Open Sql, Conn, adOpenStatic, adLockReadOnly
'        Nilai = DataRS.RecordCount
'        If Nilai < 1 Then
'            Pesan = "Ada kesalahan (Data tak bisa diambil)"
'            GoTo Salah
'        Else
'            DataRS.MoveFirst
'            Do While Not DataRS.EOF
'                Temp = DataRS.Fields("Nama")
'                Combo.AddItem Temp
'                DataRS.MoveNext
'            Loop
'        End If
'        DataRS.Close
'        Set DataRS = Nothing
'    Else
'        Pesan = "Database tidak aktif. Harap periksa koneksi LAN"
'        GoTo Salah
'    End If
'    CloseKoneksi
'    Status = True
'    Exit Sub

'Salah:
'    CloseKoneksi
'    If Pesan = "" Then Pesan = Err.Number & " - " & Err.Description
'    PesanSalah = "PROBLEM IsiCombo" & vbCrLf & Pesan
'    Status = False
'End Sub

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
            Sql = "SELECT `" & Kolom & "` FROM `" & database & "`.`" & tabel & "` ORDER BY `" & urut
        Else
            Sql = "SELECT `" & Kolom & "` FROM `" & database & "`.`" & tabel & "` WHERE `" & Kriteria & " ORDER BY `" & urut
        End If
        DataRS.CursorLocation = adUseClient
        DataRS.Open Sql, Conn, adOpenStatic, adLockReadOnly
        Nilai = DataRS.RecordCount
        If Nilai < 1 Then
            Pesan = "Ada kesalahan (Data tak bisa diambil)"
            GoTo Salah
        Else
            DataRS.MoveFirst
            Combo.AddItem ""
            Do While Not DataRS.EOF
                Temp = DataRS.Fields("Nama")
                Combo.AddItem Temp
                DataRS.MoveNext
            Loop
        End If
        DataRS.Close
        Set DataRS = Nothing
    Else
        Pesan = "Database tidak aktif. Harap periksa koneksi LAN"
        Kondisi = MesinLog("RUN  FAIL", "IsiCombo - Koneksi ke Database Terputus")
        GoTo Salah
    End If
    CloseKoneksi
    Status = True
    Exit Sub

Salah:
    CloseKoneksi
    If Pesan = "" Then Pesan = Err.Number & " - " & Err.Description
    PesanSalah = "PROBLEM IsiCombo" & vbCrLf & Pesan
    Status = False
End Sub

'Public Sub BacaSpec()
'Dim Pos As Integer
'Dim FNum As Integer
'Dim ItemStr As String
'Dim SectionHeading As String
'Dim NilaiStr As String
'Dim LineStr As String

'    FNum = FreeFile
'    ItemStr = PathData & "analogspec.txt"
'    If Dir(ItemStr) = "" Then
'        SetDefaultSpec
'        TulisSpec
'    End If
    
'    Open ItemStr For Input As FNum
''    Open App.Path & "\database\analogspec.txt" For Input As FNum
    
'    With Laser_Skr
'        Do While Not EOF(FNum)
'            Line Input #FNum, LineStr
'            Pos = InStr(LineStr, "=")
'            ItemStr = Left$(LineStr, Pos - 1)
'            NilaiStr = Mid$(LineStr, Pos + 1)
'            Select Case UCase(ItemStr)
'                Case "A1MAX"
'                    frmSelectAna.txt_A1Max.Text = NilaiStr
'                Case "A1TGT"
'                    frmSelectAna.txt_A1Tgt.Text = NilaiStr
'                Case "A1MIN"
'                    frmSelectAna.txt_A1Min.Text = NilaiStr
'                Case "A2MAX"
'                    frmSelectAna.txt_A2Max.Text = NilaiStr
'                Case "A2TGT"
'                    frmSelectAna.txt_A2Tgt.Text = NilaiStr
'                Case "A2MIN"
'                    frmSelectAna.txt_A2Min.Text = NilaiStr
'                Case "SN"
'                    .SN = NilaiStr
'                    frmSelectAna.txt_SNTest.Text = .SN
'                Case "MAXI"
'                    BatasI = NilaiStr
'                    frmSelectAna.txt_MaxI.Text = BatasI
'                Case "MAXTRIM"
'                    BatasTrim = NilaiStr
'                    frmSelectAna.Txt_MaxTrim.Text = BatasTrim
'                Case "XPOS(1)"
'                    .XPos(1) = NilaiStr
'                    frmSelectAna.txt_XPos(1).Text = .XPos(1)
'                Case "YPOS(1)"
'                    .YPos(1) = NilaiStr
'                    frmSelectAna.txt_YPos(1).Text = .YPos(1)
'                Case "XPOS(2)"
'                    .XPos(2) = NilaiStr
'                    frmSelectAna.txt_XPos(2).Text = .XPos(2)
'                Case "YPOS(2)"
'                    .YPos(2) = NilaiStr
'                    frmSelectAna.txt_YPos(2).Text = .YPos(2)
'                Case "JIGTYPE"
'                    .JigType = NilaiStr
'                Case "JIGOFF(1)"
'                    .JigOffset(1) = NilaiStr
'                Case "JIGOFF(2)"
'                    .JigOffset(2) = NilaiStr
'                Case "TARGETOFF"
'                    .TarOffset = NilaiStr
'                Case "TARGET"
'                    .Target = NilaiStr
'                Case "LASERSPEED"
'                    .LaserSpeed = NilaiStr
'                Case "LASERQS"
'                    .LaserQS = NilaiStr
'                Case "LASERPOWER"
'                    .LaserPower = NilaiStr
'                Case "CUTDIR"
'                    .CutDir = NilaiStr
'            End Select
'        Loop
'    End With
'    Close FNum
'End Sub

'Write data to INI file
'Public Function TulisSpec()
'Dim FNum As Integer

'    FNum = FreeFile
    
'    Open PathData & "AnalogSpec.txt" For Output As FNum
''    Open App.Path & "\database\AnalogSpec.txt" For Output As FNum
    
'    With Laser_Skr
'        Print #FNum, "A1Max=" & frmSelectAna.txt_A1Max.Text
'        Print #FNum, "A1Tgt=" & frmSelectAna.txt_A1Tgt.Text
'        Print #FNum, "A1Min=" & frmSelectAna.txt_A1Min.Text
'        Print #FNum, "A2Max=" & frmSelectAna.txt_A2Max.Text
'        Print #FNum, "A2Tgt=" & frmSelectAna.txt_A2Tgt.Text
'        Print #FNum, "A2Min=" & frmSelectAna.txt_A2Min.Text
'        Print #FNum, "SN=" & .SN
'        Print #FNum, "MaxI=" & BatasI
'        Print #FNum, "MaxTrim=" & BatasTrim
'        Print #FNum, "XPos(1)=" & .XPos(1)
'        Print #FNum, "YPos(1)=" & .YPos(1)
'        Print #FNum, "XPos(2)=" & .XPos(2)
'        Print #FNum, "YPos(2)=" & .YPos(2)
'        Print #FNum, "JigType=" & .JigType
'        Print #FNum, "JigOff(1)=" & .JigOffset(1)
'        Print #FNum, "JigOff(2)=" & .JigOffset(2)
'        Print #FNum, "TargetOff=" & .TarOffset
'        Print #FNum, "Target=" & .Target
'        Print #FNum, "LaserSpeed=" & .LaserSpeed
'        Print #FNum, "LaserQS=" & .LaserQS
'        Print #FNum, "LaserPower=" & .LaserPower
'        Print #FNum, "CutDir=" & .CutDir
'    End With
    
'    Close FNum

'End Function

'Public Sub SetDefaultSpec()
'    Call frmSelectAna.BatasAwal
'    With Laser_Skr
'        .SN = 5
'        BatasI = 30
'        BatasTrim = 300
'        .XPos(1) = -29.5
'        .YPos(1) = 4.2
'        .XPos(2) = -28.9
'        .YPos(2) = 4.7
'        .JigType = 2
'        .JigOffset(1) = 2.16
'        .JigOffset(2) = 2.31
'        .TarOffset = 0
'        .Target = 1
'        .LaserSpeed = 2.54
'        .LaserQS = 3
'        .LaserPower = 22
'        .CutDir = 4
'    End With
'End Sub



