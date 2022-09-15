Type TKaryawan
    NoInduk As String
    Nama As String
    JenisKelamin As String
    TanggalLahir As Date
    TanggalMasuk As Date
End Type

Sub Test()
Dim arrKaryawan(1 To 5) As TKaryawan
    With arrKaryawan(1)
        .NoInduk = "0001"
        .Nama = "Dab Atub"
        .JenisKelamin = "L"
        .TanggalLahir = "1970-08-15"
        .TanggalMasuk = "2002-03-07"
    End With

    With arrKaryawan(2)
        .NoInduk = "0002"
        .Nama = "Nahoj Bima"
        .JenisKelamin = "L"
        .TanggalLahir = "1974-06-15"
        .TanggalMasuk = "2004-02-22"
    End With

    With arrKaryawan(3)
        .NoInduk = "0003"
        .Nama = "Sam Rere"
        .JenisKelamin = "L"
        .TanggalLahir = "1989-10-12"
        .TanggalMasuk = "2001-10-28"
    End With
    
    With arrKaryawan(4)
        .NoInduk = "0004"
        .Nama = "Sam Inos"
        .JenisKelamin = "L"
        .TanggalLahir = "1998-07-12"
        .TanggalMasuk = "2006-04-17"
    End With
    
    With arrKaryawan(5)
        .NoInduk = "0005"
        .Nama = "Jaka Sampurna"
        .JenisKelamin = "L"
        .TanggalLahir = "1992-02-12"
        .TanggalMasuk = "2002-01-07"
    End With
    
    'Menampilkan Jendela Input Box
    Dim iLoop As Integer, sinputPrompt As String, sInputResult As String
    
    For iLoop = LBound(arrKaryawan) To UBound(arrKaryawan)
        sinputPrompt = sinputPrompt & iLoop & ". " & arrKaryawan(iLoop).NoInduk & ", " & arrKaryawan(iLoop).Nama & vbCrLf
    Next iLoop
    sinputPrompt = sinputPrompt & vbCrLf
    sinputPrompt = sinputPrompt & "Pilihan Anda (Ketikan dengan nomor urut [contoh 1] )"
    
    sInputResult = InputBox(sinputPrompt, "")
    
    If Trim(sInputResult) = "" Then
        MsgBox "Anda belum mengetikan pilihan", vbCritical, "Pesan Kesalahan"
        Exit Sub
    End If
    If Not IsNumeric(sInputResult) Then
        MsgBox "Masukan yang dijinkan hanya angka ya", vbCritical, "Pesan Kesalahan"
        Exit Sub
    End If
    Dim nVal As Integer, sMessage As String, sJenisKelamin As String, nUmur As Integer
    
    nVal = CInt(sInputResult)
    Select Case nVal
    Case 1 To 5
        Select Case arrKaryawan(nVal).JenisKelamin
        Case "L"
            sJenisKelamin = "Laki Laki"
        Case "P"
            sJenisKelamin = "Perempuan"
        Case Else
            sJenisKelamin = ""
        End Select
        sMessage = "No Induk: " & arrKaryawan(nVal).NoInduk & vbCrLf & _
            "Nama Karyawan: " & arrKaryawan(nVal).Nama & vbCrLf & _
            "Jenis Kelamin: " & sJenisKelamin & vbCrLf & _
            "Tanggal Lahir: " & arrKaryawan(nVal).TanggalLahir & vbCrLf & _
            "Umur: " & DateDiff("yyyy", arrKaryawan(nVal).TanggalLahir, Now())
            
        MsgBox sMessage, vbInformation, "Hasil Pilihan Anda"
    Case Else
        MsgBox "Pilihan anda tidak terdapat dalam daftar", vbCritical, "Pesan Kesalahan"
    End Select
End Sub

