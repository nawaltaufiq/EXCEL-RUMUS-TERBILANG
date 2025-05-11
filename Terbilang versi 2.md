# VBA Function: Terbilang (Versi 2)

Fungsi VBA untuk mengubah angka menjadi teks terbilang dalam format Rupiah, dengan menyebutkan angka di belakang koma secara lengkap.

## Fitur
1. **Angka Desimal Disebutkan**:
   - Bagian angka di belakang koma dikonversi menjadi teks per digit.
   - Contoh: `1234.56` akan menjadi *"SERIBU DUA RATUS TIGA PULUH EMPAT LIMA ENAM RUPIAH"*.
2. **Mendukung Bilangan Besar**:
   - Dapat menangani angka hingga jutaan, ribuan, ratusan, puluhan, dan satuan.
3. **Format Rupiah**:
   - Otomatis menambahkan kata "Rupiah" di akhir hasil konversi.
4. **Teks Kapital dan Dibungkus dengan Asteris (`*`)**:
   - Hasil akhir dikonversi menjadi huruf kapital dan dibungkus dengan asteris di awal dan akhir.
5. **Frasa Khusus**:
   - Otomatis mengganti "Satu Ribu" menjadi "Seribu" dan "Satu Ratus" menjadi "Seratus" untuk hasil yang lebih natural.

---

## Kode VBA

```vb
Function Terbilang(ByVal MyNumber As Double) As String
    Dim NumberStr As String
    Dim Parts() As String
    Dim Units As String, Decimals As String
    Dim TempStr As String
    Dim i As Integer
    Dim UnitNames As Variant

    ' Array nama angka
    UnitNames = Array("Nol", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan")

    ' Ubah angka ke string dengan titik sebagai pemisah desimal
    NumberStr = Replace(Trim(CStr(MyNumber)), ",", ".")
    Parts = Split(NumberStr, ".")

    ' Ambil bagian sebelum dan sesudah koma
    Units = Parts(0)
    If UBound(Parts) > 0 Then
        Decimals = Parts(1)
    Else
        Decimals = ""
    End If

    ' Ubah bagian sebelum koma ke kata
    TempStr = AngkaKeKata(CDbl(Units))

    ' Ubah bagian desimal ke kata (per digit)
    If Len(Decimals) > 0 Then
        For i = 1 To Len(Decimals)
            TempStr = TempStr & " " & UnitNames(Mid(Decimals, i, 1))
        Next i
    End If

    ' Tambahkan "Rupiah" dan * di awal/akhir
    TempStr = "*" & UCase(Trim(TempStr & " Rupiah")) & "*"

    Terbilang = TempStr
End Function

Function AngkaKeKata(ByVal MyNumber As Double) As String
    Dim TempStr As String
    Dim UnitNames As Variant
    Dim TensNames As Variant
    Dim BelasanNames As Variant
    Dim HundredsNames As Variant

    ' Array kata angka
    UnitNames = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan")
    TensNames = Array("", "", "Dua Puluh", "Tiga Puluh", "Empat Puluh", "Lima Puluh", "Enam Puluh", "Tujuh Puluh", "Delapan Puluh", "Sembilan Puluh")
    BelasanNames = Array("Sepuluh", "Sebelas", "Dua Belas", "Tiga Belas", "Empat Belas", "Lima Belas", "Enam Belas", "Tujuh Belas", "Delapan Belas", "Sembilan Belas")
    HundredsNames = Array("", "Seratus", "Dua Ratus", "Tiga Ratus", "Empat Ratus", "Lima Ratus", "Enam Ratus", "Tujuh Ratus", "Delapan Ratus", "Sembilan Ratus")

    If MyNumber = 0 Then
        AngkaKeKata = "Nol"
        Exit Function
    End If

    TempStr = ""

    If MyNumber >= 1000000 Then
        TempStr = TempStr & AngkaKeKata(MyNumber \ 1000000) & " Juta "
        MyNumber = MyNumber Mod 1000000
    End If

    If MyNumber >= 1000 Then
        TempStr = TempStr & AngkaKeKata(MyNumber \ 1000) & " Ribu "
        MyNumber = MyNumber Mod 1000
    End If

    If MyNumber >= 100 Then
        TempStr = TempStr & HundredsNames(MyNumber \ 100) & " "
        MyNumber = MyNumber Mod 100
    End If

    If MyNumber >= 10 And MyNumber < 20 Then
        TempStr = TempStr & BelasanNames(MyNumber - 10)
    Else
        If MyNumber >= 20 Then
            TempStr = TempStr & TensNames(MyNumber \ 10) & " "
        End If
        If MyNumber Mod 10 > 0 Then
            TempStr = TempStr & UnitNames(MyNumber Mod 10)
        End If
    End If

    ' Penyesuaian frasa
    TempStr = Replace(TempStr, "Satu Ribu", "Seribu")
    TempStr = Replace(TempStr, "Satu Ratus", "Seratus")

    AngkaKeKata = Trim(TempStr)
End Function
