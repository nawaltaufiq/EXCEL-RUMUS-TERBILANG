# VBA Function: Terbilang

Fungsi VBA untuk mengubah angka menjadi teks terbilang dalam format Rupiah, tanpa pembulatan, dan tanpa menyebut angka di belakang koma.

## Fitur
1. **Angka Tidak Dibulatkan**:
   - Hanya bagian sebelum koma (bilangan bulat) yang dikonversi menjadi teks. Angka di belakang koma diabaikan.
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
    Dim Units As Double
    Dim TempStr As String

    ' Ambil hanya angka sebelum koma, TANPA pembulatan
    Units = Int(MyNumber)

    ' Konversi angka ke kata
    TempStr = AngkaKeKata(Units)

    ' Tambahkan "Rupiah" di akhir
    TempStr = TempStr & " Rupiah"

    ' Tambahkan * di awal dan akhir, dan kapital semua huruf
    TempStr = "*" & UCase(Trim(TempStr)) & "*"

    ' Hasil akhir
    Terbilang = TempStr
End Function

Function AngkaKeKata(ByVal MyNumber As Double) As String
    Dim TempStr As String
    Dim UnitNames As Variant
    Dim TensNames As Variant
    Dim BelasanNames As Variant
    Dim HundredsNames As Variant

    ' Array angka satuan, puluhan, belasan, ratusan
    UnitNames = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan")
    TensNames = Array("", "", "Dua Puluh", "Tiga Puluh", "Empat Puluh", "Lima Puluh", "Enam Puluh", "Tujuh Puluh", "Delapan Puluh", "Sembilan Puluh")
    BelasanNames = Array("Sepuluh", "Sebelas", "Dua Belas", "Tiga Belas", "Empat Belas", "Lima Belas", "Enam Belas", "Tujuh Belas", "Delapan Belas", "Sembilan Belas")
    HundredsNames = Array("", "Seratus", "Dua Ratus", "Tiga Ratus", "Empat Ratus", "Lima Ratus", "Enam Ratus", "Tujuh Ratus", "Delapan Ratus", "Sembilan Ratus")

    If MyNumber = 0 Then
        AngkaKeKata = "Nol"
        Exit Function
    End If

    TempStr = ""

    ' Tangani jutaan
    If MyNumber >= 1000000 Then
        TempStr = TempStr & AngkaKeKata(MyNumber \ 1000000) & " Juta "
        MyNumber = MyNumber Mod 1000000
    End If

    ' Tangani ribuan
    If MyNumber >= 1000 Then
        TempStr = TempStr & AngkaKeKata(MyNumber \ 1000) & " Ribu "
        MyNumber = MyNumber Mod 1000
    End If

    ' Tangani ratusan
    If MyNumber >= 100 Then
        TempStr = TempStr & HundredsNames(MyNumber \ 100) & " "
        MyNumber = MyNumber Mod 100
    End If

    ' Tangani belasan
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

    ' Ganti frasa tertentu
    TempStr = Replace(TempStr, "Satu Ribu", "Seribu")
    TempStr = Replace(TempStr, "Satu Ratus", "Seratus")

    AngkaKeKata = Trim(TempStr)
End Function
