Attribute VB_Name = "Modül1"
Option Explicit
Function YAZIYLA(sayi As Variant) As String
' Sayýyý yazýyla yazar
'
' Mesut Akcan
' https://www.mesutakcan.blogspot.com
' makcan@gmail.com
'
' 23 Nisan 2004
' Güncelleme: 5 Haziran 2025

Dim birler(9) As String, onlar(9) As String, buyukSayi(4) As String
Dim basamak(1 To 15) As Byte, grup(1 To 3) As Byte
Dim sayiMetni As String, grupMetni As String
Dim sonuc As String
Dim negatif As Boolean
Dim i As Byte, j As Byte 'index


If (Not IsNumeric(sayi)) Or (Len(sayi) > 15) Then ' Sayý deðilse veya 15 basamaktan büyükse hata
    YAZIYLA = "#HATA!"
    Exit Function
End If

If sayi < 0 Then
  negatif = True 'Sayý negatif
  sayi = Abs(sayi)
End If

' Birler basamaðý
birler(0) = ""
birler(1) = "Bir"
birler(2) = "Ýki"
birler(3) = "Üç"
birler(4) = "Dört"
birler(5) = "Beþ"
birler(6) = "Altý"
birler(7) = "Yedi"
birler(8) = "Sekiz"
birler(9) = "Dokuz"

' Onlar basamaðý
onlar(0) = ""
onlar(1) = "On"
onlar(2) = "Yirmi"
onlar(3) = "Otuz"
onlar(4) = "Kýrk"
onlar(5) = "Elli"
onlar(6) = "Altmýþ"
onlar(7) = "Yetmiþ"
onlar(8) = "Seksen"
onlar(9) = "Doksan"

' Büyük sayýlar
buyukSayi(0) = "Trilyon "
buyukSayi(1) = "Milyar "
buyukSayi(2) = "Milyon "
buyukSayi(3) = "Bin "
buyukSayi(4) = ""

sayiMetni = Right(String(15, "0") & CStr(Fix(sayi)), 15) ' Sayýyý metne çevir ve boþluklarý kaldýr

' 1'den 15'e kadar döngü
' karakterleri tek tek al ve sayýya çevir ve diziye aktar
For i = 1 To 15
    basamak(i) = CByte(Mid(sayiMetni, i, 1))
Next

sonuc = "" ' Sonuç metni

' sayý metnini 3'erli 5 gruba ayýr ve her grubu yazýya çevir
For i = 0 To 4
  For j = 1 To 3 'gruptaki yüzler, onlar, birler basamaklarý
    grup(j) = basamak((i * 3) + j)
  Next
  
  Select Case grup(1) ' Yüzler basamaðý
    Case 0 ' sýfýr ise
      grupMetni = "" ' Yüzler basamaðý metni boþ
    Case 1 ' 1 ise
      grupMetni = "Yüz" ' Yüzler basamaðý metni "Yüz"
    Case Else ' 2-9 arasý ise
      grupMetni = birler(grup(1)) & "Yüz" ' Yüzler basamaðý metni "ÝkiYüz", "ÜçYüz" vb.
  End Select
  
  grupMetni = grupMetni & onlar(grup(2)) & birler(grup(3)) ' Onlar ve birler basamaðýný ekle
  
  If grupMetni <> "" Then
    grupMetni = grupMetni & buyukSayi(i) ' Büyük sayýlarý ekle
    If (i = 3) And (grupMetni = "BirBin ") Then
      grupMetni = "Bin" ' "BirBin" durumunu düzelt
    End If
  End If
  sonuc = sonuc & grupMetni ' Sonucu birleþtir
Next
sonuc = Trim(sonuc)
If sonuc = "" Then
  sonuc = "Sýfýr"
ElseIf negatif Then
  sonuc = "Eksi " & sonuc
End If
YAZIYLA = sonuc
End Function

