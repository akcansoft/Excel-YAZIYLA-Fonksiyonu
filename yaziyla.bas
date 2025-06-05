Attribute VB_Name = "Mod�l1"
Option Explicit
Function YAZIYLA(sayi As Variant) As String
' Say�y� yaz�yla yazar
'
' Mesut Akcan
' https://www.mesutakcan.blogspot.com
' makcan@gmail.com
'
' 23 Nisan 2004
' G�ncelleme: 5 Haziran 2025

Dim birler(9) As String, onlar(9) As String, buyukSayi(4) As String
Dim basamak(1 To 15) As Byte, grup(1 To 3) As Byte
Dim sayiMetni As String, grupMetni As String
Dim sonuc As String
Dim negatif As Boolean
Dim i As Byte, j As Byte 'index


If (Not IsNumeric(sayi)) Or (Len(sayi) > 15) Then ' Say� de�ilse veya 15 basamaktan b�y�kse hata
    YAZIYLA = "#HATA!"
    Exit Function
End If

If sayi < 0 Then
  negatif = True 'Say� negatif
  sayi = Abs(sayi)
End If

' Birler basama��
birler(0) = ""
birler(1) = "Bir"
birler(2) = "�ki"
birler(3) = "��"
birler(4) = "D�rt"
birler(5) = "Be�"
birler(6) = "Alt�"
birler(7) = "Yedi"
birler(8) = "Sekiz"
birler(9) = "Dokuz"

' Onlar basama��
onlar(0) = ""
onlar(1) = "On"
onlar(2) = "Yirmi"
onlar(3) = "Otuz"
onlar(4) = "K�rk"
onlar(5) = "Elli"
onlar(6) = "Altm��"
onlar(7) = "Yetmi�"
onlar(8) = "Seksen"
onlar(9) = "Doksan"

' B�y�k say�lar
buyukSayi(0) = "Trilyon "
buyukSayi(1) = "Milyar "
buyukSayi(2) = "Milyon "
buyukSayi(3) = "Bin "
buyukSayi(4) = ""

sayiMetni = Right(String(15, "0") & CStr(Fix(sayi)), 15) ' Say�y� metne �evir ve bo�luklar� kald�r

' 1'den 15'e kadar d�ng�
' karakterleri tek tek al ve say�ya �evir ve diziye aktar
For i = 1 To 15
    basamak(i) = CByte(Mid(sayiMetni, i, 1))
Next

sonuc = "" ' Sonu� metni

' say� metnini 3'erli 5 gruba ay�r ve her grubu yaz�ya �evir
For i = 0 To 4
  For j = 1 To 3 'gruptaki y�zler, onlar, birler basamaklar�
    grup(j) = basamak((i * 3) + j)
  Next
  
  Select Case grup(1) ' Y�zler basama��
    Case 0 ' s�f�r ise
      grupMetni = "" ' Y�zler basama�� metni bo�
    Case 1 ' 1 ise
      grupMetni = "Y�z" ' Y�zler basama�� metni "Y�z"
    Case Else ' 2-9 aras� ise
      grupMetni = birler(grup(1)) & "Y�z" ' Y�zler basama�� metni "�kiY�z", "��Y�z" vb.
  End Select
  
  grupMetni = grupMetni & onlar(grup(2)) & birler(grup(3)) ' Onlar ve birler basama��n� ekle
  
  If grupMetni <> "" Then
    grupMetni = grupMetni & buyukSayi(i) ' B�y�k say�lar� ekle
    If (i = 3) And (grupMetni = "BirBin ") Then
      grupMetni = "Bin" ' "BirBin" durumunu d�zelt
    End If
  End If
  sonuc = sonuc & grupMetni ' Sonucu birle�tir
Next
sonuc = Trim(sonuc)
If sonuc = "" Then
  sonuc = "S�f�r"
ElseIf negatif Then
  sonuc = "Eksi " & sonuc
End If
YAZIYLA = sonuc
End Function

