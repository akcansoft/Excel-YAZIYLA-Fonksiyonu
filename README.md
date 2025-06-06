# YAZIYLA - Sayıyı Yazıya Çeviren Excel VBA Fonksiyonu  

## ℹ️ Açıklama

`YAZIYLA`, bir sayıyı Türkçe yazıya çeviren, yalnızca tam sayılarla çalışan basit ve etkili bir VBA (Visual Basic for Applications) fonksiyonudur. Excel, Access ve diğer VBA destekli uygulamalarda kullanılabilir. Türkçe Excel 2021'de denenmiştir.
![image](https://github.com/user-attachments/assets/21cadf87-fd54-4488-a085-35317247efcc)


## 💻 YAZIYLA fonksiyonunu kullanma
- `yaziyla.bas` dosyasındaki kodları kopyalayın.
-  Excelde dosya açıkken <kbd>ALT+F11</kbd> tuşlarına basın (yada Şerit menüden **Geliştirici / Visual Basic** tıklayın)
- VBA Editöründe menüden Insert / Module tıklayın
- Menüden Edit / Paste ile ya da Ctrl+V ile kopyalanan kodları yapıştırın.
- Artık YAZIYLA fonksiyonunu diğer excel fonksiyonları gibi hücrelerde kullanabilirsiniz.

## 🔢 Desteklenen Sayı Aralığı

- Bu fonksiyon sayıdaki küsurları dikkate almaz.
- `-922,337,203,685,477` ile `922,337,203,685,477` arasındaki sayılar geçerlidir.
- 15 basamaktan büyük sayılarda ve sayı olmayan değerlerde `#HATA!` sonucunu verir.
  
## 𝄜 Excel Hücrelerinde Kullanımı
 `=YAZIYLA(Sayı ya da Hücre adresi)`
 
### Örnek
```excel
 =YAZIYLA(A1)
 =YAZIYLA(1453)
```

## 💻 YAZIYLA fonksiyonunu tüm excel dosyalarında kullanma.
- Bu işlemler bir defa yapılacaktır.
- Yeni bir excel dosyası oluşturun.
- Üstteki açıklmalarla VBA editörüne kodları ekleyin.
- Dosya / Farklı kaydet'i tıklayın
- Kayıt Türü listesinden "Microsoft Office Excel Eklentisi (*.xla)"  veya (*.xlam) seçin
- Kayıt Yeri'nde "Addins" belirir.
- Kaydet'i tıklayın
- Dosyayı kapatın
- Yeni bir excel dosyası ya da varolan bir excel dosyanızı açın
- Araçlar / Eklentiler'i tıklayın
- Burada "Kullanılabilir eklentiler"de "Yazıyla" göreceksiniz. Yanındaki seçeneği tıklayıp seçin.
- Tamam'ı tıklayın.
- Artık her excel dosyasında YAZIYLA fonksiyonunu başka bir işleme gerek kalmadan rahatlıkla kullanabilirsiniz.

## 🧾 Lisans

Bu proje GPL 3.0 Lisansı altında lisanslanmıştır. Daha fazla bilgi için `LICENSE` dosyasına bakın.

## 🤝 Katkı

Katkılarınız memnuniyetle karşılanır! Eğer özellik eklemek, hataları düzeltmek veya kodu geliştirmek isterseniz, bir çekme isteği açmaktan çekinmeyin.

## ✉️ İletişim

Mesut Akcan\
**Email**: <makcan@gmail.com>\
**Blog**: [akcansoft.blogspot.com](http://akcansoft.blogspot.com) - [mesutakcan.blogspot.com](http://mesutakcan.blogspot.com)\
**GitHub**: [akcansoft](http://github.com/akcansoft)\
**YouTube**: [Mesut Akcan](http://youtube.com/mesutakcan)
