# 🌟 ExcaGlow - Casper Edition

Casper Excalibur serisi dizüstü bilgisayarlar için geliştirilmiş, WMI (Windows Management Instrumentation) tabanlı gelişmiş bir klavye RGB aydınlatma kontrolcüsüdür. 

Varsayılan yazılımın ötesine geçerek ekrandaki renklere duyarlı aydınlatma (Ambiant), akıcı dalga efektleri ve detaylı renk paletleri sunar.

## ✨ Özellikler

* **🌅 Ambiant Modu:** Ekrandaki renkleri gerçek zamanlı analiz ederek klavye ışıklarını ekranla senkronize eder.
* **🎨 3 Bölge Modu:** Klavyenin sol, orta ve sağ bölgelerini ekrandaki ilgili alanların renkleriyle eşleştirir.
* **🌊 Dalga (Wave) ve 🌈 Döngü (Cycle):** Akıcı, özelleştirilebilir dalga paletleri ve RGB döngü efektleri.
* **🫁 Nefes (Breathe) ve 💡 Sabit (Static):** İsteğe bağlı hız ve parlaklık ayarlarıyla standart aydınlatma modları.
* **⚙️ İnce Ayarlar:** Güncelleme hızı (FPS), geçiş yumuşaklığı ve parlaklık ayarları için detaylı kontroller.
* **⬇️ Sistem Tepsisi (Tray) Desteği:** Arka planda sessizce çalışır ve sistem tepsisinden kolayca kontrol edilebilir.

## 🚀 Kurulum (Geliştiriciler İçin)

Python yüklü bir sistemde projeyi çalıştırmak için gerekli kütüphaneleri kurmanız yeterlidir:

1. Depoyu bilgisayarınıza klonlayın.
2. Python 3.11 veya üzeri sürüm kullandığınızdan emin olun.(Önerilen Python Sürümü: 3.11.9)
3. Gerekli paketleri yükleyin:
   ```bash
   pip install WMI mss pillow pystray
