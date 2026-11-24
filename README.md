M-Files Toolkit

PowerShell tabanlı, modüler Windows araç seti.

M-Files Toolkit; muhasebe, e-Defter, Excel, belge yönetimi ve çeşitli hesaplama işlemlerini kolaylaştırmak amacıyla geliştirilmiş bir PowerShell uygulamasıdır.

Uygulama, Windows Terminal üzerinde çalışan iki panelli bir arayüze sahiptir. Sol panelden modüller seçilir, sağ panelde seçilen modül çalıştırılır.

Özellikler

İki panelli Windows Terminal arayüzü

Fiziksel klasör yapısından otomatik modül ve grup algılama

Dinamik modül listesi ve arama

Modüler ve genişletilebilir mimari

Modül bazında yapılandırma dosyaları

Excel, PDF ve çeşitli belge işlemleri

Kullanıcı tarafından yeni modül eklenebilmesi

Modül Grupları

Uygulama toplam 23 modül içerir.

Grup

İçerik

Defter Yedekleme Akışı

E-Defter kontrolü, ZIP oluşturma, yerel yedekleme, optik medya

Defter ve GİB

GİB berat okuma

LUCA

LUCA ile ilgili yardımcı araçlar

Hesaplama

Binek oto, inşaat, başabaş, poliçe ve kur farkı hesaplamaları

Belge ve Senet

Senet, PDF, HTML ve karekod araçları

Excel

Excel filtreleme, parçalama, CSV ve dosya listeleme

Gereksinimler

Windows 10 / Windows 11

Windows PowerShell 5.1 veya uyumlu PowerShell ortamı

Windows Terminal (wt.exe)

İlgili modüllerin ihtiyaç duyduğu ek bileşenler

Bazı modüller Excel COM, WinForms veya ek kütüphanelere ihtiyaç duyabilir.

Kurulum ve Başlatma

GitHub deposundan ZIP dosyasını indirin.

Arşivi bilgisayarınızda istediğiniz klasöre çıkarın.

baslat.bat dosyasına çift tıklayın.

Uygulama Windows Terminal üzerinde açılır.

Not: Programın resmi başlangıç dosyası baslat.bat dosyasıdır.

Kullanım

Tuş / İşlem

Açıklama

Yukarı / Aşağı

Modüller arasında gezinme

Enter

Seçili modülü çalıştırma

Numara + Enter

Modülü numarasıyla seçme

Modül adı + Enter

Modül arama ve seçme

Esc

Aramayı temizleme

F5

Modül listesini ve grupları yenileme

Q + Enter

Programdan çıkış

Modüler Mimari

Menü grupları modules klasörünün birinci seviye klasörlerinden otomatik oluşturulur.

Her modül kendi klasöründe bulunur ve module.meta.json dosyasıyla tanımlanır.

Yeni modül eklemek için ilgili grup altında yeni bir modül klasörü oluşturulması ve gerekli modül tanım dosyasının eklenmesi yeterlidir.

Grup sayısı sabit değildir. Fiziksel klasör yapısı, menü organizasyonunun temelini oluşturur.

Kullanıcı Verileri

Kullanıcı ayarları ve yerel çalışma verileri kaynak kodundan ayrı tutulacak şekilde yapılandırılmıştır.

Git deposunda aşağıdaki dosya ve klasörler .gitignore kurallarıyla takip dışında bırakılmıştır:

Modül yapılandırma dosyaları (*_config.json)

E-Defter firma listesi

projeler klasörleri

logs ve raporlar klasörleri

Bu kurallar geliştirme deposunun takibini düzenler. Kullanıcı verilerinin güncelleme sırasında korunması ayrıca güncelleme mekanizmasının sorumluluğundadır.

Geliştirme ve Katkı

Projeye katkıda bulunmak isteyen geliştiriciler:

Depoyu Fork edebilir.

Kendi kopyalarında geliştirme yapabilir.

Hataları Issues bölümünden bildirebilir.

Değişikliklerini Pull Request aracılığıyla önerebilir.

Gönderilen değişiklikler ana projeye eklenmeden önce proje sahibi tarafından incelenir.

Sürüm

Mevcut sürüm: v1.4.7

Sürüm geçmişi ve değişiklikler için Releases bölümünü ziyaret edin.

Geliştirici

Mustafa Orakcı

GitHub: @orakcimustafa9656

Lisans

Bu depoya henüz bir açık kaynak lisansı eklenmemiştir. Kodun yeniden kullanımı, dağıtımı ve değiştirilmesine ilişkin izinler için bir lisans dosyası eklenmesi gerekmektedir.

Veri mimarisi

Program kodu Git'te, kullanıcı verisi Data klasöründe ve Git dışındadır (.gitignore yalnızca /Data/ içerir). Klasörler ilk ihtiyaç anında oluşur; boş klasör açılmaz.

Data\Common\Config\terminal.json: MenuWidth (varsayılan 0.30)
Data\Modules\<id>\Config | Logs | Reports | Projects | Cache: modülün kalıcı verisi
Data\Temp\Update | Sessions | Modules\<id>: geçici dosyalar

Modüllerde yol için Get-MFTModuleDataPath -ModuleId 'id' -Category Config kullanılır. Veri taşımak için Data klasörünü elle kopyalamak yeterlidir.

Modül ekleme, silme, güncelleme

Ekleme: modules\<NN- grup>\<KK- id>\ klasörü açın; içine <id>.ps1 ve module.meta.json koyun. Meta dosyası şu alanları içerir: Id (küçük harf, sabit, veri klasörü adı), Key (menü numarası), Title, Icon. Menüde F5 ile görünür.
Silme: modül klasörünü silin. Data\Modules\<id> kullanıcı verisi olduğu için kendiliğinden silinmez.
Güncelleme: modül kodunu değiştirip version.txt sürümünü artırın, GitHub'da aynı etiketle Release oluşturun. Updater yeni sürümde olmayan dosyaları modules ve lib altından siler.
Grup: modül klasörünün üst klasörüdür; menüdeki grup başlığı klasör adıdır (numara sıralama içindir).

Geliştirici kopyası

Klasörde .git varsa updater hiçbir dosyayı değiştirmez; yalnızca sürüm ve git commit bilgisini açılış ekranında gösterir. GitHub'daki Release eski kalırsa "yerel sürüm ileride" mesajı görürsünüz.

