# Proje Metodolojisi: Dinamik Çizelgeleme ve Hızlandırma Kararı

Aşağıdaki metinler, yapılan kod değişikliklerinin teknik karşılığını ve çalışmanın "Yöntem" (Methodology) bölümüne eklenebilecek teknik açıklamaları içermektedir.

## 1. Dinamik Yeniden Çizelgeleme (Dynamic Rescheduling) Yaklaşımı

Geliştirilen optimizasyon modelinde, problemin statik yapısından çıkarılarak gerçek hayat belirsizliklerini kapsayan **dinamik** bir yapıya geçilmiştir. Bu kapsamda **Senaryo Tabanlı Yuvarlanan Ufuk (Rolling Horizon / Scenario-Based)** yaklaşımının basitleştirilmiş bir versiyonu olan "3 Aşamalı Dinamik Çizelgeleme" kurgusu uygulanmıştır:

1.  **Base Plan (Temel Plan - Beklenen Durum):**
    *   İlk aşamada model, tüm hammadde ve yarı mamullerin tedarikçilerden beklenen termin sürelerinde (gecikmesiz) geleceğini varsayarak *ideal* bir üretim çizelgesi oluşturur.
    *   Bu aşama, işletmenin yıl/dönem başında yaptığı "Hedef Planı" temsil eder.

2.  **Disrupted Scenario (Kesintiye Uğramış Durum - Risk Gerçekleşmesi):**
    *   Gerçek hayat verilerinden türetilen olasılık dağılımları (Normal ve Üstel Dağılım) kullanılarak, malzemeler için rassal (stokastik) gecikme süreleri üretilir.
    *   Model, herhangi bir önleyici aksiyon (hızlandırma vb.) alınmadığı takdirde bu gecikmelerin üretim planını nasıl ötelediğini ve müşteri teslim tarihini (Due Date) ne kadar ihlal ettiğini hesaplar.
    *   Bu aşama, "Riskin Gerçekleştiği" ancak henüz "Müdahale Edilmediği" durumu simüle eder.

3.  **Recovery Scenario (İyileştirme/Kurtarma Planı - Yeniden Çizelgeleme):**
    *   Model, gerçekleşen gecikmeleri veri olarak alır ve bu gecikmelerin etkisini minimize etmek için **Hızlandırma (Expedite)** kararlarını devreye sokar.
    *   Bu aşamada modelin karar değişkenleri arasına $E_{i,j}$ (i santrali j malzemesi için hızlandırma kararı: 0 veya 1) eklenir.
    *   Amaç fonksiyonu, sadece gecikme cezasını değil, aynı zamanda hızlandırma maliyetini de içerir. Böylece model, "Gecikmeye katlanmak mı daha az maliyetli, yoksa ekstra kargo parası verip gecikmeyi önlemek mi?" sorusuna matematiksel optimizasyon ile cevap verir.

## 2. Karar Mekanizması ve Hızlandırma (Expedite) Modeli

Modelin "Recovery" aşamasında kullandığı hızlandırma mekanizması şu şekilde kurgulanmıştır:

### Matematiksel İfade
Amaç fonksiyonu (Objective Function) bir minimizasyon problemidir:

$$ Min Z = \sum_{i \in Santraller} (X_i \cdot C_{gecikme} \cdot G_i + Y_i \cdot C_{hizlandirma} \cdot E_{i}) $$

Burada:
*   $G_i$: i santralinin toplam gecikme süresi (gün).
*   $E_{i}$: i santrali için verilen hızlandırma kararları toplamı.
*   $X_i$: Müşteri Memnuniyeti / Zaman katsayısı (SLA skoruna bağlı).
*   $Y_i$: Maliyet katsayısı (Bütçe baskısına bağlı).
*   $X_i + Y_i = 1$: Ağırlıkların normalizasyonu.

### İş Kuralları (Business Rules)
Optimizasyon modeline saf matematiksel maliyetin ötesinde, işletme gerçeklerini yansıtan kısıtlar eklenmiştir:
*   **Kritik Gecikme Sınırı:** Eğer bir malzemenin tedarik gecikmesi **40 günü** aşıyorsa, model maliyete bakmaksızın o malzeme için **zorunlu hızlandırma** (Mandatory Expedite) kararı verir. Bu, "Müşteri kaybı riskinin tolere edilemeyeceği" noktayı temsil eder.

### Çıktıların Analizi
Sonuç raporunda (`rapor.xlsx` - HIZLANDIRMA KARARLARI sayfası), modelin hangi malzeme için neden hızlandırma yaptığı detaylıca sunulmaktadır:
*   Gecikme süresi (rassal üretilen).
*   Seçilen yöntem (örn. Minivan Yükleme).
*   Katlanılan ek maliyet ve kazanılan süre.

Bu yapı, çalışmanızda **"Stokastik Tedarik Süreleri Altında Çok Amaçlı Dinamik Üretim Çizelgeleme"** başlığı altında sunulabilir.
