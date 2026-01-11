# HVAC Üretim Çizelgeleme ve Hızlandırma Optimizasyonu

Bu proje, HVAC sektöründe üretim çizelgeleme problemleri için Gurobi optimizasyon kütüphanesi kullanılarak geliştirilmiş dinamik bir karar destek modelidir.

## Proje Amacı

Satınalma süreçlerinde yaşanan tedarikçi kaynaklı belirsizliklerin (stokastik gecikmeler) üretim planı üzerindeki etkilerini analiz etmek ve maliyet-etkin hızlandırma (expedite) kararları vererek teslimat gecikmelerini minimize etmektir.

## Özellikler

*   **Dinamik Çizelgeleme:** Yuvarlanan Ufuk (Rolling Horizon) yaklaşımı ile 3 aşamalı (Base Plan -> Disrupted -> Recovery) simulasyon.
*   **Stokastik Modelleme:** Malzeme gecikmeleri için geçmiş veriye dayalı olasılık dağılımları (Normal, Üstel).
*   **Optimizasyon Modeli:** Gecikme cezası ve hızlandırma maliyeti arasında trade-off yapan çok amaçlı (Maliyet & Müşteri Memnuniyeti) Gurobi modeli.
*   **Otomatik Raporlama:** Sonuçların detaylı Excel raporu olarak çıktılanması (Hızlandırma kararları, Gantt şeması verileri vb.).

## Kurulum ve Çalıştırma

1.  Gerekli kütüphaneleri yükleyin:
    ```bash
    pip install -r requirements.txt
    ```

2.  Lisans:
    *   Bu proje `gurobipy` kullanmaktadır. Çalıştırmak için geçerli bir Gurobi lisansı (akademik veya ticari) gereklidir.

3.  Çalıştırma:
    ```bash
    python gurobi_optimizasyon.py
    ```

## Dosyalar

*   `gurobi_optimizasyon.py`: Ana optimizasyon kodu.
*   `PROJE VERİ.xlsx`: Girdi veri seti (Prosesler, malzemeler, siparişler).
*   `rapor.xlsx`: Modelin ürettiği sonuç raporu.

---
**Not:** Bu proje akademik bir çalışma kapsamında geliştirilmiştir.
