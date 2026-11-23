SAP PO FAST MODE - READ ONLY
============================

AMAÇ
----
Günlük approval kuyruğunu hızlı analiz etmek:

ME28 pending PO
  -> ME23N PO detayları
  -> attachment LISTESİ (dosyayı açmadan)
  -> historical price analizi
  -> description-family price anomaly
  -> PO_Approval_Cockpit_FAST.xlsx

FAST MODE ŞUNLARI YAPMAZ
------------------------
- 05_DOCUMENT_COLLECTOR.vbs çalıştırmaz.
- PDF, Excel, Outlook, Word veya başka attachment dosyası açmaz.
- Attachment kopyalamaz ve parse etmez.
- SAP'ta approve, release, save veya veri değişikliği yapmaz.

PAKET DOSYALARI
---------------
00_RUN_PO_FAST_MODE.vbs          Günlük kullanılacak master
V15_FAST_ANALYZER.py             Analyzer kaynak kodu
V15_FAST_ANALYZER.spec           PyInstaller build tanımı
requirements-fast.txt            Sadece openpyxl + PyInstaller
BUILD_FAST_ANALYZER_EXE.bat      Windows build bilgisayarı için
build-fast-analyzer.yml          GitHub Actions workflow içeriği
SELF_TEST_FAST_ANALYZER.py       Build öncesi otomatik test
DECISION_RULES.md                Karar ve eşleştirme kuralları

HEDEF KLASÖR
------------
Mevcut SAP GUI Scripts klasörünüzde şu dosyalar birlikte olmalı:

Scripts\
  00_RUN_PO_FAST_MODE.vbs
  01_ME28_QUEUE_SCANNER.vbs
  02_ME23N_OPEN_PO.vbs
  03_ME23N_READ_OPEN_PO.vbs
  04_ATTACHMENT_SCANNER_FINAL.vbs
  V15_FAST_ANALYZER.exe
  EXPORT_20261117_082642.xlsx       (veya en yeni EXPORT_*.xlsx)
  Logs\

05_DOCUMENT_COLLECTOR.vbs klasörde kalabilir; FAST master onu çağırmaz.
Eski 00 master da klasörde kalabilir; günlük FAST kullanım için yalnız
00_RUN_PO_FAST_MODE.vbs dosyasına çift tıklayın.

GITHUB ACTIONS İLE EXE BUILD
----------------------------
Private v15-analyzer-build repository içinde:

V15_FAST_ANALYZER.py
V15_FAST_ANALYZER.spec
requirements-fast.txt
SELF_TEST_FAST_ANALYZER.py
.github\workflows\build-fast-analyzer.yml

Bu paketteki build-fast-analyzer.yml dosyasını GitHub'da şu path'e koyun:
.github/workflows/build-fast-analyzer.yml

Actions -> Build V15 FAST Analyzer EXE -> Run workflow.
Tamamlanınca V15_FAST_ANALYZER_WINDOWS artifact'ını indirin.
İçindeki V15_FAST_ANALYZER.exe dosyasını Scripts klasörüne kopyalayın.

NEDEN BUILD DAHA HAFİF
----------------------
FAST analyzer yalnız openpyxl kullanır. pdfplumber, PyMuPDF/fitz, extract-msg,
python-docx, OCR ve Tesseract yoktur. Bu nedenle EXE build'i ve günlük çalışma
document-parsing sürümünden daha basittir.

MASTER DAVRANIŞI
----------------
1. 01 queue scanner güncel pending PO listesini üretir.
2. Her PO için 02 open ve 03 reader çalışır.
3. 04 attachment scanner yalnız attachment adlarını listeler.
4. Attachment satırı yoksa PO teknik olarak FAILED olmaz:
   Attachment Scan Status = NONE yazılır.
5. Bazı PO'larda teknik hata olsa bile okunabilen PO'lar için analyzer çalışır.
   Hatalar cockpit'teki Run Status sekmesinde görünür.
6. Cockpit oluşturulur ve otomatik açılır.

ANALYZER GİRDİLERİ
------------------
Pending_PO_Details.txt           zorunlu
Pending_PO_Attachments.txt       opsiyonel; ad/list only
Fast_Run_Status.txt              master tarafından oluşturulur
EXPORT_*.xlsx                    en yeni dosya otomatik seçilir

ANALYZER ÇIKTILARI
------------------
PO_Approval_Cockpit_FAST.xlsx
V15_FAST_Analysis.json
V15_FAST_Run_Log.txt

COCKPIT SEKME LİSTESİ
---------------------
Dashboard
PO Detail
Attachment Inventory
Historical Matches
Exceptions
Run Status
Parameters

ATTACHMENT YORUMU
-----------------
CS Named = YES yalnızca dosya adının CS/comparative statement gibi göründüğünü
ifade eder. FAST mode dosyayı açmadığı için attachment içeriği, supplier award
veya fiyat eşleşmesi doğrulanmış sayılmaz.

HISTORICAL EŞLEŞTİRME ÖNCELİĞİ
------------------------------
1. Same Material + Supplier + Plant
2. Same Material + Supplier
3. Same Material
4. Same model/part token + Material Group
5. Similar Description + Material Group
6. Çok güçlü Description similarity

Currency veya UoM farklıysa fiyat karşılaştırması yapılmaz. Açık model/part
numaraları farklıysa material code aynı olmadığı sürece eşleşme reddedilir.

THRESHOLD
---------
Historical last-3 median:
  <= +3%      tolerans içinde
  >3% - 5%    CHECK
  >5%         REVIEW
  >10%        REVIEW + justification required

Description-family median:
  <= 10%      consistent
  >10% - 25%  CHECK
  >25%        REVIEW
  >50%        REVIEW + justification required

Service/1 LS itemlarda description tek başına fiyatın makul olduğunu kanıtlamaz.
History yoksa ve named scope/contract/RFQ yoksa REVIEW/HOLD oluşturulur.

GÜNLÜK KULLANIM
---------------
SAP GUI login olduktan sonra:

00_RUN_PO_FAST_MODE.vbs

çalıştırın. FAST cockpit açıldığında REVIEW/HOLD çıkan PO'lar için gerekirse
mevcut deep/document collector akışını ayrıca kullanın.
