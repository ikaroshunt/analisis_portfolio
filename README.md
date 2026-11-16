# Analisa Portfolio — Auto-Update dengan yfinance

> Update 10 aset portfolio dalam 1 perintah. Tanpa copy-paste manual lagi.

[Python](https://img.shields.io/badge/Python-3.15-blue)
[yfinance](https://img.shields.io/badge/yfinance-live-green)
[Excel](https://img.shields.io/badge/Excel-auto--update-brightgreen)
[Status](https://img.shields.io/badge/Status-VALIDATED-success)

**Artikel lengkap (English) ada di Dev.to & Codecademy Forum.**
**README ini Bahasa Indonesia, biar yang buka repo ikut pusing dikit. 😅**

---

## Kenapa Repo Ini Ada?

Dulu, cara gue update portfolio tuh nyiksa:

Buka Yahoo Finance → Cari BBCA.JK → Copy harga → Paste ke Excel → Ulangi buat BMRI, ASII, TLKM, ICBP, ANTM, BBRI, Emas, BTC, RDPU.

10 aset. 10 tab. Tiap hari.

Sekali typo, ROI langsung salah. Sekali males, portfolio langsung basi.

Akhirnya gue bikin sistem ini: **Excel buat analisa, Python buat narik harga.**

```bash
python update.py
```

Udah. Beres.

## Isi Repo

```
analisis portfolio/
├── portfolio.xlsx   # Template Excel: Dashboard, Holdings, Alokasi, Performa, Harga_Historis
├── update.py        # Script Python: yfinance -> Excel (JANGAN DIUBAH, SUDAH VALID)
├── README.md        # File ini (Bahasa Indonesia)
└── portfolio_UPDATE.xlsx  # Hasil setelah run (auto-generate)
```

> ⚠️ **PENTING:** `portfolio.xlsx` dan `update.py` sudah VALID. Jangan diubah-ubah. Kalau mau eksperimen, copy dulu.

## Fitur

- ✅ **Auto-fetch harga live** dari Yahoo Finance
- ✅ **Auto-konversi ke IDR:**
  - Saham IDX `*.JK` → sudah IDR
  - Saham US & Crypto `BTC-USD` → USD x `USDIDR=X`
  - Emas `GC=F` → USD/oz → IDR/gram (dibagi 31.1035, dikali premium Antam 1.05)
- ✅ **Auto-history 1 tahun** buat hitung Volatilitas & Drawdown
- ✅ **Auto-backup** tiap run: `portfolio_BACKUP_YYYYMMDD_HHMM.xlsx`
- ✅ **Mode dry-run** buat cek harga tanpa nyimpen
- ✅ **Skip MANUAL** buat Reksa Dana yang nggak ada di Yahoo

## Cara Pakai

### 1. Install sekali aja

```bash
pip install yfinance openpyxl pandas
```

Kalau pakai venv, pastikan venv-nya aktif dulu. Prompt-nya ada `(.venv-1)` gitu.

### 2. Taruh 2 file dalam 1 folder

`portfolio.xlsx` + `update.py` harus satu folder. Jangan dipisah.

### 3. Jalankan

```bash
python update.py
```

Tunggu 10-30 detik sampai muncul:

```
✅ File ter-update tersimpan: portfolio_UPDATE.xlsx
```

Buka file itu. Dashboard & Holdings udah ke-update semua.

### Opsi Lain

```bash
python update.py --dry-run       # Cek harga doang, nggak nyimpen
python update.py --no-history    # Cuma harga terbaru, lebih cepat
python update.py --period 6mo    # History 6 bulan aja
```

## Bukti Validasi (Real Run)

Ini hasil `--dry-run` tanggal **16 Nov 2026, 16:33** — ini yang bikin repo ini VALID:

```
Kurs USDIDR=X : Rp 18,250

BBCA  | BBCA.JK  | -> Rp 5,600   | +1.82% | Sudah IDR
BMRI  | BMRI.JK  | -> Rp 3,900   | +0.26% | Sudah IDR
ASII  | ASII.JK  | -> Rp 5,700   | +3.64% | Sudah IDR
TLKM  | TLKM.JK  | -> Rp 2,180   | +0.93% | Sudah IDR
ICBP  | ICBP.JK  | -> Rp 6,825   | +1.87% | Sudah IDR
ANTM  | ANTM.JK  | -> Rp 2,670   | -4.98% | Sudah IDR
BBRI  | BBRI.JK  | -> Rp 3,100   | +1.31% | Sudah IDR
EMAS  | GC=F     | -> Rp 2,380,380 | -0.58% | USD 3,863.70/oz * USDIDR / 31.1035 * 1.05
BTC   | BTC-USD  | -> Rp 1,173,746,540 | -1.52% | USD 64,314.88 * USDIDR
RDPU  | SKIP (MANUAL)

RINGKASAN
Berhasil update : 9 ticker
Skip MANUAL      : RDPU
```

9 dari 9 ticker berhasil. 0 gagal.

Hari itu sahamnya ijo (ASII +3.64%), Emas & BTC merah, ANTM nyungsep sendiri -4.98%.

## Tambah Saham Baru

Di sheet **Holdings**, tambah baris baru, isi 2 kolom ini aja:

| Kode | Ticker Yahoo |
| :--- | :--- |
| GOTO | `GOTO.JK` |
| BREN | `BREN.JK` |
| AMMN | `AMMN.JK` |

Terus jalankan lagi `python update.py`. Harga langsung keisi.

Ticker yang sering dicari: `GOTO.JK`, `BREN.JK`, `AMMN.JK`, `CUAN.JK`, `PGEO.JK`, `MBMA.JK`, `ETH-USD`, `SOL-USD`, `NVDA`, `MSFT`

## Bug Yang Pernah Bikin Pusing

Biar lo nggak ngulangin kesalahan gue:

### 1. `_` vs `-` Beda!

Gue pernah ketik:
```bash
python update_portfolio_yfinance_FIX.py
```
Padahal file-nya `Update-Portofolio-Yfinance-FIX.py` — pakai strip `-`, bukan underscore `_`.

Browser ngubah `_` jadi `-` pas download. Di Windows itu beda file.

**Solusi:** Sekarang namanya udah disimplifikasi jadi `update.py` dan `portfolio.xlsx`. Nggak ada lagi drama strip.

### 2. Tombol Run VS Code Nambah Petik Dobel

Kalau klik tombol Play di VS Code, dia jalaninnya jadi:
```
python.exe '"c:/.../update.py"'
```
Petik dobel → `Errno 22 Invalid argument`.

**Solusi:** Ketik manual di terminal. Jangan pakai tombol Play. Atau ketik `python up` terus pencet **Tab** biar auto-complete.

### 3. Venv Nggak Nemu yfinance

Prompt ada `(.venv-1)` tapi `yfinance` di-install di luar venv → nggak ketemu.

**Solusi:** Install pas venv-nya aktif:
```bash
pip install yfinance openpyxl pandas
```

## Struktur Excel

| Sheet | Isinya |
| :--- | :--- |
| **Dashboard** | Total nilai, Gain/Loss, ROI, CAGR, Sharpe, Top 5 holdings, Last Update |
| **Holdings** | Sumber utama. Kolom C = Ticker Yahoo, Kolom I = Harga Auto-Update |
| **Alokasi** | Bobot aktual vs target + kalkulator rebalancing |
| **Performa** | Return bulanan + metrik risk |
| **Harga_Historis** | Data harian per ticker, diisi otomatis script |
| **Config** | Kurs USDIDR, Risk Free, Premium Emas |

## Roadmap

- [x] Auto-update harga via yfinance
- [x] Konversi IDR otomatis (IDX, US, Crypto, Emas)
- [x] Dry-run & backup otomatis
- [ ] Jadwal otomatis tiap Senin 08:00 (Task Scheduler / cron)
- [ ] Ringkasan Telegram / Email tiap update
- [ ] Tracking dividen otomatis dari yfinance

## Disclaimer

Ini bukan ajakan beli/jual saham. Ini cuma tools buat tracking. Keputusan investasi tetap di tangan lo. DYOR.

---

Dibuat di Windows 10, Python 3.15, yfinance, openpyxl, dan drama `_` vs `-` yang tak terlupakan.

Kalau artikel English-nya mau dibaca, cek di Dev.to & Codecademy Forum. Kalau README ini yang lo baca, ya... selamat pusing. 😅
