
"""
update_portfolio_yfinance.py
Auto-update Template Analisa Portfolio dengan yfinance

CARA PAKAI (paling gampang):
  1. pip install yfinance openpyxl pandas
  2. python update_portfolio_yfinance.py
  3. Buka file hasil: Template_Analisa_Portfolio_YFinance_UPDATE.xlsx

Opsi:
  python update_portfolio_yfinance.py --file portfolio.xlsx --period 1y
  python update_portfolio_yfinance.py --file portofolio.xlsx --period 6mo --no-history
  python update_portfolio_yfinance.py --dry-run   (cek harga tanpa simpan)

CARA KERJA:
  - Baca kolom C di sheet Holdings = Ticker Yahoo (contoh: BBCA.JK, BTC-USD, GC=F)
  - Tarik harga terakhir via yfinance
  - Konversi ke IDR otomatis:
      * Saham .JK  -> sudah IDR, kurs = 1
      * Saham US / Crypto BTC-USD -> USD * kurs USDIDR=X
      * Emas GC=F -> USD per troy oz -> IDR per gram = price * USDIDR / 31.1035
  - Update kolom I (Harga Saat Ini), T (Last Update), U (Change 1D %)
  - Update sheet Harga_Historis dengan data harian (untuk Volatilitas, Drawdown, Chart)
  - Update sheet Config -> kurs USDIDR terbaru + waktu update terakhir
  - Semua rumus di Dashboard / Alokasi / Performa langsung ikut ke-update karena mereka ngambil dari Holdings

TICKER YANG DIDUKUNG:
  Saham IDX : BBCA.JK, BMRI.JK, ASII.JK, TLKM.JK, ICBP.JK, ANTM.JK, BBRI.JK, dll (tambah .JK)
  Saham US  : AAPL, MSFT, NVDA, BBRI? no - pakai tanpa suffix
  Crypto    : BTC-USD, ETH-USD, SOL-USD
  Emas      : GC=F (futures emas) -> dikonversi ke per gram IDR
  Forex     : USDIDR=X (kurs), EURIDR=X
  Reksa Dana / MANUAL : isi Ticker = MANUAL, script akan SKIP dan harga tetap manual
"""

import argparse
import sys
import datetime
import time

# --- Cek dependency ---
try:
    import yfinance as yf
except ImportError:
    print("❌ yfinance belum terinstall.")
    print("   Jalankan: pip install yfinance openpyxl pandas")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
except ImportError:
    print("❌ openpyxl belum terinstall.")
    print("   Jalankan: pip install openpyxl")
    sys.exit(1)

# ============================================================
# KONFIGURASI
# ============================================================
# Mapping khusus yang butuh konversi beda
# Kalau ticker kamu tidak ada di sini, script akan deteksi otomatis dari suffix-nya
SPECIAL_TICKERS = {
    "GC=F": {"type": "emas", "desc": "Emas futures USD/oz -> IDR/gram"},
    "SI=F": {"type": "perak", "desc": "Perak futures USD/oz -> IDR/gram"},
}

# Emas Antam biasanya ada premium 5-10% di atas harga spot.
# Ubah kalau mau lebih akurat dengan harga Antam. 1.0 = tanpa premium
EMAS_PREMIUM = 1.05  # 5% premium Antam
GRAM_PER_OZ = 31.1035

# ============================================================
# HELPER
# ============================================================
def get_usd_idr():
    """Ambil kurs USD/IDR terbaru dari yfinance."""
    try:
        t = yf.Ticker("USDIDR=X")
        hist = t.history(period="1d")
        if hist.empty:
            print("⚠️  Gagal ambil USDIDR, pakai fallback 16250")
            return 16250.0, "fallback"
        price = float(hist["Close"].iloc[-1])
        print(f"💵 Kurs USDIDR=X : Rp {price:,.0f}")
        return price, "yfinance"
    except Exception as e:
        print(f"⚠️  Gagal ambil kurs USDIDR ({e}), pakai 16250")
        return 16250.0, "fallback"

def get_price_and_history(ticker_symbol, period="1y"):
    """
    Return: dict {price, currency, change_1d, last_update, hist_df}
    hist_df = DataFrame harian dengan kolom Close
    """
    ticker = yf.Ticker(ticker_symbol)
    # Harga terbaru + history 1y sekaligus (1 call, lebih cepat)
    hist = ticker.history(period=period, auto_adjust=False)
    if hist.empty:
        raise ValueError(f"Data kosong untuk {ticker_symbol} - cek ticker-nya")

    last_close = float(hist["Close"].iloc[-1])
    # Change 1D
    if len(hist) >= 2:
        prev_close = float(hist["Close"].iloc[-2])
        change = (last_close - prev_close) / prev_close if prev_close else 0
    else:
        change = 0

    # Info currency (kadang kosong, kita tebak dari suffix)
    try:
        info = ticker.info
        currency = info.get("currency", "")
    except Exception:
        currency = ""

    # Tebak currency kalau info kosong
    if not currency:
        if ticker_symbol.endswith(".JK"):
            currency = "IDR"
        elif ticker_symbol.endswith("=X"):
            currency = "IDR"
        elif "-USD" in ticker_symbol or ticker_symbol in ["GC=F", "SI=F"]:
            currency = "USD"
        else:
            currency = "USD"

    last_date = hist.index[-1]
    if hasattr(last_date, "strftime"):
        last_update = last_date.strftime("%Y-%m-%d %H:%M")
    else:
        last_update = str(last_date)

    return {
        "price": last_close,
        "currency": currency,
        "change_1d": change,
        "last_update": last_update,
        "hist": hist,
        "last_date": last_date,
    }

def convert_to_idr(ticker_symbol, price_original, currency, usd_idr, kelas_aset=""):
    """
    Konversi harga asli ke IDR.
    Return: (harga_idr, kurs_yang_dipakai, keterangan)
    """
    t = ticker_symbol.strip().upper()
    # 1. Emas / Perak futures
    if t == "GC=F":
        harga_per_gram = price_original * usd_idr / GRAM_PER_OZ * EMAS_PREMIUM
        return harga_per_gram, usd_idr, f"GC=F ${price_original:,.2f}/oz * USDIDR / {GRAM_PER_OZ} * premium {EMAS_PREMIUM}"
    if t == "SI=F":
        harga_per_gram = price_original * usd_idr / GRAM_PER_OZ
        return harga_per_gram, usd_idr, "Perak futures"

    # 2. Sudah IDR (saham IDX)
    if t.endswith(".JK") or currency == "IDR":
        return price_original, 1, "Sudah IDR"

    # 3. USD -> IDR (saham US, crypto)
    if currency == "USD":
        return price_original * usd_idr, usd_idr, f"USD {price_original:,.2f} * USDIDR {usd_idr:,.0f}"

    # 4. Default: anggap USD
    return price_original * usd_idr, usd_idr, f"Default USD * USDIDR"

def format_rp(n):
    return f"Rp {n:,.0f}"

# ============================================================
# UPDATE HOLDINGS
# ============================================================
def update_holdings(wb, usd_idr, period="1y", dry_run=False):
    ws = wb["Holdings"]
    # Cari header Ticker Yahoo (kolom C di template v2)
    # Struktur v2: A No, B Kode, C Ticker Yahoo, D Nama, E Kelas, F Sektor, G Unit, H Beli Avg, I Harga Saat Ini, J Modal, K Nilai Pasar...
    # Kita cari baris header dengan "Ticker" untuk aman
    header_row = 4
    ticker_col = 3  # C
    kode_col = 2    # B
    harga_col = 9   # I = Harga Saat Ini (IDR)
    last_update_col = 20  # T
    change_col = 21       # U
    kurs_col = None  # kita simpan kurs di kolom tersembunyi? pakai Config saja

    print("\n=== UPDATE HOLDINGS ===")
    updated = []
    skipped = []
    failed = []

    # Baca semua baris holdings (5 sampai 30, skip kosong)
    for row in range(5, 35):
        kode = ws.cell(row=row, column=kode_col).value
        ticker = ws.cell(row=row, column=ticker_col).value
        if not kode or not str(kode).strip():
            continue
        kode = str(kode).strip()
        ticker = str(ticker).strip() if ticker else ""

        if not ticker or ticker.upper() in ["", "MANUAL", "-", "N/A"]:
            print(f"⏭️  {kode:6} | SKIP (MANUAL) - harga tetap manual")
            skipped.append(kode)
            continue

        print(f"🔄 {kode:6} | {ticker:12} | fetching...", end=" ")
        try:
            data = get_price_and_history(ticker, period="5d")  # untuk harga terbaru cukup 5d, lebih cepat
            price_idr, kurs, ket = convert_to_idr(ticker, data["price"], data["currency"], usd_idr)
            print(f"-> {data['currency']} {data['price']:,.2f} => {format_rp(price_idr)} | Change {data['change_1d']*100:+.2f}% | {ket}")

            if not dry_run:
                ws.cell(row=row, column=harga_col, value=round(price_idr, 2))
                ws.cell(row=row, column=last_update_col, value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
                ws.cell(row=row, column=change_col, value=data["change_1d"])
            updated.append((kode, ticker, price_idr, data["change_1d"]))
            time.sleep(0.3)  # jeda biar tidak kena rate limit yfinance
        except Exception as e:
            print(f"❌ GAGAL: {e}")
            failed.append((kode, ticker, str(e)))

    return updated, skipped, failed

# ============================================================
# UPDATE HARGA HISTORIS
# ============================================================
def update_history(wb, usd_idr, period="1y", dry_run=False):
    """
    Isi sheet Harga_Historis dengan data harian per ticker.
    Format: Tanggal | Kode | Ticker | Harga Close (Original) | Kurs | Harga Close IDR | Return Harian
    """
    if "Harga_Historis" not in wb.sheetnames:
        print("⚠️  Sheet Harga_Historis tidak ditemukan, skip")
        return

    ws_hold = wb["Holdings"]
    ws_hist = wb["Harga_Historis"]

    # Kumpulkan ticker unik dari Holdings
    tickers = []
    for row in range(5, 35):
        kode = ws_hold.cell(row=row, column=2).value
        ticker = ws_hold.cell(row=row, column=3).value
        if kode and ticker and str(ticker).strip().upper() not in ["", "MANUAL", "-", "N/A"]:
            tickers.append((str(kode).strip(), str(ticker).strip()))

    if not tickers:
        print("⚠️  Tidak ada ticker untuk history")
        return

    print(f"\n=== UPDATE HARGA HISTORIS ({period}) untuk {len(tickers)} ticker ===")
    if dry_run:
        print("   (dry-run: tidak menulis ke Excel)")
        return

    # Clear data lama (baris 4 ke bawah), tapi pertahankan header di baris 3
    ws_hist.delete_rows(4, ws_hist.max_row)

    # Header pastikan benar
    headers = ["Tanggal", "Kode", "Ticker", "Harga Close (Original)", "Kurs (IDR)", "Harga Close IDR", "Return Harian"]
    for c, h in enumerate(headers, start=1):
        cell = ws_hist.cell(row=3, column=c, value=h)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0F2A44")
        cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)

    out_row = 4
    for kode, ticker in tickers:
        print(f"📈 {kode:6} | {ticker:12} | history {period}...", end=" ")
        try:
            data = get_price_and_history(ticker, period=period)
            hist = data["hist"]
            print(f"-> {len(hist)} baris, {hist.index[0].date()} s/d {hist.index[-1].date()}")
            prev_idr = None
            for dt, row_data in hist.iterrows():
                close_orig = float(row_data["Close"])
                harga_idr, kurs, _ = convert_to_idr(ticker, close_orig, data["currency"], usd_idr)
                # Tulis
                ws_hist.cell(row=out_row, column=1, value=dt.strftime("%Y-%m-%d"))
                ws_hist.cell(row=out_row, column=2, value=kode)
                ws_hist.cell(row=out_row, column=3, value=ticker)
                ws_hist.cell(row=out_row, column=4, value=round(close_orig, 2))
                ws_hist.cell(row=out_row, column=5, value=kurs)
                ws_hist.cell(row=out_row, column=6, value=round(harga_idr, 2))
                if prev_idr is None or prev_idr == 0:
                    ws_hist.cell(row=out_row, column=7, value=0)
                else:
                    ws_hist.cell(row=out_row, column=7, value=(harga_idr - prev_idr) / prev_idr)
                # Format
                ws_hist.cell(row=out_row, column=1).number_format = 'yyyy-mm-dd'
                ws_hist.cell(row=out_row, column=4).number_format = '#,##0.00'
                ws_hist.cell(row=out_row, column=5).number_format = '#,##0'
                ws_hist.cell(row=out_row, column=6).number_format = '#,##0'
                ws_hist.cell(row=out_row, column=7).number_format = '0.00%'
                prev_idr = harga_idr
                out_row += 1
            time.sleep(0.4)
        except Exception as e:
            print(f"❌ Gagal history {kode} ({ticker}): {e}")
            continue

    # Auto filter & freeze
    ws_hist.auto_filter.ref = f"A3:G{out_row-1}"
    ws_hist.freeze_panes = "A4"
    print(f"✅ Harga_Historis terisi {out_row-4} baris")

# ============================================================
# UPDATE CONFIG & DASHBOARD TIMESTAMP
# ============================================================
def update_config(wb, usd_idr, updated_count):
    if "Config" in wb.sheetnames:
        ws = wb["Config"]
        # Cari cell Kurs USDIDR - di template v2 ada di B2
        # Kita update B2 dan B7 (Last Update)
        try:
            ws["C2"].value = usd_idr
            ws["C2"].number_format = '#,##0'
            ws["C7"].value = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws["C8"].value = f"{updated_count} ticker berhasil di-update via yfinance"
        except Exception as e:
            print(f"⚠️  Gagal update Config: {e}")

    if "Dashboard" in wb.sheetnames:
        ws = wb["Dashboard"]
        # Dashboard B3 subtitle kita update dengan timestamp - cari cell yang berisi "Last Update"
        # Di template v2, H5 = Last Update value (kita set di kode Excel)
        try:
            # Cari di Dashboard cell yang labelnya "LAST UPDATE"
            for row in ws.iter_rows(min_row=1, max_row=30, max_col=10):
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and "LAST UPDATE" in cell.value.upper():
                        # value di sebelahnya
                        ws.cell(row=cell.row, column=cell.column+1, value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        except Exception:
            pass

# ============================================================
# MAIN
# ============================================================

def find_excel_file(requested):
    """Cari file Excel kalau tidak ketemu di folder saat ini."""
    import os, glob
    # 1. Kalau persis ada, pakai itu
    if os.path.exists(requested):
        return requested
    print(f"\n🔍 File '{requested}' tidak ketemu di folder ini. Mencari di lokasi lain...")
    # 2. Cari dengan pola mirip di folder saat ini (atasi nama jadi (1), dll)
    candidates = []
    # Pola: Template*.xlsx, *Portfolio*.xlsx di cwd
    for pat in ["portfolio.xlsx", "Template*.xlsx", "*Portfolio*.xlsx", "*Analisa*.xlsx", "*.xlsx"]:
        candidates.extend(glob.glob(pat))
    # 3. Cari di Downloads dan Desktop
    home = os.path.expanduser("~")
    for base in [os.path.join(home, "Downloads"), os.path.join(home, "Desktop"), os.path.join(home, "Desktop", "analisis portfolio")]:
        if os.path.isdir(base):
            for pat in ["Template*.xlsx", "*Portfolio*.xlsx", "*Analisa*.xlsx"]:
                candidates.extend(glob.glob(os.path.join(base, pat)))
    # Bersihkan duplikat & file backup / update
    seen = set()
    uniq = []
    for c in candidates:
        if c not in seen and "_BACKUP" not in c and "$" not in c:
            seen.add(c)
            uniq.append(c)
    if not uniq:
        return None
    # Kalau ada 1, langsung pakai
    if len(uniq) == 1:
        print(f"✅ Ketemu di: {uniq[0]}")
        print(f"   Memakai file itu. Biar aman, pindahkan ke folder ini nanti.")
        return uniq[0]
    # Kalau banyak, tampilkan dan pakai yang terbaru
    print("   Ditemukan beberapa file Excel:")
    uniq_sorted = sorted(uniq, key=lambda x: os.path.getmtime(x), reverse=True)
    for i, f in enumerate(uniq_sorted, start=1):
        print(f"   {i}. {f}  (terbaru)" if i==1 else f"   {i}. {f}")
    print(f"\n✅ Memakai yang terbaru: {uniq_sorted[0]}")
    print(f"   Kalau mau pakai yang lain, jalankan: python update_portfolio_yfinance.py --file \"{uniq_sorted[1] if len(uniq_sorted)>1 else uniq_sorted[0]}\"")
    return uniq_sorted[0]

def main():

    parser = argparse.ArgumentParser(description="Auto-update Portfolio Excel dengan yfinance")
    parser.add_argument("--file", default="portfolio.xlsx", help="File Excel yang mau di-update")
    parser.add_argument("--period", default="1y", help="Periode history: 1mo, 3mo, 6mo, 1y, 2y (default 1y)")
    parser.add_argument("--no-history", action="store_true", help="Skip update Harga_Historis (lebih cepat, hanya harga terbaru)")
    parser.add_argument("--dry-run", action="store_true", help="Cek harga saja, tidak simpan")
    parser.add_argument("--output", default="", help="Nama file output, default = *_UPDATE.xlsx")
    args = parser.parse_args()

    print("="*60)
    print("📊 AUTO-UPDATE PORTFOLIO VIA YFINANCE")
    print("="*60)
    print(f"File   : {args.file}")
    print(f"Period : {args.period} | History: {'OFF' if args.no_history else 'ON'} | Dry-run: {args.dry_run}")
    print(f"Waktu  : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Auto-cari file kalau tidak ketemu
    real_file = find_excel_file(args.file)
    if real_file is None:
        print(f"❌ File tidak ditemukan: {args.file}")
        print("")
        print("   SOLUSI PALING CEPAT (pilih salah satu):")
        print("   1. Cek folder Downloads kamu, pindahkan Excel-nya ke folder ini:")
        print("      - Buka File Explorer > Downloads")
        print("      - Cari portfolio.xlsx")
        print("      - Cut + Paste ke: C:\\Users\\Win10\\Desktop\\analisis portfolio")
        print("")
        print("   2. Atau jalankan dengan path lengkap (pakai tanda petik karena ada spasi):")
        print('      python update_portfolio_yfinance.py --file "C:\\Users\\Win10\\Downloads\\portfolio.xlsx"')
        print("")
        print("   3. Atau cek isi folder ini dulu:")
        print("      dir")
        print("      Kalau di list tidak ada file .xlsx, berarti Excel-nya belum didownload / masih di Downloads.")
        sys.exit(1)
    if real_file != args.file:
        print(f"   Menggunakan file: {real_file}")
        args.file = real_file

    try:
        wb = openpyxl.load_workbook(args.file)
    except FileNotFoundError:
        print(f"❌ File tidak ditemukan: {args.file}")
        sys.exit(1)

    if "Holdings" not in wb.sheetnames:
        print("❌ Sheet Holdings tidak ditemukan. Pakai template v2 (YFinance).")
        sys.exit(1)

    # 1. Kurs USDIDR
    usd_idr, sumber = get_usd_idr()

    # 2. Update Holdings (harga terbaru)
    updated, skipped, failed = update_holdings(wb, usd_idr, period=args.period, dry_run=args.dry_run)

    # 3. Update History (kalau tidak di-skip)
    if not args.no_history and not args.dry_run:
        update_history(wb, usd_idr, period=args.period, dry_run=False)
    elif not args.no_history and args.dry_run:
        print("\n(dry-run: skip history)")

    # 4. Update Config & timestamp
    if not args.dry_run:
        update_config(wb, usd_idr, len(updated))

    # 5. Ringkasan
    print("\n" + "="*60)
    print("RINGKASAN")
    print("="*60)
    print(f"✅ Berhasil update : {len(updated)} ticker")
    for kode, ticker, price, chg in updated:
        print(f"   {kode:6} {ticker:12} -> Rp {price:,.0f} ({chg*100:+.2f}%)")
    if skipped:
        print(f"⏭️  Skip MANUAL      : {', '.join(skipped)}")
    if failed:
        print(f"❌ Gagal           : {len(failed)} ticker")
        for kode, ticker, err in failed:
            print(f"   {kode} ({ticker}): {err}")
    print(f"💵 Kurs USDIDR      : Rp {usd_idr:,.0f} ({sumber})")

    if args.dry_run:
        print("\n🔍 Dry-run selesai, file TIDAK disimpan.")
        return

    # 6. Simpan
    if args.output:
        out_file = args.output
    else:
        if "_UPDATE" in args.file:
            out_file = args.file
        else:
            out_file = args.file.replace(".xlsx", "_UPDATE.xlsx")

    # Backup otomatis: simpan juga dengan tanggal
    backup = args.file.replace(".xlsx", f"_BACKUP_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
    try:
        wb.save(backup)
        print(f"\n💾 Backup tersimpan: {backup}")
    except Exception as e:
        print(f"⚠️  Gagal backup: {e}")

    wb.save(out_file)
    print(f"✅ File ter-update tersimpan: {out_file}")
    print("\nBuka file itu, cek sheet Dashboard & Holdings. Semua rumus otomatis ke-update.")
    print("Jalankan lagi besok / tiap minggu untuk update harga terbaru.")

if __name__ == "__main__":
    main()
