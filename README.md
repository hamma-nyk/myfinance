# MyFinance - Bot Telegram Catatan Finansial

Bot Telegram untuk mencatat pemasukan & pengeluaran. Data disimpan di Google Sheets,
otomatis dipisah per tahun, dengan laporan berwarna dan export `.xlsx` lewat Telegram.

## Fitur

- Catat pemasukan / pengeluaran dari chat Telegram
- Data otomatis dipisah per sheet tahun (`2026`, `2027`, dst)
- Sheet `Laporan` otomatis menampilkan pemasukan (kiri) & pengeluaran (kanan),
  lengkap dengan total dan selisih
- Warna otomatis: header biru, baris pemasukan hijau, pengeluaran merah
- Auto-filter di setiap sheet data (filter tanggal/jam/bulan/tahun)
- Export laporan ke `.xlsx` dan kirim langsung ke Telegram
- Hanya bisa dipakai oleh pemilik bot (`ALLOWED_CHAT_ID`)
- Bisa jalan lokal (polling) atau di Vercel (webhook)

## Perintah Bot

| Perintah              | Fungsi                                        |
| --------------------- | --------------------------------------------- |
| `+50000 gaji`         | Catat pemasukan                               |
| `-20000 kopi`         | Catat pengeluaran                             |
| `50000 freelance`     | Pemasukan (tanpa tanda)                       |
| `/ringkasan`          | Ringkasan hari ini                            |
| `/bulan`              | Ringkasan bulan ini                           |
| `/laporan`            | Laporan bulan ini (detail di sheet `Laporan`) |
| `/laporan 21/10/2026` | Laporan per tanggal                           |
| `/laporan 10/2026`    | Laporan per bulan                             |
| `/laporan 2026`       | Laporan per tahun                             |
| `/export`             | Export bulan ini ke `.xlsx`                   |
| `/export 10/2026`     | Export bulan tertentu                         |
| `/export 2026`        | Export setahun                                |
| `/riwayat`            | 10 transaksi terakhir                         |
| `/hapus`              | Hapus transaksi terakhir                      |
| `/help`               | Bantuan                                       |

## Setup

### 1. Buat Telegram Bot

- Chat [@BotFather](https://t.me/BotFather) di Telegram
- Kirim `/newbot`, ikuti instruksi, dapatkan **BOT_TOKEN**
- Ambil Telegram ID kamu dari [@userinfobot](https://t.me/userinfobot) untuk `ALLOWED_CHAT_ID`

### 2. Setup Google Sheets

1. Buka [Google Cloud Console](https://console.cloud.google.com/)
2. Buat project baru → aktifkan **Google Sheets API**
3. Buat **Service Account** → download JSON key
4. Dari JSON key, ambil `client_email` dan `private_key`
5. Buat Google Spreadsheet baru
6. **Share** spreadsheet ke email service account (Editor)
7. Ambil Spreadsheet ID dari URL: `https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit`

Sheet data & laporan dibuat otomatis oleh bot — tidak perlu bikin manual.

### 3. Jalankan Lokal (polling)

Buat file `.env.local`:

```env
BOT_TOKEN=token_dari_botfather
SPREADSHEET_ID=id_spreadsheet
GOOGLE_SERVICE_ACCOUNT_EMAIL=xxx@xxx.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
ALLOWED_CHAT_ID=telegram_id_kamu
```

Lalu:

```bash
npm install
npm run dev
```

Mode lokal memakai polling dan otomatis menghapus webhook.

### 4. Deploy ke Vercel (webhook)

```bash
npm i -g vercel
vercel
```

Set environment variables di Vercel Dashboard → Settings → Environment Variables:

- `BOT_TOKEN`
- `SPREADSHEET_ID`
- `GOOGLE_SERVICE_ACCOUNT_EMAIL`
- `GOOGLE_PRIVATE_KEY`
- `ALLOWED_CHAT_ID`

Lalu daftarkan webhook (buka di browser):

```
https://api.telegram.org/bot{BOT_TOKEN}/setWebhook?url=https://{YOUR_VERCEL_DOMAIN}/api/webhook
```

> Catatan: `GOOGLE_PRIVATE_KEY` sebaiknya diisi dengan `\n` literal (satu baris),
> kode akan mengubahnya menjadi newline.

## Struktur File

```
myfinance/
├── api/webhook.js   # Handler serverless untuk Vercel
├── dev.js           # Runner lokal (polling)
├── vercel.json
├── .env.example
└── README.md
```

## Struktur Data di Sheets

Sheet tahunan (`2026`, `2027`, ...):

| A       | B   | C     | D     | E    | F         | G      |
| ------- | --- | ----- | ----- | ---- | --------- | ------ |
| Tanggal | Jam | Bulan | Tahun | Tipe | Deskripsi | Jumlah |

Sheet `Laporan`:

- Baris 1: judul laporan
- Baris 3: header tabel (Pemasukan di kiri, Pengeluaran di kanan)
- Di bawah tabel: `TOTAL PEMASUKAN`, `TOTAL PENGELUARAN`, dan `SELISIH`
- Warna: hijau untuk pemasukan/selisih positif, merah untuk pengeluaran/selisih negatif

## Catatan Keamanan

- Jangan commit `.env.local` atau file JSON service account (sudah masuk `.gitignore`)
- Jika token bot pernah bocor, revoke lewat `/revoke` di [@BotFather](https://t.me/BotFather)
- `ALLOWED_CHAT_ID` membatasi bot hanya untuk ID Telegram tertentu
