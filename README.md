# Employee-Recognition-Innovation-Hub
inovasi buat pegawai berprestasi
# 🩺 HeartCare Hub

**HeartCare Hub** adalah aplikasi web internal rumah sakit/klinik yang dirancang untuk meningkatkan budaya kerja positif, transparansi penilaian kinerja, serta kolaborasi antar dokter, perawat, dan staf medis.

---
## 🔑 Akun Demo untuk Pengujian

Gunakan akun-akun demo di bawah ini untuk mencoba berbagai fitur berdasarkan peran (*role*):

> 💡 **Catatan:** Password untuk **semua akun demo** adalah **`123`**. Jika akun terhapus atau diubah, tekan tombol **"Reset Data Demo"** di halaman login untuk mengembalikannya.

| Nama Lengkap | Username | Role / Jabatan | Hak Akses Utama |
| :--- | :--- | :--- | :--- |
| **William** | `william` | **Admin** | Akses Penuh + Tab *Kelola Admin* (Upload Excel & Hapus Akun) |
| **Dr. Dandi Siregar, Sp.JP** | `dandi` | **Dokter** | Penilaian Kinerja, Kudos, Chat, Inovasi |
| **Dr. Bryan, Sp.JP** | `bryan` | **Dokter** | Penilaian Kinerja, Kudos, Chat, Inovasi |
| **Wulan Margi, S.Kep** | `wulan` | **Perawat** | Penilaian Kinerja, Kudos, Chat, Inovasi |
| **Zhevo Wijaya** | `zhevo` | **Staff** | Penilaian Kinerja, Kudos, Chat, Inovasi |

## ✨ Fitur Utama

- **🔐 Autentikasi & Manajemen Sesi Multi-Akun**
  - Sistem login berbasis role (Admin, Dokter, Perawat, Staff).
  - Fitur *Reset Data Demo* instan untuk pengujian sistem.

- **🏆 Penilaian Kinerja & Leaderboard Real-time**
  - **Top Dokter** dan **Top Perawat & Staff** bulanan berdasarkan akumulasi poin apresiasi.
  - Form evaluasi resmi dengan sistem pemeringkatan bintang (1–5 bintang) dan catatan masukan konstruktif.
  - Umpan riwayat penilaian publik yang transparan antar tim.

- **🎈 Kudos Balon (Papan Apresiasi Publik)**
  - Mengirimkan apresiasi positif ke rekan kerja yang secara otomatis memberikan **+10 Poin**.
  - Papan umpan balik (feed) publik yang interaktif.

- **💡 Usulan Inovasi**
  - Wadah bagi seluruh staf medis untuk mengajukan ide atau solusi perbaikan operasional rumah sakit.

- **💬 Direct Chat Internal Real-time**
  - Fitur percakapan pribadi (messaging) antar pegawai dengan pembaruan instan (*auto-polling*).
  - Notifikasi visual untuk pesan masuk baru.

- **🛡️ Panel Admin IT (Manajemen Data)**
  - **Import & Export Excel**: Dukungan impor masal data pegawai menggunakan template `.xlsx`.
  - Manajemen akun pegawai (pencarian instan dan penghapusan akun).

---

## 🛠️ Teknologi yang Digunakan

- **Frontend:** HTML5, CSS3 (Flexbox & Grid Layout), Pure JavaScript (ES6+)
- **Icons & Fonts:** FontAwesome 6, Google Fonts (*Plus Jakarta Sans*)
- **Libraries:**
  - [SweetAlert2](https://sweetalert2.github.io/) — Pop-up dialog & notifikasi interaktif.
  - [SheetJS (xlsx)](https://sheetjs.com/) — Pengolahan dan impor data dari file Excel.

---

## 🚀 Cara Menjalankan Proyek

1. **Clone Repositori ini:**
   ```bash
   git clone [https://github.com/USERNAME_KAMU/heartcare-hub.git](https://github.com/USERNAME_KAMU/heartcare-hub.git)
