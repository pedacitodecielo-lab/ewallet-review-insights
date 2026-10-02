# Materi Reksa Dana Pasar Uang

Paket materi regenerasi divisi tentang Reksa Dana Pasar Uang (RDPU). Data per Oktober 2026.

| File | Isi |
|---|---|
| `Reksa_Dana_Pasar_Uang.pptx` | Presentasi 35 slide. Naskah bicara ada di speaker notes. Navigasi "Daftar isi" dan kuis bisa diklik saat mode slideshow. |
| `Naskah_Presentasi_RDPU.docx` | Naskah slide demi slide, pembagian waktu (sekitar 54 menit), dan jawaban untuk pertanyaan yang sering muncul. |
| `Modul_Belajar_RDPU.docx` | Modul belajar mandiri: konsep, contoh hitung, mitos dan fakta, 10 latihan soal dengan kunci, kamus istilah, sumber. |
| `Simulator_RDPU.html` | Simulator interaktif (buka di browser): tabungan vs deposito vs RDPU, kalkulator unit dan NAB/UP, kuis. |
| `sumber/` | Skrip pembuat file di atas, untuk memperbarui angka. |

Sebelum membawakan ulang materi, perbarui BI Rate, bunga penjaminan LPS, data KSEI, dan kisaran return RDPU. Daftar lengkapnya ada di bagian 16 modul.

Membangun ulang: `npm install pptxgenjs react react-dom react-icons sharp docx`, lalu dari folder `sumber/` jalankan `node build_ppt.js Reksa_Dana_Pasar_Uang.pptx` dan `node build_docs.js`.
