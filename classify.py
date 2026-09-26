"""Tag each review with complaint themes using a transparent keyword taxonomy (Indonesian + slang).

A review can carry several themes. Rules were drafted from a read-through of 160 sampled
reviews and then checked against a separate hand-labelled sample (see validate.py).
"""
import re
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent

THEMES = {
    "Balance lost / unauthorised deduction": [
        r"saldo\w* (?:\w+ ){0,3}(?:hilang|lenyap|raib|kosong|habis|berkurang|ke ?potong|terpotong|ketarik|tertarik|kesedot|tersedot)",
        r"(?:hilang|raib|kesedot|tersedot|ketarik|ditilep|dibobol)\w* (?:\w+ ){0,2}(?:saldo|uang|duit|dana)",
        r"uang\w* (?:\w+ ){0,2}(?:hilang|raib|lenyap|kesedot)", r"\bkesedot\b|\btersedot\b|penyedot",
        r"tanpa (?:izin|ijin|persetujuan|konfirmasi|pin)", r"auto ?debit", r"potong\w* (?:otomatis|sendiri)",
        r"tiba[\s-]?tiba (?:\w+ ){0,3}(?:ke ?potong|terpotong|ketarik|tertarik|hilang|berkurang)",
        r"(?:duit|uang|saldo)\w* (?:\w+ ){0,3}(?:ditilep|ilang|kemana|hbs|diembat|ketelen|nilep)", r"\bditilep\b|\bnyolong\b|ilang duit",
        r"\bn?y?edot (?:duit|uang|saldo)|di ?sedot|\bdiembat\b|\bketelen\b|\bnilep\b|double (?:deduction|potong|debit)",
    ],
    "Failed / pending transfer or top-up": [
        # "ga bisa masuk" means cannot log in, so "bisa" is excluded here
        r"\bpending\b", r"(?:tidak|ga|gak|gk|nggak|ngga|belum|belom|balum|blm|tdk) (?!bisa\b)(?:\w+ )?(?:masuk|terkirim|kekirim|sampai|sampe|nambah)",
        r"ke ?tahan|di ?tahan|tertahan|pencairan|pengembalian (?:dana|uang)|uang\w* (?:\w+ )?(?:ga|gak|tidak) balik",
        r"susah (?:transfer|trasfer|tf|top ?up)|(?:proses|diproses) (?:\w+ )?(?:lama|lambat)|lama (?:\w+ )?(?:di ?)?proses|tidak dapa?t (?:saya )?tarik",
        r"(?:diproses|di proses) (?:terus|mulu|lama)", r"\bnyangkut\b", r"transaksi (?:gagal|bermasalah|error|eror)",
        r"(?:gagal|ga bisa|gabisa|gak bisa|tidak bisa|tdk bisa) (?:\w+ )?(?:transfer|tf|top ?up|isi saldo|bayar|kirim|tarik)",
        r"(?:transfer|tf|top ?up|topup) (?:\w+ ){0,3}(?:gagal|lama|lambat|lelet|pending|belum)",
        r"refund", r"dana (?:\w+ )?(?:belum|belom|tidak|ga|gak|nggak) (?:\w+ )?masuk",
        r"(?:pembayaran|payment|transaksi|pencairan|top ?up) (?:\w+ )?(?:gagal|ga berhasil|gak berhasil|tidak berhasil)|sering gagal|gagal terus",
        r"tahan[\s-]?tahan|tidak dikembalikan|ga dikembalikan|gak dikembalikan|lama (?:\w+ ){0,2}masuk",
    ],
    "Login, OTP & verification": [
        r"\blog ?in\b|\blogin\b|\bligin\b|(?:tidak|ga|gak|gabisa|tdk) ?bisa masuk|sandi|password|pin salah|akun\w* (?:\w+ )?keluar",
        r"\botp\b|kode (?:verifikasi|otp)", r"verifikasi|verif\b",
        r"(?:scan|pindai|foto) wajah|face ?(?:id|verif)", r"ganti (?:nomor|no|hp)|nomor (?:lama|baru)",
        r"\bkyc\b|upgrade|premium|full ?service|data diri|\bktp\b", r"terblokir|di ?blokir|\bblokir\b|dibekukan|di ?freeze|dinonaktifkan",
    ],
    "Account security & fraud": [
        r"\bhack|\bdi ?hack|retas|\bbobol|dibajak|phishing", r"akun (?:\w+ )?(?:diambil alih|dicuri)",
        r"penipu(?:an)?|\bscam\b|\bmodus\b|tertipu|ditipu|judol",
    ],
    "Customer service": [
        r"\bcs\b|customer ?serv|call ?center|live ?chat|pusat bantuan|layanan pelanggan",
        r"(?:tidak|ga|gak|gaada|tanpa|tdk|belum) (?:ada )?(?:respon|respons|tanggapan|balasan|jawaban)",
        r"\bbot\b|robot|template|komplain|complain|pengaduan|laporan|\bagen\b|\bagent\b|pelayanan|tanggung ?jawab",
        r"respon\w* (?:\w+ )?(?:lama|lambat|lelet)|slow respon",
        r"(?:ga|gak|tidak|tdk|gabisa|susah) (?:bisa )?(?:di ?respon|dihubungi|di hubungi)|chat admin|\btelp|\btelpon|pertanggung ?jawab",
    ],
    "Fees & charges": [
        # bare "admin" is excluded: reviewers often use it to address support staff
        r"biaya|admin fee|potongan (?:\w+ )?(?:admin|\d)|dipotong \d|\bpajak\b|\bmahal\b|\bcharge\b|\bfee\b",
        # "admin" counts only when tied to an amount or size, e.g. "admin 1k", "admin gede", "+ admin"
        r"admin (?:\w+ )?(?:\d|gede|besar|mahal)|\+ ?admin|potongan (?:\w+ )?(?:tinggi|besar|banyak|gede)|banyak (?:\w+ )?potongan",
    ],
    "Hard to use": [
        r"\bribet\b|\brumit\b|\bruwet\b|\bruet\b|bingung|susah (?:dipakai|digunakan|dimengerti)|user ?friendly",
        r"mempersulit|nyusahin|menyusahkan",
    ],
    "App performance & bugs": [
        r"lemot|lelet|lambat|\blag\b|\bloading\b|\bberat\b|\berror\b|\beror\b|\bbug\b|crash|force close|keluar sendiri",
        r"maintenance|gangguan|tidak bisa dibuka|ga bisa dibuka|gabisa dibuka|\bdelay\b|sering error|update",
        r"\bslow\b|\blolot\b|baterai|battery|(?:ga|gak|tidak|gabisa) (?:bisa )?d(?:i)? ?buka|tidak tersedia|gk tersedia|ga tersedia",
    ],
    "Paylater, loans & debt collection": [
        r"pay ?later|paylater|\blater\b|pinjam|pinjol|\bcicil|kredit|\blimit\b|tagihan|penagih|nagih|debt|denda|spinjam|\butang\b|\bhutang\b",
    ],
    "Ads, notifications & promos": [
        # account/security notifications are not ads, so only promotional notifications count
        r"\biklan|pop ?up|\bspam\b|notif\w* (?:\w+ ){0,2}(?:promo|iklan|tawaran)|matikan notif|promo|cashback|voucher|kupon",
    ],
}
POSITIVE = [r"\bmudah\b|gampang|praktis|cepat|cepet|\bmantap|bagus|baik|\bgood\b|\bbest\b|terbaik|membantu|\baman\b|\bnyaman\b|keren|\bok\b|oke|\bsuka\b|puas|hemat|\bmurah\b"]

COMPILED = {t: re.compile("|".join(p), re.I) for t, p in THEMES.items()}
POS = re.compile("|".join(POSITIVE), re.I)


def tag(text):
    t = " " + re.sub(r"\s+", " ", str(text).lower()) + " "
    return [theme for theme, rx in COMPILED.items() if rx.search(t)]


def main():
    df = pd.read_csv(HERE / "reviews.csv", parse_dates=["date"])
    df["themes"] = df["review"].map(tag)
    df["n_themes"] = df["themes"].str.len()
    df["is_negative"] = df["rating"] <= 2
    df["praise"] = df["review"].str.lower().map(lambda s: bool(POS.search(s)))
    df["themes_str"] = df["themes"].map(lambda x: "; ".join(x))
    df.drop(columns=["themes"]).to_csv(HERE / "reviews_tagged.csv", index=False)

    long = df[["id", "app", "rating", "in_recent", "in_helpful", "is_negative", "themes"]].explode("themes").dropna(subset=["themes"])
    long.rename(columns={"themes": "theme"}).to_csv(HERE / "review_themes_long.csv", index=False)

    neg = df[df.is_negative]
    print(f"negative reviews: {len(neg)} | with at least one theme: {(neg.n_themes > 0).mean():.0%}")
    print(long[long.is_negative].themes.value_counts().to_string())


if __name__ == "__main__":
    main()
