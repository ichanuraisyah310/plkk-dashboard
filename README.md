# Dashboard Utilization Review PLKK (Streamlit)

Dashboard langsung tampil saat dibuka — memakai data bawaan `data/Data_Dummy.xlsx`, tanpa perlu upload file.

## Struktur folder
```
plkk-dashboard/
├── app.py                  # aplikasi Streamlit
├── requirements.txt        # dependensi
├── data/
│   └── Data_Dummy.xlsx     # data bawaan (dibaca otomatis)
├── .streamlit/config.toml  # konfigurasi Streamlit
├── .gitignore
└── README.md
```

## Menjalankan di komputer (VS Code)
```bash
python -m venv .venv
.venv\Scripts\activate          # Windows   (macOS/Linux: source .venv/bin/activate)
pip install -r requirements.txt
streamlit run app.py
```
Browser terbuka di http://localhost:8501 dan dashboard langsung muncul.

## Deploy gratis lewat GitHub + Streamlit Community Cloud
1. Upload seluruh isi folder ini ke repository GitHub (via `git push` atau drag & drop di github.com).
2. Buka https://share.streamlit.io → **New app** → pilih repository, branch `main`, Main file path: `app.py`.
3. Klik **Deploy**. Dashboard akan tampil otomatis dengan data bawaan.

## Mengganti data
- **Permanen:** timpa `data/Data_Dummy.xlsx` dengan file baru (nama file harus tetap sama), lalu commit & push.
- **Sementara:** di sidebar buka **Ganti sumber data (opsional)** lalu upload file `.xlsx`.

## Format file yang diterima
Workbook `.xlsx` dengan sheet **BMIV-01, BMIV-02, BMIV-03, BMIV-04** (boleh berakhiran, mis. `_Dummy`)
dan nama kolom sama persis dengan `Data_Dummy.xlsx`. Sheet lain (README, Column_Mapping, dst.) diabaikan.

> Catatan privasi: bila repository GitHub bersifat publik, data di dalam `data/` ikut terlihat publik.
> Untuk data asli/sensitif, gunakan repository **private** atau jangan commit file datanya.
