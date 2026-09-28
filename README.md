# BRIPEDIA Report Generator — Streamlit

## Flow user

1. Upload `data_uker.xlsx`
2. Upload `data_hc.xlsx`
3. Upload `data_bripedia.xlsx`
4. Upload `data_historis.xlsx`
5. Set `current_year`
6. Set `tanggal_data_hc`
7. Set `tanggal report`
8. Pilih `regional_office`
9. Klik **Process**
10. Output:
   - 1 ZIP berisi report Excel utama + chart PNG
   - 1 Excel `output_historis`

## Struktur input

- **data_uker**: harus punya sheet `Data Uker`; dibaca dengan `header=1` seperti notebook.
- **data_hc**: sheet pertama; kolom utama mengikuti notebook (`PERNR`, `COMPLETENAME`, `KODE BRANCH`, dst.).
- **data_bripedia**: sheet pertama; membutuhkan `Personal Number`, `Jumlah Akses`, `Jumlah Ketentuan`.
- **data_historis**: sheet pertama; membutuhkan kolom `date` dan kolom Regional Office yang akan dipilih.

## Menjalankan lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Catatan download otomatis

Aplikasi melakukan best-effort untuk memicu dua download otomatis setelah proses selesai. Beberapa browser memblokir multiple automatic downloads; karena itu dua tombol download tetap disediakan sebagai fallback.
