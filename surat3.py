import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import csv
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A4
import tempfile
import os

try:
    import win32print
    import win32api
except ImportError:
    win32print = None
    win32api = None


def simpan_data():
    nama = entry_nama.get()
    kelas = entry_kelas.get()
    hari_tanggal = entry_tanggal.get()
    alasan = entry_alasan.get()
    jam_ke = entry_jam_ke.get()

    if not (nama and kelas and hari_tanggal and alasan and jam_ke):
        messagebox.showerror("Error", "Semua field harus diisi!")
        return

    data = [nama, kelas, hari_tanggal, alasan, jam_ke]

    with open("data_surat_izin.csv", "a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(data)

    tree.insert("", "end", values=data)
    clear_form()

def clear_form():
    entry_nama.delete(0, tk.END)
    entry_kelas.delete(0, tk.END)
    entry_tanggal.set_date('')
    entry_alasan.delete(0, tk.END)
    entry_jam_ke.delete(0, tk.END)

def cetak_surat():
    selected_item = tree.focus()
    if not selected_item:
        messagebox.showwarning("Peringatan", "Pilih salah satu data yang ingin dicetak.")
        return

    data = tree.item(selected_item, "values")

    isi_surat = f"""
==============================================
     SURAT KETERANGAN IZIN KELUAR SEKOLAH
==============================================

    Guru Piket KBM SMKN 1 Probolinggo
    Menerangkan:

    Nama Siswa   : {data[0]}
    Kelas        : {data[1]}

    Ijin Meninggalkan Sekolah pada:
    Hari/Tanggal : {data[2]}
    Alasan       : {data[3]}

    Mohon Diijinkan meninggalkan KBM pada
    Jam ke : {data[4]}

    Mengetahui,
    Guru Piket / Wali Kelas

    ___________________________
    """

    popup = tk.Toplevel()
    popup.title("Preview Surat Izin")
    popup.geometry("500x500")

    text_preview = tk.Text(popup, wrap=tk.WORD)
    text_preview.insert(tk.END, isi_surat.strip())
    text_preview.config(state=tk.DISABLED)
    text_preview.pack(expand=True, fill=tk.BOTH)

    def simpan_ke_pdf():
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if file_path:
            try:
                buat_pdf(isi_surat, file_path)
                messagebox.showinfo("Sukses", f"Surat berhasil disimpan sebagai PDF:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Gagal", f"Gagal menyimpan PDF: {e}")

    def print_langsung():
        if not win32print or not win32api:
            messagebox.showerror("Tidak Didukung", "Fitur cetak langsung hanya tersedia di Windows.")
            return

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w", encoding="utf-8") as tmp_file:
                tmp_file.write(isi_surat.strip())
                tmp_file_path = tmp_file.name

            printer_name = win32print.GetDefaultPrinter()
            win32api.ShellExecute(0, "print", tmp_file_path, None, ".", 0)
        except Exception as e:
            messagebox.showerror("Gagal Cetak", f"Gagal mencetak surat: {e}")

    tk.Button(popup, text="Simpan ke PDF", command=simpan_ke_pdf).pack(pady=5)
    tk.Button(popup, text="Print Langsung", command=print_langsung).pack(pady=5)

def buat_pdf(isi_surat, path):
    c = pdf_canvas.Canvas(path, pagesize=A4)
    width, height = A4
    text = c.beginText(40, height - 50)
    text.setFont("Helvetica", 12)
    for line in isi_surat.strip().splitlines():
        text.textLine(line.strip())
    c.drawText(text)
    c.save()

def muat_data_csv():
    try:
        with open("data_surat_izin.csv", "r", newline="", encoding="utf-8") as file:
            reader = csv.reader(file)
            for row in reader:
                if row:
                    tree.insert("", "end", values=row)
    except FileNotFoundError:
        pass

def hapus_data():
    selected_item = tree.focus()
    if not selected_item:
        messagebox.showwarning("Peringatan", "Pilih data yang ingin dihapus.")
        return

    data_dipilih = tree.item(selected_item, "values")
    confirm = messagebox.askyesno("Konfirmasi", "Apakah Anda yakin ingin menghapus data ini?")
    if confirm:
        tree.delete(selected_item)
        with open("data_surat_izin.csv", "r", newline="", encoding="utf-8") as file:
            semua_data = list(csv.reader(file))
        with open("data_surat_izin.csv", "w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            for row in semua_data:
                if tuple(row) != data_dipilih:
                    writer.writerow(row)

# ==================== GUI ====================

root = tk.Tk()
root.title("Aplikasi Surat Izin / Dispensasi Siswa")
root.geometry("950x650")
root.option_add("*Font", ("Segoe UI", 10))

# --- Form Input ---
frame_form = tk.LabelFrame(root, text="Form Input Surat Izin", padx=10, pady=10)
frame_form.pack(pady=15, padx=15, anchor="w", fill="x")

tk.Label(frame_form, text="Nama Siswa:").grid(row=0, column=0, sticky="w", pady=5)
entry_nama = tk.Entry(frame_form, width=35)
entry_nama.grid(row=0, column=1, pady=5)

tk.Label(frame_form, text="Kelas:").grid(row=1, column=0, sticky="w", pady=5)
entry_kelas = tk.Entry(frame_form, width=35)
entry_kelas.grid(row=1, column=1, pady=5)

tk.Label(frame_form, text="Hari / Tanggal:").grid(row=2, column=0, sticky="w", pady=5)
entry_tanggal = DateEntry(frame_form, width=33, date_pattern='yyyy-mm-dd')
entry_tanggal.grid(row=2, column=1, pady=5)

tk.Label(frame_form, text="Alasan:").grid(row=3, column=0, sticky="w", pady=5)
entry_alasan = tk.Entry(frame_form, width=35)
entry_alasan.grid(row=3, column=1, pady=5)

tk.Label(frame_form, text="Jam ke-:").grid(row=4, column=0, sticky="w", pady=5)
entry_jam_ke = tk.Entry(frame_form, width=35)
entry_jam_ke.grid(row=4, column=1, pady=5)

btn_simpan = tk.Button(frame_form, text="Simpan Surat", command=simpan_data)
btn_simpan.grid(row=5, column=1, pady=10, sticky="e")

# --- Tabel Data ---
tk.Label(root, text="Daftar Surat Izin / Dispensasi", font=("Segoe UI", 11, "bold")).pack()

frame_table = tk.Frame(root)
frame_table.pack(pady=10, padx=10)

columns = ("Nama", "Kelas", "Hari/Tanggal", "Alasan", "Jam Ke")
tree = ttk.Treeview(frame_table, columns=columns, show="headings", height=10)

scrollbar = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=170, anchor="w")

tree.grid(row=0, column=0, sticky="nsew")
scrollbar.grid(row=0, column=1, sticky="ns")

# --- Tombol Aksi ---
frame_buttons = tk.Frame(root)
frame_buttons.pack(pady=10)

btn_cetak = tk.Button(frame_buttons, text="Cetak Surat Terpilih", command=cetak_surat, width=25)
btn_cetak.grid(row=0, column=0, padx=10)

btn_hapus = tk.Button(frame_buttons, text="Hapus Data Terpilih", command=hapus_data, width=25)
btn_hapus.grid(row=0, column=1, padx=10)

# ==================== Load Awal ====================
muat_data_csv()
root.mainloop()
