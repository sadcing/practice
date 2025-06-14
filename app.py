from flask import Flask, request, send_file, render_template
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from werkzeug.utils import secure_filename
from PIL import Image
import os
import datetime
import subprocess

app = Flask(__name__)

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.form.to_dict()

    gender = data.get('kelamin', '').lower()
    if gender == 'laki-laki':
        data['kelamin'] = 'laki-laki / male'
    elif gender == 'perempuan':
        data['kelamin'] = 'perempuan / female'

    hari_mapping = {
        "Senin": "Monday", "Selasa": "Tuesday", "Rabu": "Wednesday",
        "Kamis": "Thursday", "Jumat": "Friday", "Sabtu": "Saturday", "Minggu": "Sunday"
    }
    hari_input = data.get('hari', '')
    data['day'] = hari_mapping.get(hari_input, '')
    data['hari'] = hari_input

    nomor = data.get('nomor_regist', '000')
    kode = data.get('kode_regist', 'X')
    tahun = data.get('tahun_regist', '2025')
    data['regist_number'] = f"{nomor}/{kode}/{tahun}"

    try:
        raw_tanggal_lahir = data['tanggal_lahir']
        tgl_lahir_obj = datetime.datetime.strptime(raw_tanggal_lahir, "%Y-%m-%d").date()
        data['tanggal_lahir'] = tgl_lahir_obj.strftime("%d/%m/%Y")

        raw_tanggal_akta = data.get('tanggal_akta')
        akta_obj = datetime.datetime.strptime(raw_tanggal_akta, "%Y-%m-%d").date()
        bulan_id = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ]
        nama_bulan_akta = bulan_id[akta_obj.month - 1]
        data['tanggal_akta'] = f"{akta_obj.day} {nama_bulan_akta} {akta_obj.year}"
    except Exception as e:
        return f"Format tanggal salah. Error: {e}"

    template_path = "template.docx"
    if not os.path.exists(template_path):
        return "❌ Template tidak ditemukan."

    try:
        doc = DocxTemplate(template_path)

        ttd_file = request.files.get('ttd')
        if ttd_file and ttd_file.filename != '':
            uploads_dir = "uploads"
            os.makedirs(uploads_dir, exist_ok=True)
            filename = secure_filename(ttd_file.filename)
            ttd_path = os.path.join(uploads_dir, filename)
            ttd_file.save(ttd_path)

            img = Image.open(ttd_path)
            max_width = 500
            if img.width > max_width:
                ratio = max_width / float(img.width)
                new_height = int((float(img.height) * float(ratio)))
                img = img.resize((max_width, new_height), Image.LANCZOS)
                img.save(ttd_path)

            data['ttd'] = InlineImage(doc, ttd_path, width=Mm(35))
        else:
            data['ttd'] = ''

        doc.render(data)

        baby_name = data.get('baby_name', 'Unknown').strip().replace(' ', '_')
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)

        filename_docx = f"Surat_Kelahiran_{baby_name}_{timestamp}.docx"
        path_docx = os.path.join(output_dir, filename_docx)
        doc.save(path_docx)

        # Convert DOCX to PDF using LibreOffice
        path_pdf = path_docx.replace(".docx", ".pdf")
        subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf", path_docx, "--outdir", output_dir
        ], check=True)

        return send_file(path_pdf, as_attachment=True)

    except Exception as e:
        return f"❌ Gagal membuat file PDF. Error: {e}"

if __name__ == '__main__':
    app.run(debug=True)