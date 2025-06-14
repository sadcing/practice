import os
import datetime
import subprocess
from flask import Flask, request, send_file, render_template
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from werkzeug.utils import secure_filename
from PIL import Image

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
        tgl_lahir = datetime.datetime.strptime(data['tanggal_lahir'], "%Y-%m-%d").date()
        data['tanggal_lahir'] = tgl_lahir.strftime("%d/%m/%Y")

        tgl_akta = datetime.datetime.strptime(data['tanggal_akta'], "%Y-%m-%d").date()
        bulan_id = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ]
        data['tanggal_akta'] = f"{tgl_akta.day} {bulan_id[tgl_akta.month - 1]} {tgl_akta.year}"
    except Exception as e:
        return f"❌ Format tanggal salah: {e}"

    template_path = "template.docx"
    if not os.path.exists(template_path):
        return "❌ Template tidak ditemukan."

    try:
        doc = DocxTemplate(template_path)

        # TTD (tanda tangan)
        ttd_file = request.files.get('ttd')
        if ttd_file and ttd_file.filename != '':
            uploads_dir = "uploads"
            os.makedirs(uploads_dir, exist_ok=True)
            filename = secure_filename(ttd_file.filename)
            ttd_path = os.path.join(uploads_dir, filename)
            ttd_file.save(ttd_path)

            img = Image.open(ttd_path)
            if img.width > 500:
                ratio = 500 / float(img.width)
                new_height = int(img.height * ratio)
                img = img.resize((500, new_height), Image.LANCZOS)
                img.save(ttd_path)

            data['ttd'] = InlineImage(doc, ttd_path, width=Mm(35))
        else:
            data['ttd'] = ''

        doc.render(data)

        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)
        baby_name = data.get('baby_name', 'Unknown').strip().replace(' ', '_')
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        path_docx = os.path.join(output_dir, f"Surat_Kelahiran_{baby_name}_{timestamp}.docx")
        path_pdf = path_docx.replace('.docx', '.pdf')

        doc.save(path_docx)

        # LibreOffice Convert
        result = subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf", path_docx, "--outdir", output_dir
        ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if result.returncode != 0:
            return f"❌ Gagal convert PDF: {result.stderr.decode()}"

        if not os.path.exists(path_pdf):
            return "❌ PDF tidak ditemukan setelah convert."

        return send_file(path_pdf, as_attachment=True)

    except Exception as e:
        return f"❌ Gagal proses dokumen: {e}"

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
