from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx2pdf import convert
import os
import uuid

app = Flask(__name__)
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
UPLOAD_FOLDER = "uploads"
MODEL_PATH = "models/rapport_template.docx"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Gestion de l'image (signature)
        image_path = ""
        if "image_signature" in request.files:
            image_file = request.files["image_signature"]
            if image_file.filename:
                image_name = f"signature_{uuid.uuid4().hex}.png"
                image_path = os.path.join(UPLOAD_FOLDER, image_name)
                image_file.save(image_path)

        doc = DocxTemplate(MODEL_PATH)
        image_signature = InlineImage(doc, image_path, width=Mm(50)) if image_path else ""

        # Récupération des données du formulaire
        context = {
            "client_ste": request.form.get("client_ste"),
            "client_address": request.form.get("client_address"),
            "version": request.form.get("version"),
            "client_name": request.form.get("client_name"),
            "client_mail": request.form.get("client_mail"),
            "client_phone": request.form.get("client_phone"),
            "syn_nbr": request.form.get("syn_nbr"),
            "syn_p_totale": request.form.get("syn_p_totale"),
            "niveau_eclairage": request.form.get("niveau_eclairage"),
            "facteur_uniformité": request.form.get("facteur_uniformité"),
            "puissance_init": request.form.get("puissance_init"),
            "puissance_projetée": request.form.get("puissance_projetée"),
            "puissance_réelle_projetée": request.form.get("puissance_réelle_projetée"),
            "conso_initiale": request.form.get("conso_initiale"),
            "conso_projetée": request.form.get("conso_projetée"),
            "economie_energie": request.form.get("economie_energie"),
            "emissions": request.form.get("emissions"),
            "ste": request.form.get("ste"),
            "address": request.form.get("address"),
            "surface": request.form.get("surface"),
            "activité": request.form.get("activité"),
            "nbr_batiments": request.form.get("nbr_batiments"),
            "date_visite": request.form.get("date_visite"),
            "date_etude": request.form.get("date_etude"),
            "audit": request.form.get("audit"),
            "contact": request.form.get("contact"),
            "station_meteo": request.form.get("station_meteo"),
            "nom_client": request.form.get("nom_client"),
            "telephone_client": request.form.get("telephone_client"),
            "p_unitaire": request.form.get("p_unitaire"),
            "nombre": request.form.get("nombre"),
            "p_totale": request.form.get("p_totale"),
            "t_utilisation": request.form.get("t_utilisation"),
            "fontionnement": request.form.get("fontionnement"),
            "w_m2": request.form.get("w_m2"),
            "bâtiments": request.form.get("bâtiments"),
            "secteur_etude": request.form.get("secteur_etude"),
            "seuil_reglementaire": request.form.get("seuil_reglementaire"),
            "puissance_installée": request.form.get("puissance_installée"),
            "consommation_energie": request.form.get("consommation_energie"),
            "image_signature": image_signature
        }

        doc.render(context)

        base_filename = f"rapport_{uuid.uuid4().hex}"
        output_docx = os.path.join(OUTPUT_FOLDER, base_filename + ".docx")
        output_pdf = os.path.join(OUTPUT_FOLDER, base_filename + ".pdf")
        doc.save(output_docx)

        try:
            convert(output_docx)
        except Exception as e:
            return f"Erreur lors de la conversion PDF : {e}"

        return f"""
        <!DOCTYPE html>
        <html lang='fr'>
        <head>
            <meta charset='UTF-8'>
            <title>Rapport généré</title>
            <meta http-equiv='refresh' content='10; URL=/download/{os.path.basename(output_pdf)}'>
            <style>
                body {{ font-family: Arial, sans-serif; text-align: center; padding: 50px; background: #f2f2f2; }}
                h2 {{ color: #2ecc71; }}
                a {{ text-decoration: none; color: #3498db; font-size: 18px; }}
                .message {{ background: white; display: inline-block; padding: 30px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }}
            </style>
        </head>
        <body>
            <div class='message'>
                <h2>✅ Rapport généré avec succès !</h2>
                <p><a href='/download/{os.path.basename(output_pdf)}'>📄 Télécharger le PDF</a></p>
                <p>(Vous allez être redirigé automatiquement dans 10 secondes...)</p>
            </div>
        </body>
        </html>
        """

    return render_template("form.html")

@app.route("/download/<filename>")
def download(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
