from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from io import BytesIO
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
        image_path = ""
        if "image_signature" in request.files:
            image_file = request.files["image_signature"]
            if image_file.filename:
                image_name = f"signature_{uuid.uuid4().hex}.png"
                image_path = os.path.join(UPLOAD_FOLDER, image_name)
                image_file.save(image_path)

        doc = DocxTemplate(MODEL_PATH)
        image_signature = InlineImage(doc, image_path, width=Mm(50)) if image_path else ""
        
        # Récupération des luminaires dynamiques
        luminaires = []
        index = 0
        while f"marque_{index}" in request.form:
            marque = request.form.get(f"marque_{index}")
            reference = request.form.get(f"reference_{index}")
            puissance = float(request.form.get(f"puissance_{index}") or 0)
            nombre = float(request.form.get(f"nombre_{index}") or 0)
            puissance_totale = puissance * nombre
            luminaires.append({
                "marque": marque,
                "reference": reference,
                "puissance": puissance,
                "nombre": nombre,
                "puissance_totale": puissance_totale
            })
            index += 1
            
            
        zones_1 = request.form.getlist('zones[]')  # Cela te donne une liste comme ['COMMERCE', 'INDUSTRIEL', ...]
            
        # Récupération des équipements
        equipements = []
        i = 0
        while f"type_{i}" in request.form:
            type_ = request.form.get(f"type_{i}")
            p_unitaire = float(request.form.get(f"puissance_unitaire_{i}") or 0)
            nombre = float(request.form.get(f"nombre_{i}") or 0)
            p_totale = p_unitaire * nombre
            t_utilisation = float(request.form.get(f"temps_utilisation_{i}") or 0)
            fonctionnement = request.form.get(f"fonctionnement_{i}") or ""
            w_m2 = float(request.form.get(f"w_m2_{i}") or 0)

            equipements.append({
                "type": type_,
                "puissance_unitaire": p_unitaire,
                "nombre": nombre,
                "puissance_totale": p_totale,
                "temps_utilisation": t_utilisation,
                "fonctionnement": fonctionnement,
                "w_m2": w_m2,
            })
            i += 1
            
            
        # Récupération de l'inventaire
        inventaire = []
        index = 0
        while f"marque_inv_{index}" in request.form:
            marque = request.form.get(f"marque_inv_{index}")
            reference = request.form.get(f"reference_inv_{index}")
            puissance = float(request.form.get(f"puissance_inv_{index}") or 0)
            nombre = float(request.form.get(f"nombre_inv_{index}") or 0)
            temps = float(request.form.get(f"temps_utilisation_inv_{index}") or 0)
            fonctionnement = request.form.get(f"fonctionnement_inv_{index}") or ""
            w_m2 = float(request.form.get(f"w_m2_inv_{index}") or 0)
            puissance_totale = puissance * nombre
            inventaire.append({
                "marque": marque,
                "reference": reference,
                "puissance": puissance,
                "nombre": nombre,
                "puissance_totale": puissance_totale,
                "temps_utilisation": temps,
                "fonctionnement": fonctionnement,
                "w_m2": w_m2
            })
            index += 1
            
        
        # Récupération des zones
        zones = []
        i = 0
        while f"nom_zone_{i}" in request.form:
            zone = {
                "nom": request.form.get(f"nom_zone_{i}"),
                "usage": request.form.get(f"usage_zone_{i}", "COMMERCE"),  # ✅ Ajout ici
                "surface": float(request.form.get(f"surface_zone_{i}") or 0),
                "lux_projete": float(request.form.get(f"lux_projete_zone_{i}") or 0),
                "nbr_luminaires": int(request.form.get(f"nbr_luminaires_zone_{i}") or 0),
                "coeff_gradation": float(request.form.get(f"coeff_gradation_zone_{i}") or 0),
                "w_m2_theorique": float(request.form.get(f"w_m2_theorique_zone_{i}") or 0),
                "w_m2_reel": float(request.form.get(f"w_m2_reel_zone_{i}") or 0),
                "w_m2_100lux_reel": float(request.form.get(f"w_m2_100lux_reel_zone_{i}") or 0),
            }
            zones.append(zone)
            i += 1
            
          
        # Calcul du global
        if zones:
            global_zone = {
                "surface": sum(float(z["surface"]) for z in zones),
                #"lux_projete": sum(float(z["lux_projete"]) for z in zones),
                "lux_projete": round(sum(z["lux_projete"] for z in zones) / len(zones), 2),
                "nbr_luminaires": sum(int(z["nbr_luminaires"]) for z in zones),
                #"coeff_gradation": round(sum(float(z["coeff_gradation"]) for z in zones) / len(zones), 2),
                "coeff_gradation": round(sum(z["coeff_gradation"] for z in zones) / len(zones), 2),
                "w_m2_theorique": round(sum(float(z["w_m2_theorique"]) for z in zones) / len(zones), 2),
                "w_m2_reel": round(sum(float(z["w_m2_reel"]) for z in zones) / len(zones), 2),
                "w_m2_100lux_reel": round(sum(float(z["w_m2_100lux_reel"]) for z in zones) / len(zones), 2),
            }
        else:
            global_zone = {
                "surface": 0,
                "lux_projete": 0,
                "nbr_luminaires": 0,
                "coeff_gradation": 0,
                "w_m2_theorique": 0,
                "w_m2_reel": 0,
                "w_m2_100lux_reel": 0,
            }

            
        # Récupération des images
        images = []
        for i in range(1, 6):
            file = request.files.get(f"image_{i}")
            if file and file.filename:
                upload_folder = os.path.join("static", "uploads")
                os.makedirs(upload_folder, exist_ok=True)  # Crée le dossier s'il n'existe pas

                filename = file.filename  # Récupère le nom du fichier
                filepath = os.path.join(upload_folder, filename)  # Construit le chemin complet

                file.save(filepath)  # Sauvegarde le fichier

                images.append(InlineImage(doc, filepath, width=Mm(25)))  # largeur personnalisable
            else:
                images.append("")  # Placeholder si aucune image


        # Totaux
        total_nombre = sum(i["nombre"] for i in inventaire)
        total_puissance_inventaires = sum(i["puissance_totale"] for i in inventaire)
        total_temps = sum(i["temps_utilisation"] for i in inventaire)



        # Totaux
        total_nombre = sum(e["nombre"] for e in equipements)
        total_puissance_totale = sum(e["puissance_totale"] for e in equipements)
        total_temps_utilisation = sum(e["temps_utilisation"] for e in equipements)


        # Totaux
        total_luminaires = sum(l["nombre"] for l in luminaires)
        total_puissance_luminaires = sum(l["puissance_totale"] for l in luminaires)

        # --- Calculs automatiques ---
        conso_initiale = (total_puissance_totale * total_temps_utilisation) / 1000
        conso_projete = (total_puissance_inventaires * total_temps) / 1000
        economie_energie = conso_initiale - conso_projete
        emissions_evitees = economie_energie * 0.08
        puissance_theorique = float(request.form.get("puissance_theorique") or 0)
        puissance_reelle_projete = puissance_theorique * global_zone["coeff_gradation"]



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
            #"puissance_réelle_projetée": request.form.get("puissance_réelle_projetée"),
            #"conso_initiale": request.form.get("conso_initiale"),
            #"conso_projetée": request.form.get("conso_projetée"),
            #"economie_energie": request.form.get("economie_energie"),
            #"emissions": request.form.get("emissions"),
            "conso_initiale": conso_initiale,
            "conso_projetée": conso_projete,
            "economie_energie": economie_energie,
            "emissions": emissions_evitees,
            "puissance_theorique": puissance_theorique,
            "puissance_réelle_projetée": puissance_reelle_projete,
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
            "W_m2": request.form.get("W_m2"),
            "bâtiments": request.form.get("bâtiments"),
            "secteur_etude": request.form.get("secteur_etude"),
            "seuil_reglementaire": request.form.get("seuil_reglementaire"),
            "puissance_installée": request.form.get("puissance_installée"),
            "consommation_energie": request.form.get("consommation_energie"),
            "luminaires": luminaires,
            "total_luminaires": total_luminaires,
            "total_puissance": total_puissance_inventaires,
            'zones': zones,
            'zones_1': zones_1,
            "global": global_zone,
            "equipements": equipements,
            "total_nombre": total_nombre,
            "total_puissance_totale": total_puissance_totale,
            "total_temps_utilisation": total_temps_utilisation,
            "inventaire": inventaire,
            "total_nombre": total_nombre,
            "total_puissance": total_puissance_luminaires,
            "total_temps": total_temps,
            "nom": request.form.get(f"nom_zone_{i}"),
            "usage": request.form.get(f"usage_zone_{i}", "COMMERCE"),
            "secteur_etude": request.form.get("secteur_etude", "Locaux de vente (35.1)"),
            "seuil_reglementaire": float(request.form.get("seuil_reglementaire") or 300),
            "eclairement_moyen": float(request.form.get("eclairement_moyen") or 0),
            "seuil_uniformite": float(request.form.get("seuil_uniformite") or 0.4),
            "uniformite_modelee": float(request.form.get("uniformite_modelee") or 0),
            "puissance_theorique": float(request.form.get("puissance_theorique") or 0),
            "puissance_reelle": float(request.form.get("puissance_reelle") or 0),
            "ratio_lux_max": float(request.form.get("ratio_lux_max") or 1.6),
            "ratio_lux_projete": float(request.form.get("ratio_lux_projete") or 0),
            "conso_projete": float(request.form.get("conso_projete") or 0),
            "economies_energie": float(request.form.get("economies_energie") or 0),
            "emissions_evitees": float(request.form.get("emissions_evitees") or 0),
            "image_signature": image_signature,
            "image_1": images[0],
            "image_2": images[1],
            "image_3": images[2],
            "image_4": images[3],
            "image_5": images[4],
        }

        base_filename = f"rapport_{uuid.uuid4().hex}"
        output_docx = os.path.join(OUTPUT_FOLDER, base_filename + ".docx")
        doc.render(context)
        doc.save(output_docx)


        return f"""
        <!DOCTYPE html>
        <html lang='fr'>
        <head>
            <meta charset='UTF-8'>
            <title>Rapport généré</title>
            <meta http-equiv='refresh' content='10; URL=/download/{os.path.basename(output_docx)}'>
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
                <p><a href='/download/{os.path.basename(output_docx)}'>📄 Télécharger le fichier Word</a></p>
                <p>Vous pouvez ensuite l'ouvrir avec Word et l'exporter en PDF via <strong>Fichier &gt; Exporter en PDF</strong>.</p>
                <p>(Vous allez être redirigé automatiquement dans 10 secondes...)</p>
            </div>
        </body>
        </html>
        """
        


    # première fois → afficher le formulaire vide
    return render_template("form.html")


@app.route("/generate", methods=["POST"])
def generate():
    context = request.form.to_dict()  # on récupère les données validées

    doc = DocxTemplate(MODEL_PATH)
    doc.render(context)

    filename = f"rapport_{uuid.uuid4().hex}.docx"
    output_path = os.path.join(OUTPUT_FOLDER, filename)
    doc.save(output_path)

    return render_template("success.html", filename=filename)
    
    

@app.route("/download/<filename>")
def download(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
    
#if __name__ == "__main__":
 #   app.run(debug=True)
    
    





'''from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
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
        image_path = ""
        if "image_signature" in request.files:
            image_file = request.files["image_signature"]
            if image_file.filename:
                image_name = f"signature_{uuid.uuid4().hex}.png"
                image_path = os.path.join(UPLOAD_FOLDER, image_name)
                image_file.save(image_path)

        doc = DocxTemplate(MODEL_PATH)
        image_signature = InlineImage(doc, image_path, width=Mm(50)) if image_path else ""

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
            "W_m2": request.form.get("W_m2"),
            "bâtiments": request.form.get("bâtiments"),
            "secteur_etude": request.form.get("secteur_etude"),
            "seuil_reglementaire": request.form.get("seuil_reglementaire"),
            "puissance_ins
            tallée": request.form.get("puissance_installée"),
            "consommation_energie": request.form.get("consommation_energie"),
            "image_signature": image_signature
        }

        base_filename = f"rapport_{uuid.uuid4().hex}"
        output_docx = os.path.join(OUTPUT_FOLDER, base_filename + ".docx")
        doc.render(context)
        doc.save(output_docx)

        return f"""
        <!DOCTYPE html>
        <html lang='fr'>
        <head>
            <meta charset='UTF-8'>
            <title>Rapport généré</title>
            <meta http-equiv='refresh' content='10; URL=/download/{os.path.basename(output_docx)}'>
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
                <p><a href='/download/{os.path.basename(output_docx)}'>📄 Télécharger le fichier Word</a></p>
                <p>Vous pouvez ensuite l'ouvrir avec Word et l'exporter en PDF via <strong>Fichier &gt; Exporter en PDF</strong>.</p>
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
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <title>Génération de Rapport</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 0;
      background-color: #f5f7fa;
      color: #333;
    }

    .container {
      max-width: 1000px;
      margin: 40px auto;
      padding: 30px;
      background-color: #fff;
      border-radius: 16px;
      box-shadow: 0 0 20px rgba(0,0,0,0.05);
    }

    h1 {
      text-align: center;
      color: #007bff;
      margin-bottom: 30px;
    }

    h2 {
      border-left: 5px solid #007bff;
      padding-left: 10px;
      color: #007bff;
      margin-top: 40px;
    }

    form {
      display: grid;
      gap: 20px;
    }

    .form-group {
      display: grid;
      gap: 6px;
    }

    label {
      font-weight: 500;
    }

    input, textarea, select {
      padding: 10px;
      border-radius: 8px;
      border: 1px solid #ccc;
      font-size: 1rem;
      transition: border 0.2s ease;
      width: 100%;
    }

    input:focus, textarea:focus, select:focus {
      border-color: #007bff;
      outline: none;
    }

    button {
      background-color: #007bff;
      color: white;
      padding: 15px 25px;
      border: none;
      border-radius: 10px;
      font-size: 1rem;
      cursor: pointer;
      transition: background-color 0.3s ease;
      margin-top: 20px;
    }

    button:hover {
      background-color: #0056b3;
    }

    .grid-2 {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 20px;
    }

    @media (max-width: 768px) {
      .grid-2 {
        grid-template-columns: 1fr;
      }
    }
  </style>
</head>
<body>

  <div class="container">
    <h1>📝 Génération de Rapport</h1>

    <form method="POST" enctype="multipart/form-data">

      <!-- Informations du Client -->
      <h2>Informations du Client</h2>
      <div class="grid-2">
        <div class="form-group">
          <label>Société</label>
          <input type="text" name="client_ste" required>
        </div>
        <div class="form-group">
          <label>Adresse</label>
          <input type="text" name="client_address" required>
        </div>
        <div class="form-group">
          <label>Version</label>
          <input type="text" name="version" required>
        </div>
        <div class="form-group">
          <label>Nom du Contact</label>
          <input type="text" name="client_name" required>
        </div>
        <div class="form-group">
          <label>Email</label>
          <input type="email" name="client_mail" required>
        </div>
        <div class="form-group">
          <label>Téléphone</label>
          <input type="text" name="client_phone" required>
        </div>
      </div>

      <!-- Fiche d'identité du site -->
      <h2>Fiche Identité du Site</h2>
      <div class="grid-2">
        <div class="form-group"><label>Nom du Site</label><input type="text" name="ste" required></div>
        <div class="form-group"><label>Adresse du Chantier</label><input type="text" name="address" required></div>
        <div class="form-group"><label>Surface éclairée (m²)</label><input type="number" name="surface" required></div>
        <div class="form-group"><label>Secteur d'activité</label><input type="text" name="activité" required></div>
        <div class="form-group"><label>Nombre de bâtiments/zones</label><input type="number" name="nbr_batiments" required></div>
        <div class="form-group"><label>Date de visite</label><input type="date" name="date_visite" required></div>
        <div class="form-group"><label>Date de l'étude</label><input type="date" name="date_etude" required></div>
        <div class="form-group"><label>Opqbi be</label><input type="text" name="audit"></div>
        <div class="form-group"><label>Contact</label><input type="text" name="contact"></div>
        <div class="form-group"><label>Station Météo</label><input type="text" name="station_meteo"></div>
        <div class="form-group"><label>Nom du Client</label><input type="text" name="nom_client"></div>
        <div class="form-group"><label>Téléphone du Client</label><input type="text" name="telephone_client"></div>
      </div>

      <!-- Résultats de simulation -->
      <h2>Inventaire et Résultats de Simulation</h2>
      <div class="grid-2">
        <div class="form-group"><label>Puissance unitaire (W)</label><input type="number" name="p_unitaire" step="0.1"></div>
        <div class="form-group"><label>Nombre de luminaires</label><input type="number" name="nombre"></div>
        <div class="form-group"><label>Puissance totale (W)</label><input type="number" name="p_totale" step="0.1"></div>
        <div class="form-group"><label>Temps d'utilisation (h/an)</label><input type="number" name="t_utilisation"></div>
        <div class="form-group"><label>Fonctionnement</label><input type="text" name="fontionnement"></div>
        <div class="form-group"><label>W/m²</label><input type="text" name="W_m2"></div>
        <div class="form-group"><label>Nom des bâtiments</label><input type="text" name="bâtiments"></div>
        <div class="form-group"><label>Secteur étude (DIALUX)</label><input type="text" name="secteur_etude"></div>
        <div class="form-group"><label>Seuil réglementaire (Lux)</label><input type="number" name="seuil_reglementaire"></div>
        <div class="form-group"><label>Puissance installée (W)</label><input type="number" name="puissance_installée"></div>
        <div class="form-group"><label>Consommation d'énergie (kWh/an)</label><input type="number" name="consommation_energie"></div>
        <div class="form-group"><label>Nombre total synthèse</label><input type="number" name="syn_nbr"></div>
        <div class="form-group"><label>Puissance totale synthèse</label><input type="number" name="syn_p_totale"></div>
        <div class="form-group"><label>Niveau d'éclairement moyen (Lux)</label><input type="number" name="niveau_eclairage"></div>
        <div class="form-group"><label>Facteur d'uniformité</label><input type="text" name="facteur_uniformité"></div>
        <div class="form-group"><label>Puissance initiale (W)</label><input type="number" name="puissance_init"></div>
        <div class="form-group"><label>Puissance projetée (W)</label><input type="number" name="puissance_projetée"></div>
        <div class="form-group"><label>Puissance réelle projetée (W)</label><input type="number" name="puissance_réelle_projetée"></div>
        <div class="form-group"><label>Consommation initiale (kWh/an)</label><input type="number" name="conso_initiale"></div>
        <div class="form-group"><label>Consommation projetée (kWh/an)</label><input type="number" name="conso_projetée"></div>
        <div class="form-group"><label>Économies d'énergie (kWh/an)</label><input type="number" name="economie_energie"></div>
        <div class="form-group"><label>Émissions CO₂ évitées (kgeqCO₂)</label><input type="number" name="emissions"></div>
      </div>

      <!-- Upload d'image -->
      <h2>Image de Signature</h2>
      <div class="form-group">
        <label for="image_signature">Télécharger une image (signature, logo...)</label>
        <input type="file" name="image_signature" accept="image/*">
      </div>

      <!-- Bouton de soumission -->
      <button type="submit">🚀 Générer le rapport</button>
    </form>
  </div>

</body>
</html>

'''