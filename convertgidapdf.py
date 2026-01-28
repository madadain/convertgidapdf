
import os
import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog, messagebox, BooleanVar, Toplevel, Checkbutton, Label, Button

# =============================
# Utilitaires
# =============================

def to_decimal_hours(time_str):
    """Convertit 'hh:mm' en heures décimales (float, arrondi à 2 décimales)."""
    if pd.isna(time_str):
        return 0
    s = str(time_str).strip()
    if not s:
        return 0

    # Normalisation de formats fréquents
    s = s.replace("h", ":").replace("H", ":").replace(" ", "")
    if ":" in s:
        try:
            parts = s.split(":")
            if len(parts) >= 2:
                h = int(parts[0]) if parts[0] else 0
                m = int(parts[1]) if parts[1] else 0
                return round(h + m / 60, 2)
        except Exception:
            return 0
    return 0

def get_company_name(pdf_path):
    """Retourne le nom de société = nom du fichier sans extension .pdf."""
    return os.path.splitext(os.path.basename(pdf_path))[0].strip()

def safe_table_to_df(table):
    """
    Convertit une table (liste de lignes) en DataFrame et normalise le nombre de colonnes.
    Retourne un DF avec colonnes Col1..ColN (au moins 6 colonnes pour supporter Col6).
    """
    if not table:
        return None
    df = pd.DataFrame(table)
    df.columns = [f"Col{i+1}" for i in range(len(df.columns))]
    needed = 6 - len(df.columns)
    if needed > 0:
        for _ in range(needed):
            df[f"Col{len(df.columns) + 1}"] = None
    return df

def process_pdf(pdf_path):
    """
    Extrait et transforme les tables d'un PDF:
    - Concatène toutes les tables
    - Renomme colonnes Col1..ColN
    - Filtre lignes contenant 'Total' dans Col1
    - Garde Col1, Col2, Col6 -> Libellé, Valeur, Temps_hhmm
    - Ajoute Heures_decimales
    - Ajoute Societe (nom du fichier sans .pdf) en première colonne
    """
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            try:
                tables = page.extract_tables() or []
            except Exception:
                tables = []
            for table in tables:
                df = safe_table_to_df(table)
                if df is not None and not df.empty:
                    all_tables.append(df)

    if not all_tables:
        return pd.DataFrame(columns=["Societe", "Libellé", "Valeur", "Temps_hhmm", "Heures_decimales"])

    final_df = pd.concat(all_tables, ignore_index=True)
    final_df.columns = [f"Col{i+1}" for i in range(len(final_df.columns))]

    # Filtrer 'Total' dans Col1
    if "Col1" in final_df.columns:
        mask_total = final_df["Col1"].astype(str).str.contains("Total", case=False, na=False)
        final_df = final_df[mask_total]
    else:
        return pd.DataFrame(columns=["Societe", "Libellé", "Valeur", "Temps_hhmm", "Heures_decimales"])

    for col in ["Col2", "Col6"]:
        if col not in final_df.columns:
            final_df[col] = None

    final_df = final_df[["Col1", "Col2", "Col6"]].copy()
    final_df.columns = ["Libellé", "Valeur", "Temps_hhmm"]

    # Nettoyage Temps_hhmm
    final_df["Temps_hhmm"] = final_df["Temps_hhmm"].apply(
        lambda x: str(x).strip().replace("h", ":").replace("H", ":") if pd.notna(x) else None
    )
    # Conversion en heures décimales
    final_df["Heures_decimales"] = final_df["Temps_hhmm"].apply(to_decimal_hours)

    # Première colonne: nom de société (nom du fichier sans .pdf)
    final_df.insert(0, "Societe", get_company_name(pdf_path))

    return final_df

# =============================
# UI & Contrôle principal
# =============================

def main():
    root = Tk()
    root.withdraw()

    # Choix récursif (sous-dossiers)
    recursive_choice = BooleanVar(value=False)

    top = Toplevel()
    top.title("Options d'extraction PDF")
    Label(top, text="Parcourir aussi les sous-dossiers ?").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    chk = Checkbutton(top, text="Oui, inclure les sous-dossiers", variable=recursive_choice)
    chk.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="w")

    def validate_choice():
        top.destroy()

    Button(top, text="Continuer", command=validate_choice).grid(row=2, column=0, padx=10, pady=(0, 10), sticky="e")
    top.lift()
    top.attributes("-topmost", True)
    root.wait_window(top)

    # Sélection du dossier contenant les PDF
    folder = filedialog.askdirectory(title="Choisir le dossier contenant les PDF")
    if not folder:
        messagebox.showerror("Erreur", "Aucun dossier sélectionné.")
        return

    # Sélection du fichier Excel de sortie
    output_file = filedialog.asksaveasfilename(
        title="Enregistrer le fichier Excel",
        defaultextension=".xlsx",
        filetypes=[("Fichiers Excel", "*.xlsx")]
    )
    if not output_file:
        messagebox.showerror("Erreur", "Aucun chemin de sortie sélectionné.")
        return

    # Liste des PDF
    pdf_paths = []
    if recursive_choice.get():
        for root_dir, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".pdf"):
                    pdf_paths.append(os.path.join(root_dir, f))
    else:
        for f in os.listdir(folder):
            if f.lower().endswith(".pdf"):
                pdf_paths.append(os.path.join(folder, f))

    if not pdf_paths:
        messagebox.showwarning("Attention", "Aucun fichier PDF trouvé dans le dossier sélectionné.")
        return

    # Traitement
    all_results = []
    errors = []
    for pdf_path in pdf_paths:
        try:
            df = process_pdf(pdf_path)
            if df is not None and not df.empty:
                all_results.append(df)
        except Exception as e:
            errors.append((pdf_path, str(e)))

    if not all_results:
        msg = "Aucune table 'Total' trouvée dans les PDF."
        if errors:
            msg += f"\n\nErreurs sur {len(errors)} fichier(s):\n" + "\n".join([f"- {os.path.basename(p)} : {err}" for p, err in errors[:10]])
        messagebox.showwarning("Attention", msg)
        return

    final_df = pd.concat(all_results, ignore_index=True)

    try:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Synthèse")
        resume = [
            f"Extraction terminée ✅",
            f"Fichier Excel : {output_file}",
            f"PDF traités : {len(pdf_paths)}",
            f"Lignes extraites : {len(final_df)}",
        ]
        if errors:
            resume.append(f"PDF en erreur : {len(errors)}")
            resume.append("Aperçu erreurs :")
            resume.extend([f"- {os.path.basename(p)} : {err}" for p, err in errors[:5]])
        messagebox.showinfo("Succès", "\n".join(resume))
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible d'écrire le fichier Excel : {e}")


if __name__ == "__main__":
    main()
