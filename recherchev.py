import pandas as pd

# =========================
# CONFIGURATION
# =========================

fichier_cegid = "CEGID.xlsx"
fichier_pegase = "PEGASE.xlsx"
fichier_sortie = "COMPARAISON_CEGID_PEGASE.xlsx"

colonne_cle = "Numero"


print("Chargement des fichiers en cours...")

cegid = pd.read_excel(fichier_cegid)
pegase = pd.read_excel(fichier_pegase)

# Vérifier colonne
if colonne_cle not in cegid.columns:
    raise Exception(f"Colonne '{colonne_cle}' absente dans CEGID")

if colonne_cle not in pegase.columns:
    raise Exception(f"Colonne '{colonne_cle}' absente dans PEGASE")

# Nettoyage
cegid[colonne_cle] = cegid[colonne_cle].astype(str).str.strip()
pegase[colonne_cle] = pegase[colonne_cle].astype(str).str.strip()


set_cegid = set(cegid[colonne_cle])
set_pegase = set(pegase[colonne_cle])

# COMPARAISONS

print("Comparaison CEGID → PEGASE...")
cegid["Existe_dans_PEGASE"] = cegid[colonne_cle].apply(
    lambda x: "trouvé" if x in set_pegase else "non trouvé"
)

print("Comparaison PEGASE → CEGID...")
pegase["Existe_dans_CEGID"] = pegase[colonne_cle].apply(
    lambda x: "trouvé" if x in set_cegid else "non trouvé"
)


resume = pd.DataFrame({
    "Description": [
        "Total CEGID",
        "Total PEGASE",
        "CEGID trouvés dans PEGASE",
        "CEGID non trouvés dans PEGASE",
        "PEGASE trouvés dans CEGID",
        "PEGASE non trouvés dans CEGID"
    ],
    "Valeur": [
        len(cegid),
        len(pegase),
        (cegid["Existe_dans_PEGASE"] == "trouvé").sum(),
        (cegid["Existe_dans_PEGASE"] == "non trouvé").sum(),
        (pegase["Existe_dans_CEGID"] == "trouvé").sum(),
        (pegase["Existe_dans_CEGID"] == "non trouvé").sum()
    ]
})


# SAUVEGARDE DANS UN SEUL FICHIER EXCEL AVEC 3 FEUILLES

print("Création du fichier Excel avec 3 feuilles...")

with pd.ExcelWriter(fichier_sortie, engine="openpyxl") as writer:
    
    cegid.to_excel(
        writer,
        sheet_name="CEGID_vs_PEGASE",
        index=False
    )
    
    pegase.to_excel(
        writer,
        sheet_name="PEGASE_vs_CEGID",
        index=False
    )
    
    resume.to_excel(
        writer,
        sheet_name="RESUME",
        index=False
    )

print("================== <3 ==================")
print("Fichier créé avec succès ou pas :", fichier_sortie)
print("================== <3 ================")

