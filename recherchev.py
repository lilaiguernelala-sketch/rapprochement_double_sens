import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Comparateur CEGID vs PEGASE", layout="wide")
st.title("📊 Comparateur CEGID vs PEGASE")

st.write("""
Chargez vos fichiers Excel pour comparer les données dans la colonne `Numero`.
L'application vous donnera un résumé et un fichier Excel avec les résultats.
""")

# Upload des fichiers
fichier_cegid = st.file_uploader("Upload fichier CEGID", type=["xlsx"])
fichier_pegase = st.file_uploader("Upload fichier PEGASE", type=["xlsx"])

colonne_cle = "Numero"

if fichier_cegid and fichier_pegase:

    try:
        cegid = pd.read_excel(fichier_cegid)
        pegase = pd.read_excel(fichier_pegase)
    except Exception as e:
        st.error(f"Erreur lors de la lecture des fichiers Excel : {e}")
        st.stop()

    # Vérifier que la colonne clé existe
    if colonne_cle not in cegid.columns:
        st.error(f"Colonne '{colonne_cle}' absente dans CEGID")
        st.stop()
    if colonne_cle not in pegase.columns:
        st.error(f"Colonne '{colonne_cle}' absente dans PEGASE")
        st.stop()

    # Nettoyage
    cegid[colonne_cle] = cegid[colonne_cle].astype(str).str.strip()
    pegase[colonne_cle] = pegase[colonne_cle].astype(str).str.strip()

    set_cegid = set(cegid[colonne_cle])
    set_pegase = set(pegase[colonne_cle])

    # Comparaison
    cegid["Existe_dans_PEGASE"] = cegid[colonne_cle].apply(
        lambda x: "trouvé" if x in set_pegase else "non trouvé"
    )
    pegase["Existe_dans_CEGID"] = pegase[colonne_cle].apply(
        lambda x: "trouvé" if x in set_cegid else "non trouvé"
    )

    # Résumé
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

    st.success("✅ Comparaison terminée")
    st.subheader("Résumé")
    st.dataframe(resume)

    # Création du fichier Excel en mémoire
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cegid.to_excel(writer, sheet_name="CEGID_vs_PEGASE", index=False)
        pegase.to_excel(writer, sheet_name="PEGASE_vs_CEGID", index=False)
        resume.to_excel(writer, sheet_name="RESUME", index=False)
    output.seek(0)

    # Bouton de téléchargement
    st.download_button(
        label="📥 Télécharger le fichier Excel",
        data=output,
        file_name="COMPARAISON_CEGID_PEGASE.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
