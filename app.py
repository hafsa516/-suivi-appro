import streamlit as st
import pandas as pd
import re
import datetime
import io
import os

# ─────────────────────────────────────────────
#  Utilitaires
# ─────────────────────────────────────────────

def trouver_col(df, nom):
    return next((c for c in df.columns if c.strip().upper() == nom.upper()), None)


def week_to_friday(week_str):
    try:
        match = re.search(r'W(\d+)', str(week_str).upper())
        if not match:
            return ''
        week_num = int(match.group(1))
        annee    = datetime.datetime.now().year
        vendredi = datetime.datetime.strptime(
            f'{annee}-W{week_num:02d}-5', '%G-W%V-%u'
        )
        return vendredi.strftime('%d.%m.%Y')
    except Exception:
        return ''


def calculer_confirmation(date_str):
    try:
        if date_str == '':
            return ''
        date_conf   = datetime.datetime.strptime(date_str, '%d.%m.%Y')
        aujourd_hui = datetime.datetime.now()
        diff        = (date_conf - aujourd_hui).days
        if diff < 0:     return ''
        elif diff > 30:  return 'CONFD1'
        elif diff >= 15: return 'CONFD2'
        else:            return 'CONFD3'
    except Exception:
        return ''


def get_friday_of_current_week():
    """Retourne le vendredi de la semaine en cours au format DD.MM.YYYY"""
    aujourd_hui = datetime.datetime.now()
    days_until_friday = (4 - aujourd_hui.weekday()) % 7  # 4 = Friday
    vendredi = aujourd_hui + datetime.timedelta(days=days_until_friday)
    return vendredi.strftime('%d.%m.%Y')


# ─────────────────────────────────────────────
#  Lecture Excel
# ─────────────────────────────────────────────

def lire_excel_rapide(file_obj, feuilles):
    for engine in ('calamine', 'openpyxl'):
        try:
            file_obj.seek(0)
            data = pd.read_excel(
                file_obj,
                sheet_name=feuilles,
                dtype=str,
                engine=engine
            )
            return data, engine
        except Exception:
            continue
    raise RuntimeError("Impossible de lire le fichier Excel.")


# ─────────────────────────────────────────────
#  Écriture vers BytesIO (xlsxwriter)
# ─────────────────────────────────────────────

def ecrire_xlsxwriter(df_final, df_out, col_lt):
    import xlsxwriter

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {
        'in_memory': True,
        'strings_to_numbers': True
    })

    fmt_header = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
    fmt_bleu   = wb.add_format({'bg_color': '#ADD8E6'})
    fmt_rouge  = wb.add_format({'bg_color': '#F29999'})
    fmt_normal = wb.add_format({})

    # ── Feuille Résultat ──────────────────────────────────────
    ws_res   = wb.add_worksheet('Résultat')
    colonnes = df_final.columns.tolist()
    lt_idx   = colonnes.index(col_lt) if col_lt and col_lt in colonnes else -1

    for c_idx, col_name in enumerate(colonnes):
        ws_res.write(0, c_idx, col_name, fmt_header)

    for r_idx, row in enumerate(df_final.itertuples(index=False), start=1):
        row_data = list(row)
        lt_val   = str(row_data[lt_idx]).upper() if lt_idx >= 0 else ''
        if lt_val == 'W/O FRNS':
            fmt = fmt_bleu
        elif re.search(r'W\d+', lt_val) or lt_val == 'W/O RA':
            fmt = fmt_rouge
        else:
            fmt = fmt_normal
        ws_res.set_row(r_idx, None, fmt)
        ws_res.write_row(r_idx, 0, row_data)

    # ── Feuille Tableau Confirmation ──────────────────────────
    ws_conf = wb.add_worksheet('Tableau Confirmation')

    if not df_out.empty:
        for c_idx, col_name in enumerate(df_out.columns.tolist()):
            ws_conf.write(0, c_idx, col_name, fmt_header)
        for r_idx, row in enumerate(df_out.itertuples(index=False), start=1):
            ws_conf.write_row(r_idx, 0, list(row))
    else:
        cols_conf = ['NO commande', 'n° poste', 'Fournisseur', 'référence',
                     'Désignation', 'date de confirmation', 'Qte confirmée',
                     'Référence confirmation', 'Unité', 'CA']
        for c_idx, col_name in enumerate(cols_conf):
            ws_conf.write(0, c_idx, col_name, fmt_header)

    wb.close()
    output.seek(0)
    return output


# ─────────────────────────────────────────────
#  Logique métier principale
# ─────────────────────────────────────────────

def traiter_excel(file_obj, log_fn, progress_fn):
    debut = datetime.datetime.now()

    # Étape 1 : Lecture
    log_fn("📂 Lecture du fichier...", "info")
    progress_fn(5)

    try:
        sheets, engine_used = lire_excel_rapide(file_obj, ['Commandes', 'Suivi Appro'])
        log_fn(f"   Moteur utilisé : {engine_used}", "info")
    except Exception as e:
        log_fn(f"❌ Lecture impossible : {e}", "error")
        return None, None, None

    if 'Commandes' not in sheets or 'Suivi Appro' not in sheets:
        log_fn("❌ Feuille 'Commandes' ou 'Suivi Appro' introuvable.", "error")
        return None, None, None
    progress_fn(20)

    # Étape 2 : Nettoyage et filtrage
    log_fn("📋 Filtrage 'Commandes'...", "info")
    df = sheets['Commandes'].fillna('').astype(str)
    df.columns = df.columns.str.strip()

    result = df[
        (df['Date confirmée'].str.strip() == '') &
        (df['Date_Reception'].str.strip() == '') &
        (df['Infos-ach'].str.strip().str.contains(
            r'(?i)cr[eé][eé]\s*par\s*:\s*F[A-Za-z0-9]+',
            regex=True, na=False
        ))
    ].copy()
    log_fn(f"   {len(result)} lignes retenues", "info")
    progress_fn(35)

    df_suivi = sheets['Suivi Appro'].fillna('').astype(str)
    df_suivi.columns = df_suivi.columns.str.strip()

    # Étape 3 : Jointure
    log_fn("🔗 Fusion des feuilles...", "info")
    cols_result = result.columns.tolist()
    cols_suivi  = df_suivi.columns.tolist()

    result['Liste Cdes']       = result['Liste Cdes'].str.strip().str.lower()
    result['Liste Poste Cdes'] = result['Liste Poste Cdes'].str.strip().str.lower()
    df_suivi['NO DE COMMANDE'] = df_suivi['NO DE COMMANDE'].str.strip().str.lower()
    df_suivi['POSTE CDE']      = df_suivi['POSTE CDE'].str.strip().str.lower()

    df_final = result.merge(
        df_suivi,
        left_on=['Liste Cdes', 'Liste Poste Cdes'],
        right_on=['NO DE COMMANDE', 'POSTE CDE'],
        how='left',
        suffixes=('_cmd', '_suivi')
    )
    df_final = df_final.drop(columns=['NO DE COMMANDE', 'POSTE CDE'], errors='ignore')

    cols_a_supprimer = []
    for col_cmd in cols_result:
        for col_s in cols_suivi:
            if col_s in cols_a_supprimer: continue
            if col_cmd.strip().lower() != col_s.strip().lower(): continue
            if col_s not in df_final.columns or col_cmd not in df_final.columns: continue
            cols_a_supprimer.append(col_s)

    df_final = df_final.drop(columns=cols_a_supprimer, errors='ignore').fillna('')
    log_fn(f"   {len(df_final)} lignes, {len(df_final.columns)} colonnes", "info")
    progress_fn(55)

    # Étape 4 : Tableau Confirmation
    log_fn("📊 Construction tableau confirmation...", "info")
    col_leadtime = trouver_col(df_final, 'LEADTIME')
    if col_leadtime:
        # Filtrer les lignes avec W\d+ ou W/O RA
        df_conf = df_final[
            df_final[col_leadtime].str.upper().str.contains(r'W\d+', regex=True, na=False) |
            (df_final[col_leadtime].str.upper() == 'W/O RA')
        ].copy()

        col_cde_cmd   = trouver_col(df_conf, 'Doc_achat')
        col_poste_cmd = trouver_col(df_conf, 'Poste')
        col_fourn     = trouver_col(df_conf, 'Fourn/Div_fourn')
        col_art       = trouver_col(df_conf, 'Article')
        col_desig     = trouver_col(df_conf, 'Designation')
        col_uac       = trouver_col(df_conf, 'UAc')
        col_qte       = trouver_col(df_conf, 'A_livrer')

        noms_charge = ['Chargé appro', 'Charge appro', "Chargé d'appro",
                       'Chargé Appro', 'CHARGE APPRO', 'CA', 'CHARGE APPRO_cmd']
        col_charge = next(
            (trouver_col(df_conf, n) for n in noms_charge
             if trouver_col(df_conf, n) is not None), None)
        if col_charge is None:
            for n in noms_charge:
                if (n + '_cmd') in df_conf.columns:
                    col_charge = n + '_cmd'
                    break

        # Fonction pour traiter chaque ligne (date et référence)
        def process_row(row):
            lt_val = str(row[col_leadtime]).upper()
            
            if lt_val == 'W/O RA':
                # Pour W/O RA : vendredi de la semaine en cours + CONFD3
                date_conf = get_friday_of_current_week()
                ref_conf = 'CONFD3'
            else:
                # Pour W\d+ : calcul normal
                date_conf = week_to_friday(lt_val)
                ref_conf = calculer_confirmation(date_conf)
            
            return pd.Series({
                'date_conf': date_conf,
                'ref_conf': ref_conf
            })

        # Appliquer le traitement à chaque ligne
        df_conf[['date_de_confirmation', 'reference_confirmation']] = df_conf.apply(process_row, axis=1)

        # Construction du DataFrame final de confirmation
        df_out = pd.DataFrame()
        df_out['NO commande']            = df_conf[col_cde_cmd].str.strip().str.upper() + ',' if col_cde_cmd else ''
        df_out['n° poste']               = df_conf[col_poste_cmd].str.strip() + ',' if col_poste_cmd else ''
        df_out['Fournisseur']            = df_conf[col_fourn].str.strip() if col_fourn else ''
        df_out['référence']              = df_conf[col_art].str.strip() if col_art else ''
        df_out['Désignation']            = df_conf[col_desig].str.strip() if col_desig else ''
        df_out['date de confirmation']   = df_conf['date_de_confirmation']
        df_out['Qte confirmée']          = df_conf[col_qte].str.strip() if col_qte else ''
        df_out['Référence confirmation'] = df_conf['reference_confirmation']
        df_out['Unité']                  = df_conf[col_uac].str.strip() if col_uac else ''
        df_out['CA']                     = df_conf[col_charge].str.strip() if col_charge else ''

        # Compter avant filtrage
        nb_avant = len(df_out)
        
        # Filtrer les lignes avec référence confirmation vide
        df_out = df_out[df_out['Référence confirmation'] != ''].copy()
        
        # Remplir les valeurs vides et réinitialiser l'index
        df_out = df_out.fillna('').reset_index(drop=True)
        
        # Compter les lignes par type après filtrage
        if not df_out.empty and col_leadtime in df_conf.columns:
            # Récupérer les indices des lignes conservées
            indices_conserves = df_out.index.tolist()
            # Compter les W/O RA dans les lignes conservées
            nb_w_ra = len(df_conf.loc[df_conf.index.isin(indices_conserves) & 
                                      (df_conf[col_leadtime].str.upper() == 'W/O RA')])
            nb_w_digit = len(df_out) - nb_w_ra
            log_fn(f"   {len(df_out)} lignes confirmation ({nb_avant - len(df_out)} ignorées) - W\\d+: {nb_w_digit}, W/O RA: {nb_w_ra}", "info")
        else:
            log_fn(f"   {len(df_out)} lignes confirmation ({nb_avant - len(df_out)} ignorées)", "info")
        
    else:
        df_out = pd.DataFrame()
        log_fn("⚠️  Colonne LEADTIME introuvable", "warning")

    progress_fn(70)

    # Étape 5 : Écriture
    log_fn("💾 Génération du fichier Excel...", "info")
    progress_fn(80)
    try:
        output_buffer = ecrire_xlsxwriter(df_final, df_out, col_leadtime)
    except Exception as e:
        log_fn(f"❌ Erreur lors de l'écriture : {e}", "error")
        return None, None, None

    progress_fn(100)
    duree = (datetime.datetime.now() - debut).seconds
    log_fn(f"🎉 Terminé en {duree}s — fichier prêt au téléchargement !", "success")

    return output_buffer, df_final, df_out


# ─────────────────────────────────────────────
#  Interface Streamlit
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Suivi Appro",
    page_icon="📊",
    layout="wide"
)

# CSS personnalisé
st.markdown("""
<style>
    .stApp { background-color: #f0f2f5; }
    .main-header {
        background: linear-gradient(135deg, #1a1a2e, #16213e);
        color: white;
        padding: 2rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        text-align: center;
    }
    .main-header h1 { margin: 0; font-size: 2rem; }
    .main-header p  { margin: 0.5rem 0 0; color: #aab4be; font-size: 0.9rem; }
    .log-box {
        background: #1e2433;
        color: #a8d8a8;
        font-family: Consolas, monospace;
        font-size: 0.85rem;
        padding: 1rem;
        border-radius: 8px;
        min-height: 200px;
        max-height: 350px;
        overflow-y: auto;
        white-space: pre-wrap;
    }
    .stDownloadButton > button {
        background-color: #2ecc71 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        padding: 0.6rem 2rem !important;
        font-size: 1rem !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📊 Suivi Appro</h1>
    <p>Traitement automatique des commandes — Fusion & Tableau Confirmation</p>
</div>
""", unsafe_allow_html=True)

# Upload
uploaded_file = st.file_uploader(
    "Sélectionnez votre fichier Excel (.xlsx)",
    type=["xlsx", "xlsm"],
    help="Le fichier doit contenir les feuilles 'Commandes' et 'Suivi Appro'"
)

if uploaded_file is not None:
    st.info(f"📄 Fichier chargé : **{uploaded_file.name}** ({uploaded_file.size // 1024} Ko)")

    if st.button("▶ Exécuter le traitement", type="primary", use_container_width=True):

        log_messages = []
        log_placeholder   = st.empty()
        progress_bar      = st.progress(0, text="Initialisation...")

        def log_fn(msg, level="info"):
            log_messages.append(msg)
            log_placeholder.markdown(
                '<div class="log-box">' +
                '\n'.join(log_messages) +
                '</div>',
                unsafe_allow_html=True
            )

        def progress_fn(value):
            labels = {
                5: "Lecture du fichier...",
                20: "Feuilles chargées...",
                35: "Filtrage...",
                55: "Fusion...",
                70: "Tableau confirmation...",
                80: "Génération Excel...",
                100: "Terminé !"
            }
            progress_bar.progress(value / 100, text=labels.get(value, "Traitement..."))

        output_buffer, df_final, df_out = traiter_excel(
            uploaded_file, log_fn, progress_fn
        )

        if output_buffer:
            st.success("✅ Traitement terminé avec succès !")

            # Aperçu
            with st.expander("🔍 Aperçu — Feuille Résultat (20 premières lignes)"):
                st.dataframe(df_final.head(20), use_container_width=True)

            if df_out is not None and not df_out.empty:
                with st.expander(f"📋 Aperçu — Tableau Confirmation ({len(df_out)} lignes)"):
                    st.dataframe(df_out, use_container_width=True)

            # Téléchargement
            base_name = os.path.splitext(uploaded_file.name)[0]
            st.download_button(
                label="⬇️ Télécharger le fichier Excel résultat",
                data=output_buffer,
                file_name=f"{base_name}_RÉSULTAT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("❌ Le traitement a échoué. Consultez le journal ci-dessus.")

else:
    st.markdown("""
    <div style="text-align:center; color:#888; padding: 3rem;">
        <p style="font-size:3rem;">📁</p>
        <p>Importez votre fichier Excel pour commencer</p>
        <p style="font-size:0.85rem;">Format attendu : <strong>.xlsx</strong> avec les feuilles <em>Commandes</em> et <em>Suivi Appro</em></p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.markdown(
    "<p style='text-align:center; color:#aaa; font-size:0.8rem;'>Suivi Appro • Propulsé par Streamlit</p>",
    unsafe_allow_html=True
)
