import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from io import BytesIO

parameters = { "background_color": '#f0f2f6',
               "text_color": '#333333',
               "header_color": "#4C7AAF", 
               "font_size": '16px' }

st.set_page_config(page_title="PODP RECURRENCE PRODUIT", layout="wide")

st.title("PODP RECURRENCE PRODUIT")
st.header("Analyse de la récurrence des produits PODP.")

# Upload du fichier Excel
st.subheader(" Charger un fichier Excel")
uploaded_file = st.file_uploader(
    "Choisissez un fichier Excel (.xlsx ou .xls)",
    type=['xlsx', 'xls'],
    help="Téléchargez un fichier Excel contenant les données de formation"
)

# Fonction pour charger et traiter les données
@st.cache_data
def load_data(file):
    df = pd.read_excel(file)
    
    # Nettoyer les noms de colonnes (supprimer les retours à la ligne et espaces multiples)
    df.columns = df.columns.str.replace('\n', ' ').str.replace('\r', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()
    
    # Convertir les colonnes de date en datetime
    date_columns = [col for col in df.columns if 'date' in col.lower() or 'début' in col.lower() or 'debut' in col.lower() or 'fin' in col.lower()]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

# Fonction utilitaire pour rechercher une colonne
def find_column(df, keywords_list):
    """Recherche une colonne dans le dataframe en fonction de mots-clés"""
    for keywords in keywords_list:
        for col in df.columns:
            col_normalized = col.upper().replace(' ', '').replace('\n', '').replace('\r', '')
            if all(kw.upper() in col_normalized for kw in keywords):
                return col
    return None

# Vérifier si un fichier a été uploadé
if uploaded_file is not None:
    # Charger les données
    df = load_data(uploaded_file)
    
    # Afficher les informations sur le dataframe
    st.subheader(" Données chargées")
    st.write(f"Nombre total de lignes : {len(df)}")
    st.write(f"Nombre de colonnes : {len(df.columns)}")
    
    # Afficher les colonnes disponibles
    with st.expander("Voir les colonnes disponibles"):
        st.write(df.columns.tolist())
    
    # Section de filtrage
    st.subheader(" Filtres sur les dates de formation")
    
    # Identifier les colonnes de dates de début et fin
    date_debut_col = find_column(df, [['date', 'debut', 'formation'], ['date', 'début', 'formation']])
    date_fin_col = find_column(df, [['date', 'fin', 'formation']])
    
    # Interface simplifiée - Un seul choix de date avec 18 mois glissants automatique
    st.info("📅 Mode 18 mois glissants : sélectionnez une date de début, la période de 18 mois sera calculée automatiquement.")
    
    if date_debut_col and df[date_debut_col].notna().any():
        min_date_debut = df[date_debut_col].min()
        max_date_debut = df[date_debut_col].max()
        
        date_debut_selected = st.date_input(
            "Date de début de formation",
            value=min_date_debut,
            min_value=min_date_debut,
            max_value=max_date_debut,
            help="Sélectionnez la date de début. La période de 18 mois sera automatiquement calculée."
        )
        
        # Calculer automatiquement la date de fin à 18 mois
        date_fin_calculated = date_debut_selected + relativedelta(months=18)
        
        st.success(f" Période analysée : du **{date_debut_selected.strftime('%d/%m/%Y')}** au **{date_fin_calculated.strftime('%d/%m/%Y')}** (18 mois)")
        
    else:
        st.warning("Colonne de date de début non trouvée")
        date_debut_selected = None
        date_fin_calculated = None
    
    # Appliquer les filtres sur les dates de début
    df_filtered = df.copy()
    
    if date_debut_col and date_debut_selected and date_fin_calculated:
        # Filtrer les formations dont la date de début est dans la période de 18 mois
        mask_debut = (
            (df_filtered[date_debut_col] >= pd.Timestamp(date_debut_selected)) &
            (df_filtered[date_debut_col] <= pd.Timestamp(date_fin_calculated))
        )
        df_filtered = df_filtered[mask_debut]
    
    # Afficher les résultats filtrés
    st.subheader(" Données filtrées")
    st.write(f"Nombre de lignes après filtrage : {len(df_filtered)}")
    
   
    st.subheader(" Analyse de récurrence ")
    
    # Identifier les colonnes nécessaires
    centre_col = find_column(df_filtered, [['centre']])
    produit_col = find_column(df_filtered, [['numero', 'produit']])
    etablissement_col = find_column(df_filtered, [['libelle', 'etablissement'], ['libellé', 'etablissement']])
    intitule_offre_col = find_column(df_filtered, [['intitule', 'offre'], ['intitulé', 'offre']])
    region_col = find_column(df_filtered, [['libelle', 'région'], ['libellé', 'région']])
    
    # Identifier toutes les colonnes CAP (CAP1, CAP2, CAP3, etc.)
    # Exclure les colonnes qui ne sont pas CAP1/2/3/4/5 (comme CAPACITE, etc.)
    cap_cols = [col for col in df_filtered.columns if col.upper().startswith('CAP') and col.upper() in ['CAP1', 'CAP2', 'CAP3', 'CAP4', 'CAP5']]
    
    # Créer une colonne combinée avec toutes les valeurs CAP uniques pour chaque ligne
    if cap_cols:
        df_filtered['CAP_COMBINED'] = df_filtered[cap_cols].apply(
            lambda row: '; '.join(sorted(set(str(v) for v in row.dropna().values if v != ''))), 
            axis=1
        )
        cap1_col = 'CAP1' if 'CAP1' in df_filtered.columns else None
    else:
        cap1_col = None
    
    if centre_col and produit_col:
        # Grouper par produit, centre et CAP_COMBINED (pour distinguer les produits avec des CAP différents)
        group_cols = [produit_col, centre_col]
        if cap_cols and 'CAP_COMBINED' in df_filtered.columns:
            group_cols.append('CAP_COMBINED')
        
        recurrence = df_filtered.groupby(group_cols, dropna=False).size().reset_index(name='nb_programmations')
        
        # Ajouter les colonnes supplémentaires via des agrégations (beaucoup plus efficace que iterrows)
        agg_dict = {}
        
        if etablissement_col:
            agg_dict[etablissement_col] = 'first'
        if intitule_offre_col:
            agg_dict[intitule_offre_col] = 'first'
        if region_col:
            agg_dict[region_col] = 'first'
        # Ajouter toutes les colonnes CAP
        for cap_col in cap_cols:
            agg_dict[cap_col] = 'first'
        # Ne pas ajouter CAP_COMBINED dans agg_dict si elle est déjà dans group_cols
        if cap_cols and 'CAP_COMBINED' in df_filtered.columns and 'CAP_COMBINED' not in group_cols:
            agg_dict['CAP_COMBINED'] = 'first'
        if date_debut_col:
            agg_dict[date_debut_col] = 'min'  # Date de début la plus ancienne
        if date_fin_col:
            # Utiliser la date de début la plus récente comme date de fin
            agg_dict[date_fin_col] = lambda x: df_filtered.loc[x.index, date_debut_col].max()
        
        if agg_dict:
            # Utiliser groupby avec agg pour récupérer les informations supplémentaires
            info_supplementaire = df_filtered.groupby(group_cols, dropna=False).agg(agg_dict).reset_index()
            # Merger avec le dataframe de récurrence
            recurrence = recurrence.merge(info_supplementaire, on=group_cols, how='left')
        
        # Réorganiser les colonnes pour un meilleur affichage
        colonnes_ordre = [produit_col]
        if intitule_offre_col and intitule_offre_col in recurrence.columns:
            colonnes_ordre.append(intitule_offre_col)
        if etablissement_col and etablissement_col in recurrence.columns:
            colonnes_ordre.append(etablissement_col)
        if region_col and region_col in recurrence.columns:
            colonnes_ordre.append(region_col)
        # Ajouter toutes les colonnes CAP dans l'ordre
        for cap_col in cap_cols:
            if cap_col in recurrence.columns:
                colonnes_ordre.append(cap_col)
        colonnes_ordre.append(centre_col)
        colonnes_ordre.append('nb_programmations')
        if date_debut_col and date_debut_col in recurrence.columns:
            colonnes_ordre.append(date_debut_col)
        if date_fin_col and date_fin_col in recurrence.columns:
            colonnes_ordre.append(date_fin_col)
        
        # Garder seulement les colonnes qui existent
        colonnes_ordre = [col for col in colonnes_ordre if col in recurrence.columns]
        
        # Identifier la colonne capacité pour l'exclure
        capacite_col = find_column(recurrence, [['capacite', 'offre'], ['capacité', 'offre']])
        if capacite_col and capacite_col in colonnes_ordre:
            colonnes_ordre.remove(capacite_col)
        
        recurrence = recurrence[colonnes_ordre]
        
        # Créer une liste de colonnes à afficher (sans CENTRE)
        colonnes_affichage = [col for col in colonnes_ordre if col != centre_col]
        
        # Section de filtrage interactif
        st.subheader("🔍 Filtres d'affichage")
        
        col_filter1, col_filter2, col_filter3 = st.columns(3)
        
        with col_filter1:
            # Filtre par région
            if region_col and region_col in recurrence.columns:
                regions_disponibles = sorted([r for r in recurrence[region_col].unique() if pd.notna(r)])
                regions_selectionnees = st.multiselect(
                    "Filtrer par région",
                    options=regions_disponibles,
                    default=regions_disponibles,
                    help="Sélectionnez une ou plusieurs régions"
                )
            else:
                regions_selectionnees = None
        
        with col_filter2:
            # Filtre par CAP - recherche dans toutes les colonnes CAP
            cap_columns = [col for col in df_filtered.columns if col.upper().startswith('CAP') and col != 'CAP_COMBINED']
            
            if cap_columns:
                # Extraire toutes les valeurs CAP uniques de toutes les colonnes CAP
                all_cap_values = set()
                for cap_col in cap_columns:
                    values = df_filtered[cap_col].dropna().astype(str).unique()
                    # Garder uniquement les valeurs qui contiennent au moins une lettre
                    all_cap_values.update([v for v in values if v and v != '' and v != 'nan' and any(c.isalpha() for c in v)])
                
                if all_cap_values:
                    cap_disponibles = sorted(list(all_cap_values))
                    cap_selectionnes = st.multiselect(
                        "Filtrer par CAP",
                        options=cap_disponibles,
                        default=cap_disponibles,
                        help="Sélectionnez un ou plusieurs types CAP (recherche dans toutes les colonnes CAP)"
                    )
                else:
                    cap_selectionnes = None
            else:
                cap_selectionnes = None
        
        with col_filter3:
            # Filtre par établissement - adapté aux régions sélectionnées
            if etablissement_col and etablissement_col in recurrence.columns:
                if regions_selectionnees and region_col and region_col in recurrence.columns:
                    # Filtrer les établissements en fonction des régions sélectionnées
                    etablissements_disponibles = sorted(recurrence[recurrence[region_col].isin(regions_selectionnees)][etablissement_col].unique())
                else:
                    # Tous les établissements si pas de région sélectionnée ou pas de colonne région
                    etablissements_disponibles = sorted([e for e in recurrence[etablissement_col].unique() if pd.notna(e)])
                
                etablissements_selectionnes = st.multiselect(
                    "Filtrer par établissement",
                    options=etablissements_disponibles,
                    default=etablissements_disponibles,
                    help="Sélectionnez un ou plusieurs établissements"
                )
            else:
                etablissements_selectionnes = None
        
        # Appliquer les filtres
        recurrence_filtered = recurrence.copy()
        
        if regions_selectionnees and region_col and region_col in recurrence_filtered.columns:
            recurrence_filtered = recurrence_filtered[recurrence_filtered[region_col].isin(regions_selectionnees)]
        
        if cap_selectionnes and len(cap_selectionnes) > 0:
            # Filtrer si au moins une des valeurs CAP sélectionnées est présente dans CAP1, CAP2, CAP3, CAP4 ou CAP5
            cap_filter_cols = [col for col in cap_cols if col in recurrence_filtered.columns]
            
            if cap_filter_cols:
                def row_contains_selected_cap(row):
                    # Vérifier dans chaque colonne CAP
                    for val in row:
                        if pd.notna(val) and cap_selectionnes and str(val) in cap_selectionnes:
                            return True
                    return False
                
                mask = recurrence_filtered[cap_filter_cols].apply(row_contains_selected_cap, axis=1)
                recurrence_filtered = recurrence_filtered[mask]
        
        if etablissements_selectionnes and etablissement_col and etablissement_col in recurrence_filtered.columns:
            recurrence_filtered = recurrence_filtered[recurrence_filtered[etablissement_col].isin(etablissements_selectionnes)]
        
        # Afficher les métriques filtrées
        st.subheader("📊 Vue d'ensemble")
        m1, m2, m3, m4 = st.columns(4)
        
        # Compter le nombre de produits uniques
        total_combos = recurrence_filtered[produit_col].nunique()
        
        # Calculer la récurrence par PRODUIT (somme de toutes les programmations par produit)
        recurrence_par_produit = recurrence_filtered.groupby(produit_col)['nb_programmations'].sum().reset_index()
        combos_faible = len(recurrence_par_produit[recurrence_par_produit['nb_programmations'] < 3])
        combos_bonne = len(recurrence_par_produit[recurrence_par_produit['nb_programmations'] >= 3])
        taux_faible = (combos_faible / len(recurrence_par_produit) * 100) if len(recurrence_par_produit) > 0 else 0
        
        with m1:
            st.metric("Produits uniques", total_combos)
        with m2:
            st.metric("Récurrence < 3", combos_faible, delta=f"{taux_faible:.1f}%", delta_color="inverse")
        with m3:
            st.metric("Récurrence ≥ 3", combos_bonne)
        with m4:
            avg_recurrence = recurrence_filtered['nb_programmations'].mean() if len(recurrence_filtered) > 0 else 0
            st.metric("Moyenne programmations", f"{avg_recurrence:.1f}")
        
        # Répartition temporelle des programmations (0-6, 6-12, 12-18 mois)
        if date_debut_col and date_debut_selected and date_fin_calculated and date_debut_col in df_filtered.columns:
            st.write("**📅 Répartition temporelle des programmations :**")
            
            # Appliquer les mêmes filtres que recurrence_filtered sur df_filtered
            df_temp = df_filtered.copy()
            
            if regions_selectionnees and region_col and region_col in df_temp.columns:
                df_temp = df_temp[df_temp[region_col].isin(regions_selectionnees)]
            
            if cap_selectionnes and len(cap_selectionnes) > 0:
                cap_filter_cols_temp = [col for col in cap_cols if col in df_temp.columns]
                if cap_filter_cols_temp:
                    def row_contains_cap_temp(row):
                        for val in row:
                            if pd.notna(val) and cap_selectionnes and str(val) in cap_selectionnes:
                                return True
                        return False
                    mask_temp = df_temp[cap_filter_cols_temp].apply(row_contains_cap_temp, axis=1)
                    df_temp = df_temp[mask_temp]
            
            if etablissements_selectionnes and etablissement_col and etablissement_col in df_temp.columns:
                df_temp = df_temp[df_temp[etablissement_col].isin(etablissements_selectionnes)]
            
            # Calculer les périodes
            date_debut_ts = pd.Timestamp(date_debut_selected)
            date_6mois = date_debut_ts + relativedelta(months=6)
            date_12mois = date_debut_ts + relativedelta(months=12)
            date_18mois = pd.Timestamp(date_fin_calculated)
            
            # Compter les produits uniques dans chaque période
            total_prog = df_temp[produit_col].nunique()
            prog_0_6 = df_temp[(df_temp[date_debut_col] >= date_debut_ts) & (df_temp[date_debut_col] < date_6mois)][produit_col].nunique()
            prog_6_12 = df_temp[(df_temp[date_debut_col] >= date_6mois) & (df_temp[date_debut_col] < date_12mois)][produit_col].nunique()
            prog_12_18 = df_temp[(df_temp[date_debut_col] >= date_12mois) & (df_temp[date_debut_col] <= date_18mois)][produit_col].nunique()
            
            # Calculer les pourcentages
            pct_0_6 = (prog_0_6 / total_prog * 100) if total_prog > 0 else 0
            pct_6_12 = (prog_6_12 / total_prog * 100) if total_prog > 0 else 0
            pct_12_18 = (prog_12_18 / total_prog * 100) if total_prog > 0 else 0
            
            # Afficher les métriques
            col_t1, col_t2, col_t3 = st.columns(3)
            with col_t1:
                st.metric("0-6 mois", f"{pct_0_6:.1f}%", f"{prog_0_6} produits")
            with col_t2:
                st.metric("6-12 mois", f"{pct_6_12:.1f}%", f"{prog_6_12} produits")
            with col_t3:
                st.metric("12-18 mois", f"{pct_12_18:.1f}%", f"{prog_12_18} produits")
        
        # Afficher le tableau principal (sans la colonne CENTRE)
        st.subheader("📋 Récurrence par produit")
        recurrence_sorted = recurrence_filtered.sort_values('nb_programmations', ascending=False)
        st.dataframe(recurrence_sorted[colonnes_affichage], use_container_width=True, height=400)
        
        # Statistiques descriptives
        st.write("**Statistiques sur les données filtrées :**")
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        with col_s1:
            st.write(f"Moyenne : {recurrence_filtered['nb_programmations'].mean():.2f}")
        with col_s2:
            st.write(f"Médiane : {recurrence_filtered['nb_programmations'].median():.0f}")
        with col_s3:
            st.write(f"Maximum : {recurrence_filtered['nb_programmations'].max()}")
        with col_s4:
            st.write(f"Minimum : {recurrence_filtered['nb_programmations'].min()}")
        
        # Options de téléchargement (sans la colonne CENTRE)
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            csv_tous = recurrence_sorted[colonnes_affichage].to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Télécharger en CSV",
                data=csv_tous,
                file_name="recurrence_produits.csv",
                mime="text/csv"
            )
        
        with col_dl2:
            # Créer un fichier Excel en mémoire (sans la colonne CENTRE)
            buffer_tous = BytesIO()
            with pd.ExcelWriter(buffer_tous, engine='openpyxl') as writer:
                recurrence_sorted[colonnes_affichage].to_excel(writer, index=False, sheet_name='Récurrence produits')
            buffer_tous.seek(0)
            
            st.download_button(
                label="📥 Télécharger en XLSX",
                data=buffer_tous,
                file_name="recurrence_produits.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Statistiques globales (non filtrées)
        st.subheader("📈 KPIs globaux (toutes données)")
        k1, k2, k3, k4 = st.columns(4)

        total_formations = len(df_filtered)
        total_produits = df_filtered[produit_col].nunique() if produit_col and produit_col in df_filtered.columns else 0
        total_centres = df_filtered[centre_col].nunique() if centre_col and centre_col in df_filtered.columns else 0
        avg_prog_global = recurrence['nb_programmations'].mean() if len(recurrence) > 0 else 0

        with k1:
            st.metric("Total formations", f"{total_formations}")
        with k2:
            st.metric("Produits uniques", f"{total_produits}")
        with k3:
            st.metric("Centres actifs", f"{total_centres}")
        with k4:
            st.metric("Moy. programmations", f"{avg_prog_global:.2f}")

        # Durée moyenne des formations si dates disponibles
        if date_debut_col and date_fin_col and date_debut_col in df_filtered.columns and date_fin_col in df_filtered.columns:
            # Calculer la différence en jours (soustraction de deux colonnes datetime donne des timedelta)
            df_durations = df_filtered[[date_debut_col, date_fin_col]].dropna()
            if len(df_durations) > 0:
                durations = (df_durations[date_fin_col] - df_durations[date_debut_col]).dt.days  # type: ignore
                avg_duration = durations.mean()
                if avg_duration > 0:
                    st.write(f"**Durée moyenne des formations :** {avg_duration:.1f} jours")

        # Tendance mensuelle des programmations
        if date_debut_col and date_debut_col in df_filtered.columns and df_filtered[date_debut_col].notna().any():
            st.subheader("📅 Tendance mensuelle des programmations")
            
            # Appliquer les mêmes filtres que pour recurrence_filtered
            df_month = df_filtered.copy()
            
            if regions_selectionnees and region_col and region_col in df_month.columns:
                df_month = df_month[df_month[region_col].isin(regions_selectionnees)]
            
            if cap_selectionnes and len(cap_selectionnes) > 0 and 'CAP_COMBINED' in df_month.columns:
                df_month = df_month[
                    df_month['CAP_COMBINED'].apply(
                        lambda x: any(cap in str(x) for cap in cap_selectionnes) if pd.notna(x) and cap_selectionnes else False
                    )
                ]
            
            if etablissements_selectionnes and etablissement_col and etablissement_col in df_month.columns:
                df_month = df_month[df_month[etablissement_col].isin(etablissements_selectionnes)]
            
            df_month['mois'] = df_month[date_debut_col].dt.to_period('M').dt.to_timestamp()  # type: ignore
            monthly_counts = df_month.groupby('mois').size().reset_index(name='nb_programmations')
            monthly_counts = monthly_counts.sort_values('mois')
            st.bar_chart(data=monthly_counts.set_index('mois'))
    
    else:
        st.warning("⚠️ Impossible de trouver les colonnes 'CENTRE' et/ou 'NUMERO DU PRODUIT' dans le fichier.")
        st.write("Colonnes disponibles :", df_filtered.columns.tolist())
    
    # Afficher le dataframe filtré complet
    st.subheader("📄 Détail des données filtrées")
    st.dataframe(df_filtered, use_container_width=True)

else:
    st.info("👆 Veuillez charger un fichier Excel pour commencer l'analyse.")
