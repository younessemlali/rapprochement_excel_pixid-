import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from thefuzz import fuzz, process
import re

# Configuration de la page
st.set_page_config(
    page_title="Excel Analyzer Pro",
    page_icon="📊",
    layout="wide"
)

# Titre principal
st.title("📊 Excel Analyzer Pro - Analyse intelligente de contrats")
st.markdown("### Embellissez, analysez et recherchez dans vos fichiers Excel")

# Fonction pour nettoyer les données
def clean_data(df):
    """Nettoie les données du DataFrame"""
    df = df.dropna(how='all')
    df = df.dropna(axis=1, how='all')
    
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()
    
    # Remplacer 'nan' par des chaînes vides
    df = df.replace('nan', '')
    df = df.fillna('')
    
    return df

# Fonction pour styliser une feuille Excel
def style_worksheet(worksheet, df):
    """Applique un style professionnel à une feuille Excel"""
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    even_row_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    odd_row_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    border_style = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )
    
    alignment_center = Alignment(horizontal='center', vertical='center')
    alignment_left = Alignment(horizontal='left', vertical='center')
    
    # Style des en-têtes
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = alignment_center
        cell.border = border_style
    
    # Style des lignes de données
    for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row), start=2):
        fill = even_row_fill if row_idx % 2 == 0 else odd_row_fill
        for cell in row:
            cell.fill = fill
            cell.border = border_style
            cell.alignment = alignment_left if cell.column <= 2 else alignment_center
    
    # Ajuster la largeur des colonnes
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Figer la première ligne
    worksheet.freeze_panes = 'A2'
    worksheet.auto_filter.ref = worksheet.dimensions

# Fonction pour créer un fichier Excel avec analyses complètes
def create_comprehensive_excel(df, filename="analyse_complete.xlsx"):
    """Crée un fichier Excel avec plusieurs onglets d'analyse"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # ONGLET 1: Données nettoyées
        df_clean = df.copy()
        df_clean.to_excel(writer, index=False, sheet_name='Données nettoyées')
        style_worksheet(writer.sheets['Données nettoyées'], df_clean)
        
        # ONGLET 2: Vue d'ensemble
        total = len(df)
        ok_count = len(df[df['Statut_Final'].str.upper() == 'OK'])
        ko_count = len(df[df['Statut_Final'].str.upper() != 'OK'])
        taux_reussite = round((ok_count / total * 100), 2) if total > 0 else 0
        
        initiaux = len(df[df['Initial/Avenant'].str.contains('Initial', case=False, na=False)])
        avenants = len(df[df['Initial/Avenant'].str.contains('Avenant', case=False, na=False)])
        unites = df['Code_Unite'].nunique() if 'Code_Unite' in df.columns else 0
        
        # Période
        if 'Date_Integration' in df.columns:
            df['Date_Integration'] = pd.to_datetime(df['Date_Integration'], errors='coerce')
            date_min = df['Date_Integration'].min()
            date_max = df['Date_Integration'].max()
            periode = f"Du {date_min.strftime('%d/%m/%Y') if pd.notna(date_min) else 'N/A'} au {date_max.strftime('%d/%m/%Y') if pd.notna(date_max) else 'N/A'}"
        else:
            periode = "N/A"
        
        summary_data = {
            'Métrique': [
                'Nombre total de contrats',
                'Nombre de contrats OK',
                'Nombre de contrats KO',
                'Taux de réussite (%)',
                'Nombre de contrats initiaux',
                'Nombre d\'avenants',
                'Nombre d\'agences (Code_Unite)',
                'Période couverte'
            ],
            'Valeur': [total, ok_count, ko_count, f"{taux_reussite}%", initiaux, avenants, unites, periode]
        }
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, index=False, sheet_name='Vue d\'ensemble')
        style_worksheet(writer.sheets['Vue d\'ensemble'], df_summary)
        
        # ONGLET 3: Analyse par agence (Code_Unite)
        if 'Code_Unite' in df.columns:
            ws_agence = writer.book.create_sheet('Analyse par agence')
            current_row = 1
            
            # Titre
            ws_agence.cell(row=current_row, column=1).value = "ANALYSE COMPLÈTE PAR AGENCE (CODE_UNITE)"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=14, color="366092")
            current_row += 2
            
            # 1. Volume total par agence
            volume_agence = df['Code_Unite'].value_counts().reset_index()
            volume_agence.columns = ['Agence', 'Nombre total']
            volume_agence['% du total'] = round((volume_agence['Nombre total'] / total * 100), 2)
            
            ws_agence.cell(row=current_row, column=1).value = "1. VOLUME TOTAL PAR AGENCE"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            
            for r_idx, row in enumerate(dataframe_to_rows(volume_agence, index=False, header=True), current_row):
                for c_idx, value in enumerate(row, 1):
                    ws_agence.cell(row=r_idx, column=c_idx).value = value
            current_row += len(volume_agence) + 3
            
            # 2. Taux de réussite par agence
            agence_list = []
            for agence in df['Code_Unite'].unique():
                df_agence = df[df['Code_Unite'] == agence]
                total_ag = len(df_agence)
                ok_ag = (df_agence['Statut_Final'].str.upper() == 'OK').sum()
                ko_ag = total_ag - ok_ag
                taux_ag = round((ok_ag / total_ag * 100), 2) if total_ag > 0 else 0
                agence_list.append({
                    'Code_Unite': agence,
                    'Total': total_ag,
                    'OK': ok_ag,
                    'KO': ko_ag,
                    'Taux réussite (%)': taux_ag
                })
            
            agence_status = pd.DataFrame(agence_list)
            agence_status = agence_status.sort_values('KO', ascending=False)
            
            ws_agence.cell(row=current_row, column=1).value = "2. TAUX DE RÉUSSITE PAR AGENCE (Classement par nombre de KO)"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            
            for r_idx, row in enumerate(dataframe_to_rows(agence_status, index=False, header=True), current_row):
                for c_idx, value in enumerate(row, 1):
                    ws_agence.cell(row=r_idx, column=c_idx).value = value
            current_row += len(agence_status) + 3
            
            # 3. Top agences avec le plus de rejets (KO)
            top_ko_agences = agence_status.nlargest(10, 'KO')[['Code_Unite', 'KO', 'Taux réussite (%)']]
            
            ws_agence.cell(row=current_row, column=1).value = "3. TOP 10 AGENCES AVEC LE PLUS DE REJETS"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12, color="DC3545")
            current_row += 1
            
            for r_idx, row in enumerate(dataframe_to_rows(top_ko_agences, index=False, header=True), current_row):
                for c_idx, value in enumerate(row, 1):
                    cell = ws_agence.cell(row=r_idx, column=c_idx)
                    cell.value = value
                    if r_idx == current_row:  # Header
                        cell.fill = PatternFill(start_color="DC3545", end_color="DC3545", fill_type="solid")
                        cell.font = Font(bold=True, color="FFFFFF")
            current_row += len(top_ko_agences) + 3
            
            # 4. Agences × Types d'erreurs
            df_ko = df[df['Statut_Final'].str.upper() != 'OK']
            if len(df_ko) > 0:
                try:
                    agence_erreur = pd.crosstab(df_ko['Code_Unite'], df_ko['Statut_Final'], margins=True)
                    agence_erreur = agence_erreur.reset_index()
                    
                    ws_agence.cell(row=current_row, column=1).value = "4. CROISEMENT AGENCES × TYPES D'ERREURS"
                    ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                    current_row += 1
                    
                    for r_idx, row in enumerate(dataframe_to_rows(agence_erreur, index=False, header=True), current_row):
                        for c_idx, value in enumerate(row, 1):
                            ws_agence.cell(row=r_idx, column=c_idx).value = value
                    current_row += len(agence_erreur) + 3
                except Exception as e:
                    # Si le crosstab échoue, on passe
                    ws_agence.cell(row=current_row, column=1).value = "4. CROISEMENT AGENCES × TYPES D'ERREURS - Données insuffisantes"
                    current_row += 2
            
            # 5. Agences × Types de contrats
            try:
                agence_type = pd.crosstab(df['Code_Unite'], df['Type (libellé)'], margins=True)
                agence_type = agence_type.reset_index()
                
                ws_agence.cell(row=current_row, column=1).value = "5. VOLUME D'INTÉGRATIONS PAR AGENCE ET TYPE DE CONTRAT"
                ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                current_row += 1
                
                for r_idx, row in enumerate(dataframe_to_rows(agence_type, index=False, header=True), current_row):
                    for c_idx, value in enumerate(row, 1):
                        ws_agence.cell(row=r_idx, column=c_idx).value = value
            except Exception as e:
                # Si le crosstab échoue, on passe
                ws_agence.cell(row=current_row, column=1).value = "5. VOLUME D'INTÉGRATIONS PAR AGENCE ET TYPE DE CONTRAT - Données insuffisantes"
        
        # ONGLET 4: Contrats OK
        if ok_count > 0:
            df_ok = df[df['Statut_Final'].str.upper() == 'OK'].copy()
            
            ok_by_type = df_ok['Type (libellé)'].value_counts().reset_index()
            ok_by_type.columns = ['Type de contrat', 'Nombre']
            ok_by_type['Pourcentage'] = round((ok_by_type['Nombre'] / ok_count * 100), 2)
            
            ok_by_unite = df_ok['Code_Unite'].value_counts().reset_index()
            ok_by_unite.columns = ['Agence (Code_Unite)', 'Nombre']
            ok_by_unite['Pourcentage'] = round((ok_by_unite['Nombre'] / ok_count * 100), 2)
            
            ok_summary = pd.DataFrame({
                'Métrique': ['Total contrats OK', 'Nombre de types différents', 'Nombre d\'agences'],
                'Valeur': [ok_count, df_ok['Type (libellé)'].nunique(), df_ok['Code_Unite'].nunique()]
            })
            
            ok_summary.to_excel(writer, index=False, sheet_name='Contrats OK', startrow=0)
            ok_by_type.to_excel(writer, index=False, sheet_name='Contrats OK', startrow=len(ok_summary) + 3)
            ok_by_unite.to_excel(writer, index=False, sheet_name='Contrats OK', startrow=len(ok_summary) + len(ok_by_type) + 6)
            
            ws_ok = writer.sheets['Contrats OK']
            ws_ok.insert_rows(len(ok_summary) + 2)
            ws_ok.cell(row=len(ok_summary) + 2, column=1).value = "Répartition par type de contrat:"
            ws_ok.cell(row=len(ok_summary) + 2, column=1).font = Font(bold=True, size=12)
            
            ws_ok.insert_rows(len(ok_summary) + len(ok_by_type) + 5)
            ws_ok.cell(row=len(ok_summary) + len(ok_by_type) + 5, column=1).value = "Répartition par agence:"
            ws_ok.cell(row=len(ok_summary) + len(ok_by_type) + 5, column=1).font = Font(bold=True, size=12)
            
            style_worksheet(ws_ok, ok_summary)
        
        # ONGLET 5: Contrats KO avec analyses détaillées
        if ko_count > 0:
            df_ko = df[df['Statut_Final'].str.upper() != 'OK'].copy()
            
            ko_summary = pd.DataFrame({
                'Métrique': ['Total contrats KO', 'Taux d\'échec (%)', 'Nombre de types d\'erreurs', 'Nombre d\'agences concernées'],
                'Valeur': [
                    ko_count,
                    f"{round((ko_count / total * 100), 2)}%",
                    df_ko['Statut_Final'].nunique(),
                    df_ko['Code_Unite'].nunique()
                ]
            })
            
            ko_by_status = df_ko['Statut_Final'].value_counts().reset_index()
            ko_by_status.columns = ['Type d\'erreur', 'Nombre']
            ko_by_status['Pourcentage'] = round((ko_by_status['Nombre'] / ko_count * 100), 2)
            
            # KO par agence
            ko_by_agence = df_ko['Code_Unite'].value_counts().reset_index()
            ko_by_agence.columns = ['Agence', 'Nombre de rejets']
            ko_by_agence['% des rejets'] = round((ko_by_agence['Nombre de rejets'] / ko_count * 100), 2)
            
            # Messages d'erreur
            error_messages = []
            if 'Message_Integration' in df_ko.columns:
                msg_int = df_ko[df_ko['Message_Integration'] != '']['Message_Integration'].value_counts()
                if len(msg_int) > 0:
                    error_messages.append(('Message_Integration', msg_int))
            
            if 'Message_Transfert' in df_ko.columns:
                msg_trans = df_ko[df_ko['Message_Transfert'] != '']['Message_Transfert'].value_counts()
                if len(msg_trans) > 0:
                    error_messages.append(('Message_Transfert', msg_trans))
            
            ko_by_type = df_ko['Type (libellé)'].value_counts().reset_index()
            ko_by_type.columns = ['Type de contrat', 'Nombre KO']
            
            # Écrire dans l'onglet
            current_row = 0
            ko_summary.to_excel(writer, index=False, sheet_name='Contrats KO', startrow=current_row)
            current_row += len(ko_summary) + 3
            
            ws_ko = writer.sheets['Contrats KO']
            ws_ko.cell(row=current_row, column=1).value = "RÉPARTITION DES ERREURS PAR STATUT:"
            ws_ko.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            ko_by_status.to_excel(writer, index=False, sheet_name='Contrats KO', startrow=current_row)
            current_row += len(ko_by_status) + 3
            
            ws_ko.cell(row=current_row, column=1).value = "REJETS PAR AGENCE (TOP):"
            ws_ko.cell(row=current_row, column=1).font = Font(bold=True, size=12, color="DC3545")
            current_row += 1
            ko_by_agence.to_excel(writer, index=False, sheet_name='Contrats KO', startrow=current_row)
            current_row += len(ko_by_agence) + 3
            
            # Messages d'erreur
            for msg_type, msg_counts in error_messages:
                ws_ko.cell(row=current_row, column=1).value = f"MESSAGES D'ERREUR - {msg_type}:"
                ws_ko.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                current_row += 1
                
                msg_df = pd.DataFrame({
                    'Message': msg_counts.index,
                    'Occurrences': msg_counts.values
                }).head(15)
                msg_df.to_excel(writer, index=False, sheet_name='Contrats KO', startrow=current_row)
                current_row += len(msg_df) + 3
            
            ws_ko.cell(row=current_row, column=1).value = "CONTRATS KO PAR TYPE:"
            ws_ko.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            ko_by_type.to_excel(writer, index=False, sheet_name='Contrats KO', startrow=current_row)
            
            style_worksheet(ws_ko, ko_summary)
        
        # ONGLET 6: Types et Avenants
        if 'Initial/Avenant' in df.columns and 'Type (libellé)' in df.columns:
            init_avenant = df['Initial/Avenant'].value_counts().reset_index()
            init_avenant.columns = ['Catégorie', 'Nombre']
            init_avenant['Pourcentage'] = round((init_avenant['Nombre'] / total * 100), 2)
            
            types_detail = df['Type (libellé)'].value_counts().reset_index()
            types_detail.columns = ['Type de contrat', 'Nombre']
            types_detail['Pourcentage'] = round((types_detail['Nombre'] / total * 100), 2)
            
            cross_type_status = None
            if 'Statut_Final' in df.columns:
                try:
                    cross_type_status = pd.crosstab(
                        df['Type (libellé)'],
                        df['Statut_Final'],
                        margins=True,
                        margins_name='Total'
                    ).reset_index()
                except Exception:
                    pass
            
            current_row = 0
            init_avenant.to_excel(writer, index=False, sheet_name='Types et Avenants', startrow=current_row)
            ws_types = writer.sheets['Types et Avenants']
            current_row += len(init_avenant) + 3
            
            ws_types.cell(row=current_row, column=1).value = "Détail par type de contrat:"
            ws_types.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            types_detail.to_excel(writer, index=False, sheet_name='Types et Avenants', startrow=current_row)
            current_row += len(types_detail) + 3
            
            if cross_type_status is not None:
                ws_types.cell(row=current_row, column=1).value = "Croisement Type × Statut:"
                ws_types.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                current_row += 1
                cross_type_status.to_excel(writer, index=False, sheet_name='Types et Avenants', startrow=current_row)
            
            style_worksheet(ws_types, init_avenant)
        
        # ONGLET 7: Analyse temporelle
        if 'Date_Integration' in df.columns:
            df_temp = df.copy()
            df_temp['Date_Integration'] = pd.to_datetime(df_temp['Date_Integration'], errors='coerce')
            df_temp = df_temp.dropna(subset=['Date_Integration'])
            
            if len(df_temp) > 0:
                df_temp['Date'] = df_temp['Date_Integration'].dt.date
                timeline_day = df_temp.groupby('Date').size().reset_index(name='Nombre de contrats')
                
                df_temp['Mois'] = df_temp['Date_Integration'].dt.to_period('M').astype(str)
                timeline_month = df_temp.groupby('Mois').size().reset_index(name='Nombre de contrats')
                
                current_row = 0
                pd.DataFrame({
                    'Métrique': ['Date la plus ancienne', 'Date la plus récente', 'Nombre de jours couverts'],
                    'Valeur': [
                        df_temp['Date_Integration'].min().strftime('%d/%m/%Y'),
                        df_temp['Date_Integration'].max().strftime('%d/%m/%Y'),
                        (df_temp['Date_Integration'].max() - df_temp['Date_Integration'].min()).days
                    ]
                }).to_excel(writer, index=False, sheet_name='Analyse temporelle', startrow=current_row)
                ws_time = writer.sheets['Analyse temporelle']
                current_row += 5
                
                ws_time.cell(row=current_row, column=1).value = "Volume par jour:"
                ws_time.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                current_row += 1
                timeline_day.to_excel(writer, index=False, sheet_name='Analyse temporelle', startrow=current_row)
                current_row += len(timeline_day) + 3
                
                ws_time.cell(row=current_row, column=1).value = "Volume par mois:"
                ws_time.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                current_row += 1
                timeline_month.to_excel(writer, index=False, sheet_name='Analyse temporelle', startrow=current_row)
                
                style_worksheet(ws_time, timeline_day)
    
    output.seek(0)
    return output

# Upload du fichier
uploaded_file = st.file_uploader(
    "📁 Choisissez votre fichier Excel",
    type=['xlsx', 'xls'],
    help="Formats supportés: .xlsx, .xls"
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        with st.expander("👁️ Aperçu des données brutes", expanded=False):
            st.dataframe(df.head(10), width='stretch')
        
        df_clean = clean_data(df)
        
        st.success(f"✅ Fichier chargé avec succès : {len(df_clean)} lignes, {len(df_clean.columns)} colonnes")
        
        # Créer des onglets
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "🔍 Recherche intelligente",
            "📋 Données", 
            "📊 Analyses détaillées", 
            "📈 Visualisations", 
            "💾 Export multi-onglets"
        ])
        
        # TAB 0: Recherche intelligente de contrats
        with tab1:
            st.subheader("🔍 Recherche intelligente de contrats")
            st.markdown("Recherchez des contrats par numéro, agence, statut, type ou tout autre critère")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                search_column = st.selectbox(
                    "Rechercher dans la colonne:",
                    options=df_clean.columns.tolist(),
                    index=df_clean.columns.tolist().index('Contrat') if 'Contrat' in df_clean.columns else 0
                )
            
            with col2:
                search_value = st.text_input("Valeur recherchée:", placeholder="Ex: 001-NVM-173169")
            
            with col3:
                search_type = st.radio("Type de recherche:", ["Contient", "Égal à", "Commence par"], horizontal=True)
            
            if search_value:
                if search_type == "Contient":
                    mask = df_clean[search_column].astype(str).str.contains(search_value, case=False, na=False)
                elif search_type == "Égal à":
                    mask = df_clean[search_column].astype(str).str.upper() == search_value.upper()
                else:  # Commence par
                    mask = df_clean[search_column].astype(str).str.startswith(search_value, na=False)
                
                results = df_clean[mask]
                
                st.markdown(f"### Résultats de la recherche: **{len(results)}** contrat(s) trouvé(s)")
                
                if len(results) > 0:
                    st.dataframe(results, width='stretch', height=400)
                    
                    # Bouton pour exporter les résultats
                    csv = results.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Télécharger les résultats (CSV)",
                        data=csv,
                        file_name=f"recherche_{search_value}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
                    
                    # Statistiques rapides sur les résultats
                    if len(results) > 1:
                        st.markdown("#### Statistiques sur les résultats")
                        col1, col2, col3 = st.columns(3)
                        
                        if 'Statut_Final' in results.columns:
                            with col1:
                                ok_res = len(results[results['Statut_Final'].str.upper() == 'OK'])
                                st.metric("OK", ok_res, delta=f"{round(ok_res/len(results)*100, 1)}%")
                        
                        if 'Code_Unite' in results.columns:
                            with col2:
                                st.metric("Agences", results['Code_Unite'].nunique())
                        
                        if 'Type (libellé)' in results.columns:
                            with col3:
                                st.metric("Types", results['Type (libellé)'].nunique())
                else:
                    st.warning(f"Aucun résultat trouvé pour '{search_value}' dans la colonne '{search_column}'")
            
            # Recherche avancée
            with st.expander("🎯 Recherche avancée (filtres multiples)", expanded=False):
                st.markdown("Combinez plusieurs critères pour affiner votre recherche")
                
                filter_col1, filter_col2 = st.columns(2)
                
                with filter_col1:
                    if 'Statut_Final' in df_clean.columns:
                        statuts_select = st.multiselect(
                            "Filtrer par statut:",
                            options=df_clean['Statut_Final'].unique().tolist(),
                            default=None
                        )
                    
                    if 'Type (libellé)' in df_clean.columns:
                        types_select = st.multiselect(
                            "Filtrer par type:",
                            options=df_clean['Type (libellé)'].unique().tolist(),
                            default=None
                        )
                
                with filter_col2:
                    if 'Code_Unite' in df_clean.columns:
                        agences_select = st.multiselect(
                            "Filtrer par agence:",
                            options=sorted(df_clean['Code_Unite'].unique().tolist()),
                            default=None
                        )
                    
                    if 'Initial/Avenant' in df_clean.columns:
                        init_avenant_select = st.multiselect(
                            "Filtrer par Initial/Avenant:",
                            options=df_clean['Initial/Avenant'].unique().tolist(),
                            default=None
                        )
                
                if st.button("🔍 Appliquer les filtres avancés", width='stretch'):
                    filtered_df = df_clean.copy()
                    
                    if 'statuts_select' in locals() and statuts_select:
                        filtered_df = filtered_df[filtered_df['Statut_Final'].isin(statuts_select)]
                    
                    if 'types_select' in locals() and types_select:
                        filtered_df = filtered_df[filtered_df['Type (libellé)'].isin(types_select)]
                    
                    if 'agences_select' in locals() and agences_select:
                        filtered_df = filtered_df[filtered_df['Code_Unite'].isin(agences_select)]
                    
                    if 'init_avenant_select' in locals() and init_avenant_select:
                        filtered_df = filtered_df[filtered_df['Initial/Avenant'].isin(init_avenant_select)]
                    
                    st.success(f"✅ {len(filtered_df)} contrats correspondent aux critères")
                    st.dataframe(filtered_df, width='stretch', height=400)
                    
                    # Export des résultats filtrés
                    csv_filtered = filtered_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Télécharger les résultats filtrés (CSV)",
                        data=csv_filtered,
                        file_name=f"recherche_filtree_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
        
        # TAB 1: Données nettoyées
        with tab2:
            st.subheader("Données nettoyées et formatées")
            st.dataframe(df_clean, width='stretch', height=400)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Lignes", len(df_clean))
            with col2:
                st.metric("Colonnes", len(df_clean.columns))
            with col3:
                duplicates = df_clean.duplicated().sum()
                st.metric("Doublons", duplicates)
            with col4:
                if 'Statut_Final' in df_clean.columns:
                    ok_count = len(df_clean[df_clean['Statut_Final'].str.upper() == 'OK'])
                    st.metric("Contrats OK", ok_count)
        
        # TAB 2: Analyses détaillées
        with tab3:
            st.subheader("📊 Analyses approfondies")
            
            # Analyse des statuts
            if 'Statut_Final' in df_clean.columns:
                st.markdown("### 🎯 Analyse des statuts")
                col1, col2, col3 = st.columns(3)
                
                total = len(df_clean)
                ok_count = len(df_clean[df_clean['Statut_Final'].str.upper() == 'OK'])
                ko_count = total - ok_count
                
                with col1:
                    st.metric("Total contrats", total)
                with col2:
                    st.metric("✅ OK", ok_count, delta=f"{round(ok_count/total*100, 1)}%")
                with col3:
                    st.metric("❌ KO", ko_count, delta=f"{round(ko_count/total*100, 1)}%", delta_color="inverse")
                
                # Détail des erreurs
                if ko_count > 0:
                    st.markdown("#### 🔴 Détail des erreurs par type")
                    df_ko = df_clean[df_clean['Statut_Final'].str.upper() != 'OK']
                    error_types = df_ko['Statut_Final'].value_counts().reset_index()
                    error_types.columns = ['Type d\'erreur', 'Nombre']
                    error_types['%'] = round(error_types['Nombre'] / ko_count * 100, 1)
                    st.dataframe(error_types, width='stretch', hide_index=True)
            
            # Analyse par agence (Code_Unite)
            if 'Code_Unite' in df_clean.columns:
                st.markdown("### 🏢 Analyse par agence (Code_Unite)")
                
                # Top agences avec rejets
                if 'Statut_Final' in df_clean.columns:
                    # Créer le DataFrame d'analyse par agence
                    agence_list = []
                    for agence in df_clean['Code_Unite'].unique():
                        df_agence = df_clean[df_clean['Code_Unite'] == agence]
                        total_ag = len(df_agence)
                        ok_ag = (df_agence['Statut_Final'].str.upper() == 'OK').sum()
                        ko_ag = total_ag - ok_ag
                        taux_ag = round((ok_ag / total_ag * 100), 1) if total_ag > 0 else 0
                        agence_list.append({
                            'Code_Unite': agence,
                            'Total': total_ag,
                            'OK': ok_ag,
                            'KO': ko_ag,
                            'Taux réussite (%)': taux_ag
                        })
                    
                    agence_status = pd.DataFrame(agence_list)
                    agence_status = agence_status.sort_values('KO', ascending=False)
                    
                    st.markdown("#### 🔴 Top 10 agences avec le plus de rejets")
                    top_rejets = agence_status.head(10)
                    st.dataframe(top_rejets, width='stretch', hide_index=True)
                    
                    st.markdown("#### ✅ Top 10 agences avec le meilleur taux de réussite")
                    top_reussite = agence_status.sort_values('Taux réussite (%)', ascending=False).head(10)
                    st.dataframe(top_reussite, width='stretch', hide_index=True)
                
                # Volume par agence
                st.markdown("#### 📊 Volume d'intégrations par agence")
                volume_agence = df_clean['Code_Unite'].value_counts().reset_index()
                volume_agence.columns = ['Agence', 'Nombre']
                volume_agence['%'] = round(volume_agence['Nombre'] / len(df_clean) * 100, 1)
                st.dataframe(volume_agence, width='stretch', hide_index=True)
            
            # Analyse Initial/Avenant
            if 'Initial/Avenant' in df_clean.columns:
                st.markdown("### 📄 Analyse Initial vs Avenants")
                init_avenant = df_clean['Initial/Avenant'].value_counts()
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Contrats Initiaux", init_avenant.get('Initial', 0))
                with col2:
                    st.metric("Avenants", init_avenant.get('Avenant', 0))
            
            # Analyse des types
            if 'Type (libellé)' in df_clean.columns:
                st.markdown("### 📋 Répartition par type de contrat")
                types_count = df_clean['Type (libellé)'].value_counts().reset_index()
                types_count.columns = ['Type', 'Nombre']
                types_count['%'] = round(types_count['Nombre'] / len(df_clean) * 100, 1)
                st.dataframe(types_count, width='stretch', hide_index=True)
            
            # Croisement Agences × Types d'erreurs
            if 'Code_Unite' in df_clean.columns and 'Statut_Final' in df_clean.columns and ko_count > 0:
                st.markdown("### 🔀 Croisement Agences × Types d'erreurs")
                df_ko_only = df_clean[df_clean['Statut_Final'].str.upper() != 'OK']
                try:
                    cross_agence_erreur = pd.crosstab(
                        df_ko_only['Code_Unite'],
                        df_ko_only['Statut_Final'],
                        margins=True,
                        margins_name='Total'
                    )
                    st.dataframe(cross_agence_erreur, width='stretch')
                except Exception as e:
                    st.warning("Impossible de générer le croisement - données insuffisantes")
        
        # TAB 3: Visualisations
        with tab4:
            st.subheader("📈 Visualisations interactives")
            
            col1, col2 = st.columns(2)
            
            # Graphique statuts OK/KO
            if 'Statut_Final' in df_clean.columns:
                with col1:
                    st.markdown("#### Distribution OK vs KO")
                    ok_count = len(df_clean[df_clean['Statut_Final'].str.upper() == 'OK'])
                    ko_count = len(df_clean) - ok_count
                    fig = px.pie(
                        values=[ok_count, ko_count],
                        names=['OK', 'KO'],
                        title="Répartition Statut Final",
                        hole=0.4,
                        color_discrete_map={'OK': '#28a745', 'KO': '#dc3545'}
                    )
                    st.plotly_chart(fig, width='stretch')
            
            # Graphique types
            if 'Type (libellé)' in df_clean.columns:
                with col2:
                    st.markdown("#### Types de contrats")
                    type_counts = df_clean['Type (libellé)'].value_counts()
                    fig = px.bar(
                        x=type_counts.index,
                        y=type_counts.values,
                        title="Nombre par type",
                        labels={'x': 'Type', 'y': 'Nombre'},
                        color=type_counts.values,
                        color_continuous_scale='Blues'
                    )
                    fig.update_layout(showlegend=False)
                    st.plotly_chart(fig, width='stretch')
            
            # Graphiques par agence
            if 'Code_Unite' in df_clean.columns:
                st.markdown("#### 🏢 Analyse par agence")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Volume par agence
                    volume_agence = df_clean['Code_Unite'].value_counts().head(15)
                    fig = px.bar(
                        x=volume_agence.values,
                        y=volume_agence.index,
                        orientation='h',
                        title="Top 15 agences par volume",
                        labels={'x': 'Nombre de contrats', 'y': 'Agence'},
                        color=volume_agence.values,
                        color_continuous_scale='Viridis'
                    )
                    fig.update_layout(showlegend=False, yaxis={'categoryorder':'total ascending'})
                    st.plotly_chart(fig, width='stretch')
                
                with col2:
                    # Taux de réussite par agence
                    if 'Statut_Final' in df_clean.columns:
                        agence_success = df_clean.groupby('Code_Unite')['Statut_Final'].apply(
                            lambda x: (x.str.upper() == 'OK').sum() / len(x) * 100
                        ).sort_values(ascending=False).head(15)
                        
                        fig = px.bar(
                            x=agence_success.values,
                            y=agence_success.index,
                            orientation='h',
                            title="Top 15 agences - Taux de réussite (%)",
                            labels={'x': 'Taux de réussite (%)', 'y': 'Agence'},
                            color=agence_success.values,
                            color_continuous_scale='RdYlGn'
                        )
                        fig.update_layout(showlegend=False, yaxis={'categoryorder':'total ascending'})
                        st.plotly_chart(fig, width='stretch')
            
            # Évolution temporelle
            if 'Date_Integration' in df_clean.columns:
                st.markdown("#### 📅 Évolution temporelle")
                df_temp = df_clean.copy()
                df_temp['Date_Integration'] = pd.to_datetime(df_temp['Date_Integration'], errors='coerce')
                df_temp = df_temp.dropna(subset=['Date_Integration'])
                df_temp = df_temp.copy()  # Créer une vraie copie pour éviter le warning
                df_temp['Date'] = df_temp['Date_Integration'].dt.date
                timeline = df_temp.groupby('Date').size().reset_index(name='Nombre')
                
                fig = px.line(
                    timeline,
                    x='Date',
                    y='Nombre',
                    title="Volume de contrats par jour",
                    markers=True
                )
                st.plotly_chart(fig, width='stretch')
            
            # Analyse croisée Type × Statut
            if 'Type (libellé)' in df_clean.columns and 'Statut_Final' in df_clean.columns:
                st.markdown("#### 🔀 Analyse croisée Type × Statut")
                try:
                    cross_data = pd.crosstab(df_clean['Type (libellé)'], df_clean['Statut_Final'])
                    fig = px.bar(
                        cross_data,
                        barmode='group',
                        title="Répartition des statuts par type de contrat"
                    )
                    st.plotly_chart(fig, width='stretch')
                except Exception as e:
                    st.warning("Impossible de générer le graphique croisé - données insuffisantes")
            
            # Heatmap Agences × Types d'erreurs
            if 'Code_Unite' in df_clean.columns and 'Statut_Final' in df_clean.columns and ko_count > 0:
                st.markdown("#### 🔥 Heatmap : Agences × Types d'erreurs")
                try:
                    df_ko_heat = df_clean[df_clean['Statut_Final'].str.upper() != 'OK']
                    
                    # Limiter aux top agences pour la lisibilité
                    top_agences_ko = df_ko_heat['Code_Unite'].value_counts().head(10).index
                    df_ko_heat = df_ko_heat[df_ko_heat['Code_Unite'].isin(top_agences_ko)]
                    
                    heatmap_data = pd.crosstab(df_ko_heat['Code_Unite'], df_ko_heat['Statut_Final'])
                    
                    fig = px.imshow(
                        heatmap_data,
                        labels=dict(x="Type d'erreur", y="Agence", color="Nombre"),
                        title="Concentration des erreurs par agence (Top 10)",
                        color_continuous_scale='Reds',
                        aspect="auto"
                    )
                    st.plotly_chart(fig, width='stretch')
                except Exception as e:
                    st.warning("Impossible de générer la heatmap - données insuffisantes")
        
        # TAB 4: Export multi-onglets
        with tab5:
            st.subheader("💾 Télécharger l'analyse complète Excel")
            
            st.markdown("""
            ### 📑 Le fichier Excel généré contient 7 onglets d'analyse :
            
            1. **Données nettoyées** - Toutes vos données formatées et nettoyées
            2. **Vue d'ensemble** - Métriques clés et statistiques générales
            3. **🆕 Analyse par agence** - Analyses complètes des performances par Code_Unite
            4. **Contrats OK** - Analyse détaillée des contrats réussis par type et agence
            5. **Contrats KO** - Analyse des erreurs avec messages détaillés + rejets par agence
            6. **Types et Avenants** - Répartition des types de contrats et avenants
            7. **Analyse temporelle** - Évolution dans le temps
            """)
            
            excel_file = create_comprehensive_excel(df_clean)
            
            st.download_button(
                label="⬇️ Télécharger l'analyse Excel complète (7 onglets)",
                data=excel_file,
                file_name=f"analyse_complete_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch'
            )
            
            st.success("✅ Fichier Excel avec 7 onglets d'analyse prêt au téléchargement !")
    
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier : {str(e)}")
        st.exception(e)
        st.info("Vérifiez que votre fichier Excel est valide et non corrompu.")

else:
    st.info("👆 Commencez par uploader un fichier Excel pour l'analyser")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666;'>Excel Analyzer Pro - Analyse intelligente de contrats</div>",
    unsafe_allow_html=True
)
