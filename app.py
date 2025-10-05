# TAB 2: Dashboard Agences
        with tab3:
            st.subheader("🏢 Dashboard Agences - Vue Exécutive")
            
            if 'Code_Unite' in df_clean.columns and 'Statut_Final' in df_clean.columns:
                
                # Calcul des métriques par agence
                agence_metrics = []
                for agence in df_clean['Code_Unite'].unique():
                    df_ag = df_clean[df_clean['Code_Unite'] == agence]
                    total = len(df_ag)
                    ok = (df_ag['Statut_Final'].str.upper() == 'OK').sum()
                    ko = total - ok
                    taux = round((ok / total * 100), 1) if total > 0 else 0
                    agence_metrics.append({
                        'Agence': agence,
                        'Total': total,
                        'OK': ok,
                        'KO': ko,
                        'Taux (%)': taux
                    })
                
                df_agences = pd.DataFrame(agence_metrics)
                taux_moyen = df_agences['Taux (%)'].mean()
                df_agences['Écart vs Moyenne'] = df_agences['Taux (%)'] - taux_moyen
                df_agences = df_agences.sort_values('Taux (%)', ascending=False)
                
                # === MÉTRIQUES CLÉS ===
                st.markdown("### 🎯 Métriques Clés")
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    st.metric(
                        "🏆 Meilleure",
                        df_agences.iloc[0]['Agence'],
                        f"{df_agences.iloc[0]['Taux (%)']}%"
                    )
                
                with col2:
                    st.metric(
                        "🔴 Pire",
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

# Fonction de parsing de langage naturel
def parse_natural_language_query(query, df):
    """Parse une requête en langage naturel et extrait les filtres"""
    filters = {}
    query_lower = query.lower()
    
    # Détecter les statuts
    if any(word in query_lower for word in ['ko', 'échec', 'erreur', 'rejet', 'échoué']):
        filters['statut'] = 'KO'
    elif any(word in query_lower for word in ['ok', 'réussi', 'succès', 'validé']):
        filters['statut'] = 'OK'
    
    # Détecter les agences (Code_Unite)
    if 'Code_Unite' in df.columns:
        agences = df['Code_Unite'].unique()
        for agence in agences:
            if str(agence).lower() in query_lower:
                filters['agence'] = agence
                break
    
    # Détecter les mois
    mois_map = {
        'janvier': 1, 'jan': 1,
        'février': 2, 'fevrier': 2, 'fév': 2, 'fev': 2,
        'mars': 3, 'mar': 3,
        'avril': 4, 'avr': 4,
        'mai': 5,
        'juin': 6,
        'juillet': 7, 'juil': 7,
        'août': 8, 'aout': 8,
        'septembre': 9, 'sept': 9, 'sep': 9,
        'octobre': 10, 'oct': 10,
        'novembre': 11, 'nov': 11,
        'décembre': 12, 'decembre': 12, 'déc': 12, 'dec': 12
    }
    
    for nom_mois, num_mois in mois_map.items():
        if nom_mois in query_lower:
            filters['mois'] = num_mois
            break
    
    # Détecter Initial/Avenant
    if any(word in query_lower for word in ['initial', 'initiaux']):
        filters['init_avenant'] = 'Initial'
    elif any(word in query_lower for word in ['avenant', 'avenants']):
        filters['init_avenant'] = 'Avenant'
    
    # Détecter les types de contrats
    if 'Type (libellé)' in df.columns:
        types = df['Type (libellé)'].unique()
        for type_contrat in types:
            if str(type_contrat).lower() in query_lower:
                filters['type'] = type_contrat
                break
    
    return filters

# Fonction de recherche floue
def fuzzy_search(query, df, column, limit=10):
    """Recherche floue dans une colonne spécifique"""
    if column not in df.columns:
        return []
    
    # Extraire les valeurs uniques non vides
    values = df[column].dropna().astype(str).unique().tolist()
    values = [v for v in values if v.strip()]
    
    if not values or not query.strip():
        return []
    
    # Recherche floue
    matches = process.extract(query, values, limit=limit, scorer=fuzz.token_sort_ratio)
    
    # Retourner les résultats avec leur score
    return [(match[0], match[1]) for match in matches if match[1] > 50]  # Score minimum 50

# Fonction pour calculer le score de pertinence
def calculate_relevance_score(row, query, filters):
    """Calcule un score de pertinence pour chaque ligne"""
    score = 0
    query_lower = query.lower()
    
    # Score basé sur le contrat
    if 'Contrat' in row.index:
        contrat_str = str(row['Contrat']).lower()
        if query_lower in contrat_str:
            score += 100  # Correspondance exacte
        else:
            score += fuzz.partial_ratio(query_lower, contrat_str) * 0.5  # Correspondance partielle
    
    # Score basé sur les filtres détectés
    if filters.get('agence') and 'Code_Unite' in row.index:
        if row['Code_Unite'] == filters['agence']:
            score += 50
    
    if filters.get('statut') and 'Statut_Final' in row.index:
        if filters['statut'] == 'KO' and row['Statut_Final'].upper() != 'OK':
            score += 50
        elif filters['statut'] == 'OK' and row['Statut_Final'].upper() == 'OK':
            score += 50
    
    if filters.get('type') and 'Type (libellé)' in row.index:
        if row['Type (libellé)'] == filters['type']:
            score += 40
    
    if filters.get('init_avenant') and 'Initial/Avenant' in row.index:
        if filters['init_avenant'].lower() in str(row['Initial/Avenant']).lower():
            score += 30
    
    if filters.get('mois') and 'Date_Integration' in row.index:
        try:
            date = pd.to_datetime(row['Date_Integration'])
            if date.month == filters['mois']:
                score += 40
        except:
            pass
    
    return score

# Fonction pour obtenir des suggestions intelligentes
def get_smart_suggestions(partial_input, df, limit=5):
    """Génère des suggestions intelligentes basées sur l'entrée partielle"""
    suggestions = []
    
    if not partial_input or len(partial_input) < 2:
        return suggestions
    
    partial_lower = partial_input.lower()
    
    # Suggestions de contrats
    if 'Contrat' in df.columns:
        contrats = df['Contrat'].dropna().astype(str)
        contrats_matches = contrats[contrats.str.contains(partial_input, case=False, na=False)].head(limit)
        for contrat in contrats_matches:
            suggestions.append({
                'type': '📄 Contrat',
                'value': contrat,
                'score': fuzz.partial_ratio(partial_lower, contrat.lower())
            })
    
    # Suggestions d'agences
    if 'Code_Unite' in df.columns:
        agences = df['Code_Unite'].dropna().astype(str).unique()
        for agence in agences:
            if partial_lower in agence.lower():
                suggestions.append({
                    'type': '🏢 Agence',
                    'value': agence,
                    'score': fuzz.ratio(partial_lower, agence.lower())
                })
    
    # Suggestions de statuts
    if any(word in partial_lower for word in ['k', 'o', 'ko', 'ok']):
        if 'ko' in partial_lower or 'k' == partial_lower:
            suggestions.append({'type': '❌ Statut', 'value': 'KO', 'score': 100})
        if 'ok' in partial_lower or 'o' == partial_lower:
            suggestions.append({'type': '✅ Statut', 'value': 'OK', 'score': 100})
    
    # Suggestions de mois
    mois_suggestions = {
        'jan': 'janvier', 'fev': 'février', 'mar': 'mars', 'avr': 'avril',
        'mai': 'mai', 'juin': 'juin', 'juil': 'juillet', 'aout': 'août',
        'sept': 'septembre', 'oct': 'octobre', 'nov': 'novembre', 'dec': 'décembre'
    }
    for abbr, mois in mois_suggestions.items():
        if abbr.startswith(partial_lower) or mois.startswith(partial_lower):
            suggestions.append({'type': '📅 Mois', 'value': mois, 'score': 90})
    
    # Trier par score et limiter
    suggestions = sorted(suggestions, key=lambda x: x['score'], reverse=True)[:limit]
    
    return suggestions

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
        
        # ONGLET 3: Analyse par agence (Code_Unite) - VERSION ENRICHIE
        if 'Code_Unite' in df.columns:
            ws_agence = writer.book.create_sheet('Analyse par agence')
            current_row = 1
            
            # TITRE PRINCIPAL
            ws_agence.cell(row=current_row, column=1).value = "ANALYSE COMPLÈTE PAR AGENCE (CODE_UNITE)"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=14, color="366092")
            ws_agence.merge_cells(f'A{current_row}:F{current_row}')
            current_row += 2
            
            # === SECTION 0: DASHBOARD EXÉCUTIF ===
            ws_agence.cell(row=current_row, column=1).value = "🎯 DASHBOARD EXÉCUTIF"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=13, color="FF0000")
            current_row += 1
            
            # Calculer les métriques globales
            agence_list_exec = []
            for agence in df['Code_Unite'].unique():
                df_agence = df[df['Code_Unite'] == agence]
                total_ag = len(df_agence)
                ok_ag = (df_agence['Statut_Final'].str.upper() == 'OK').sum()
                ko_ag = total_ag - ok_ag
                taux_ag = round((ok_ag / total_ag * 100), 2) if total_ag > 0 else 0
                agence_list_exec.append({
                    'Agence': agence,
                    'Total': total_ag,
                    'OK': ok_ag,
                    'KO': ko_ag,
                    'Taux': taux_ag
                })
            
            df_exec = pd.DataFrame(agence_list_exec)
            
            # Métriques clés
            taux_moyen = df_exec['Taux'].mean()
            meilleure_agence = df_exec.loc[df_exec['Taux'].idxmax()]
            pire_agence = df_exec.loc[df_exec['Taux'].idxmin()]
            agences_alerte = len(df_exec[df_exec['Taux'] < 60])
            agences_au_dessus = len(df_exec[df_exec['Taux'] >= taux_moyen])
            
            synthese = pd.DataFrame({
                'Indicateur': [
                    '🏆 Meilleure agence',
                    '🔴 Pire agence',
                    '📊 Taux moyen national',
                    '⚠️ Agences en alerte (< 60%)',
                    '✅ Agences au-dessus moyenne',
                    '📈 Total agences'
                ],
                'Valeur': [
                    f"{meilleure_agence['Agence']} ({meilleure_agence['Taux']:.1f}%)",
                    f"{pire_agence['Agence']} ({pire_agence['Taux']:.1f}%)",
                    f"{taux_moyen:.1f}%",
                    str(agences_alerte),
                    f"{agences_au_dessus}/{len(df_exec)}",
                    str(len(df_exec))
                ]
            })
            
            for r_idx, row in enumerate(dataframe_to_rows(synthese, index=False, header=True), current_row):
                for c_idx, value in enumerate(row, 1):
                    cell = ws_agence.cell(row=r_idx, column=c_idx)
                    cell.value = value
                    if r_idx == current_row:  # Header
                        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                        cell.font = Font(bold=True, color="FFFFFF")
                    else:
                        cell.border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin')
                        )
            
            current_row += len(synthese) + 3
            
            # === SECTION 1: CLASSEMENT GÉNÉRAL ===
            ws_agence.cell(row=current_row, column=1).value = "1. 🏆 CLASSEMENT GÉNÉRAL DES AGENCES"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            
            # Créer le classement avec écart à la moyenne
            df_classement = df_exec.copy()
            df_classement['Écart vs Moyenne'] = df_classement['Taux'] - taux_moyen
            df_classement['Écart vs Moyenne'] = df_classement['Écart vs Moyenne'].round(1)
            df_classement['Rang'] = df_classement['Taux'].rank(ascending=False, method='min').astype(int)
            df_classement = df_classement.sort_values('Rang')
            
            # Ajouter colonne de statut
            def get_status(taux):
                if taux >= 80:
                    return '🟢 Excellent'
                elif taux >= 60:
                    return '🟡 Moyen'
                else:
                    return '🔴 Critique'
            
            df_classement['Statut'] = df_classement['Taux'].apply(get_status)
            df_classement = df_classement[['Rang', 'Agence', 'Total', 'OK', 'KO', 'Taux', 'Écart vs Moyenne', 'Statut']]
            
            for r_idx, row in enumerate(dataframe_to_rows(df_classement, index=False, header=True), current_row):
                for c_idx, value in enumerate(row, 1):
                    cell = ws_agence.cell(row=r_idx, column=c_idx)
                    cell.value = value
                    if r_idx == current_row:  # Header
                        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                        cell.font = Font(bold=True, color="FFFFFF")
                    else:
                        # Mise en forme conditionnelle selon le statut
                        if c_idx == 8 and isinstance(value, str):  # Colonne Statut
                            if '🟢' in value:
                                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                            elif '🔴' in value:
                                cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                            elif '🟡' in value:
                                cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            
            current_row += len(df_classement) + 3
            
            # === SECTION 2: AGENCES À RISQUE ===
            agences_risque = df_classement[df_classement['Taux'] < 60]
            if len(agences_risque) > 0:
                ws_agence.cell(row=current_row, column=1).value = "2. ⚠️ AGENCES À RISQUE (Taux < 60%)"
                ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12, color="C00000")
                current_row += 1
                
                agences_risque['Action recommandée'] = 'Audit urgent + Plan d\'action'
                
                for r_idx, row in enumerate(dataframe_to_rows(agences_risque, index=False, header=True), current_row):
                    for c_idx, value in enumerate(row, 1):
                        cell = ws_agence.cell(row=r_idx, column=c_idx)
                        cell.value = value
                        if r_idx == current_row:
                            cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
                            cell.font = Font(bold=True, color="FFFFFF")
                        else:
                            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                
                current_row += len(agences_risque) + 3
            
            # === SECTION 3: TOP PERFORMERS ===
            top_performers = df_classement.head(5)
            ws_agence.cell(row=current_row, column=1).value = "3. 🌟 TOP 5 PERFORMERS"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12, color="00B050")
            current_row += 1
            
            for r_idx, row in enumerate(dataframe_to_rows(top_performers, index=False, header=True), current_row):
                for c_idx, value in enumerate(row, 1):
                    cell = ws_agence.cell(row=r_idx, column=c_idx)
                    cell.value = value
                    if r_idx == current_row:
                        cell.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
                        cell.font = Font(bold=True, color="FFFFFF")
                    else:
                        cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            
            current_row += len(top_performers) + 3
            
            # === SECTION 4: Volume total par agence (existant) ===
            # === SECTION 4: Volume total par agence (existant) ===
            ws_agence.cell(row=current_row, column=1).value = "4. 📊 VOLUME TOTAL PAR AGENCE"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            
            volume_agence = df['Code_Unite'].value_counts().reset_index()
            volume_agence.columns = ['Agence', 'Nombre total']
            volume_agence['% du total'] = round((volume_agence['Nombre total'] / total * 100), 2)
            
            for r_idx, row in enumerate(dataframe_to_rows(volume_agence, index=False, header=True), current_row):
                for c_idx, value in enumerate(row, 1):
                    ws_agence.cell(row=r_idx, column=c_idx).value = value
            current_row += len(volume_agence) + 3
            
            # === SECTION 5: Taux de réussite par agence ===
            ws_agence.cell(row=current_row, column=1).value = "5. 📈 TAUX DE RÉUSSITE PAR AGENCE (Classement par nombre de KO)"
            ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
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
            
            # === SECTION 6: Top agences avec le plus de rejets ===
            top_ko_agences = agence_status.nlargest(10, 'KO')[['Code_Unite', 'KO', 'Taux réussite (%)']]
            
            ws_agence.cell(row=current_row, column=1).value = "6. 🔴 TOP 10 AGENCES AVEC LE PLUS DE REJETS"
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
            
            # === SECTION 7: Agences × Types d'erreurs ===
            df_ko = df[df['Statut_Final'].str.upper() != 'OK']
            if len(df_ko) > 0:
                try:
                    agence_erreur = pd.crosstab(df_ko['Code_Unite'], df_ko['Statut_Final'], margins=True)
                    agence_erreur = agence_erreur.reset_index()
                    
                    ws_agence.cell(row=current_row, column=1).value = "7. 🔀 CROISEMENT AGENCES × TYPES D'ERREURS"
                    ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                    current_row += 1
                    
                    for r_idx, row in enumerate(dataframe_to_rows(agence_erreur, index=False, header=True), current_row):
                        for c_idx, value in enumerate(row, 1):
                            ws_agence.cell(row=r_idx, column=c_idx).value = value
                    current_row += len(agence_erreur) + 3
                except Exception as e:
                    # Si le crosstab échoue, on passe
                    ws_agence.cell(row=current_row, column=1).value = "7. CROISEMENT AGENCES × TYPES D'ERREURS - Données insuffisantes"
                    current_row += 2
            
            # === SECTION 8: Agences × Types de contrats ===
            try:
                agence_type = pd.crosstab(df['Code_Unite'], df['Type (libellé)'], margins=True)
                agence_type = agence_type.reset_index()
                
                ws_agence.cell(row=current_row, column=1).value = "8. 📊 VOLUME D'INTÉGRATIONS PAR AGENCE ET TYPE DE CONTRAT"
                ws_agence.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                current_row += 1
                
                for r_idx, row in enumerate(dataframe_to_rows(agence_type, index=False, header=True), current_row):
                    for c_idx, value in enumerate(row, 1):
                        ws_agence.cell(row=r_idx, column=c_idx).value = value
            except Exception as e:
                # Si le crosstab échoue, on passe
                ws_agence.cell(row=current_row, column=1).value = "8. VOLUME D'INTÉGRATIONS PAR AGENCE ET TYPE DE CONTRAT - Données insuffisantes"
        
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
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "🔍 Recherche intelligente",
            "📋 Données",
            "🏢 Dashboard Agences",
            "📊 Analyses détaillées", 
            "📈 Visualisations", 
            "💾 Export multi-onglets"
        ])
        
        # TAB 0: Recherche intelligente de contrats
        with tab1:
            st.subheader("🔍 Recherche Hybride Intelligente")
            st.markdown("Recherchez en langage naturel ou par correspondance floue")
            
            # Initialiser l'historique de recherche dans session_state
            if 'search_history' not in st.session_state:
                st.session_state['search_history'] = []
            
            # Zone de recherche principale
            col1, col2 = st.columns([4, 1])
            
            with col1:
                search_query = st.text_input(
                    "🔎 Recherche intelligente",
                    placeholder="Ex: contrats ko nvm septembre, ou 001-NVM-173, ou agence 169 initial...",
                    help="Tapez en langage naturel ou une partie d'un numéro de contrat"
                )
            
            with col2:
                search_mode = st.selectbox(
                    "Mode",
                    ["🧠 Hybride", "🎯 Exact", "🔤 Flou"],
                    help="Hybride: Combine tous les modes | Exact: Correspondance exacte | Flou: Tolère les fautes"
                )
            
            # Afficher des suggestions en temps réel
            if search_query and len(search_query) >= 2:
                suggestions = get_smart_suggestions(search_query, df_clean, limit=5)
                if suggestions:
                    with st.expander("💡 Suggestions", expanded=True):
                        cols = st.columns(len(suggestions))
                        for idx, sugg in enumerate(suggestions):
                            with cols[idx]:
                                if st.button(f"{sugg['type']}: {sugg['value']}", key=f"sugg_{idx}"):
                                    search_query = sugg['value']
                                st.caption(f"Score: {sugg['score']}%")
            
            # Bouton de recherche
            if st.button("🔍 RECHERCHER", type="primary", use_container_width=True) or search_query:
                
                if search_query:
                    # Ajouter à l'historique
                    if search_query not in st.session_state['search_history']:
                        st.session_state['search_history'].insert(0, search_query)
                        st.session_state['search_history'] = st.session_state['search_history'][:10]  # Garder les 10 dernières
                    
                    with st.spinner("🔎 Recherche en cours..."):
                        results = df_clean.copy()
                        
                        if search_mode == "🧠 Hybride":
                            # 1. Parser le langage naturel
                            filters = parse_natural_language_query(search_query, df_clean)
                            
                            # Afficher les filtres détectés
                            if filters:
                                st.info(f"🧠 Filtres détectés: {', '.join([f'{k}: {v}' for k, v in filters.items()])}")
                            
                            # 2. Appliquer les filtres
                            if filters.get('statut'):
                                if filters['statut'] == 'KO':
                                    results = results[results['Statut_Final'].str.upper() != 'OK']
                                else:
                                    results = results[results['Statut_Final'].str.upper() == 'OK']
                            
                            if filters.get('agence'):
                                results = results[results['Code_Unite'] == filters['agence']]
                            
                            if filters.get('type'):
                                results = results[results['Type (libellé)'] == filters['type']]
                            
                            if filters.get('init_avenant'):
                                results = results[results['Initial/Avenant'].str.contains(filters['init_avenant'], case=False, na=False)]
                            
                            if filters.get('mois') and 'Date_Integration' in results.columns:
                                results['Date_Integration'] = pd.to_datetime(results['Date_Integration'], errors='coerce')
                                results = results[results['Date_Integration'].dt.month == filters['mois']]
                            
                            # 3. Calculer les scores de pertinence
                            results['_score'] = results.apply(
                                lambda row: calculate_relevance_score(row, search_query, filters),
                                axis=1
                            )
                            
                            # 4. Filtrer les résultats avec score > 0 et trier
                            results = results[results['_score'] > 0].sort_values('_score', ascending=False)
                        
                        elif search_mode == "🎯 Exact":
                            # Recherche exacte dans toutes les colonnes
                            mask = pd.Series([False] * len(results))
                            for col in results.columns:
                                mask = mask | results[col].astype(str).str.contains(search_query, case=False, na=False)
                            results = results[mask]
                            results['_score'] = 100
                        
                        elif search_mode == "🔤 Flou":
                            # Recherche floue sur le champ Contrat
                            if 'Contrat' in results.columns:
                                fuzzy_matches = fuzzy_search(search_query, results, 'Contrat', limit=50)
                                if fuzzy_matches:
                                    matched_values = [match[0] for match in fuzzy_matches]
                                    results = results[results['Contrat'].isin(matched_values)]
                                    
                                    # Ajouter les scores
                                    score_dict = {match[0]: match[1] for match in fuzzy_matches}
                                    results['_score'] = results['Contrat'].map(score_dict)
                                    results = results.sort_values('_score', ascending=False)
                                else:
                                    results = pd.DataFrame()
                            else:
                                st.warning("Colonne 'Contrat' non trouvée pour la recherche floue")
                                results = pd.DataFrame()
                        
                        # Afficher les résultats
                        if len(results) > 0:
                            st.success(f"✅ {len(results)} résultat(s) trouvé(s)")
                            
                            # Afficher les scores si disponibles
                            if '_score' in results.columns:
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("Score moyen", f"{results['_score'].mean():.1f}%")
                                with col2:
                                    st.metric("Meilleur score", f"{results['_score'].max():.1f}%")
                                with col3:
                                    st.metric("Score minimum", f"{results['_score'].min():.1f}%")
                            
                            # Afficher les résultats avec scores
                            display_results = results.copy()
                            if '_score' in display_results.columns:
                                # Déplacer la colonne score au début
                                cols = ['_score'] + [col for col in display_results.columns if col != '_score']
                                display_results = display_results[cols]
                                display_results = display_results.rename(columns={'_score': '🎯 Score'})
                            
                            st.dataframe(display_results, width='stretch', height=400)
                            
                            # Boutons d'export
                            col1, col2 = st.columns(2)
                            with col1:
                                csv = results.to_csv(index=False).encode('utf-8')
                                st.download_button(
                                    label="📥 Télécharger CSV",
                                    data=csv,
                                    file_name=f"recherche_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                    mime="text/csv",
                                    use_container_width=True
                                )
                            
                            with col2:
                                # Exporter en Excel
                                output = io.BytesIO()
                                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                    results.to_excel(writer, index=False, sheet_name='Résultats')
                                output.seek(0)
                                
                                st.download_button(
                                    label="📥 Télécharger Excel",
                                    data=output,
                                    file_name=f"recherche_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                            
                            # Statistiques sur les résultats
                            if len(results) > 1:
                                st.markdown("### 📊 Statistiques sur les résultats")
                                col1, col2, col3, col4 = st.columns(4)
                                
                                with col1:
                                    if 'Statut_Final' in results.columns:
                                        ok_res = len(results[results['Statut_Final'].str.upper() == 'OK'])
                                        st.metric("✅ OK", ok_res, delta=f"{round(ok_res/len(results)*100, 1)}%")
                                
                                with col2:
                                    if 'Statut_Final' in results.columns:
                                        ko_res = len(results[results['Statut_Final'].str.upper() != 'OK'])
                                        st.metric("❌ KO", ko_res, delta=f"{round(ko_res/len(results)*100, 1)}%")
                                
                                with col3:
                                    if 'Code_Unite' in results.columns:
                                        st.metric("🏢 Agences", results['Code_Unite'].nunique())
                                
                                with col4:
                                    if 'Type (libellé)' in results.columns:
                                        st.metric("📋 Types", results['Type (libellé)'].nunique())
                        
                        else:
                            st.warning(f"❌ Aucun résultat trouvé pour '{search_query}'")
                            st.info("💡 Essayez : \n- Des mots-clés différents\n- Le mode 'Flou' pour plus de tolérance\n- Une recherche plus générale")
            
            # Historique de recherche
            if st.session_state['search_history']:
                with st.expander("📜 Historique des recherches"):
                    st.markdown("Cliquez pour relancer une recherche précédente")
                    cols = st.columns(min(len(st.session_state['search_history']), 5))
                    for idx, hist_query in enumerate(st.session_state['search_history'][:5]):
                        with cols[idx]:
                            if st.button(f"🔄 {hist_query}", key=f"hist_{idx}"):
                                search_query = hist_query
            
            # Aide et exemples
            with st.expander("❓ Aide et exemples de recherche"):
                st.markdown("""
                ### 🧠 Mode Hybride (Recommandé)
                Combine recherche floue + langage naturel
                
                **Exemples de requêtes :**
                - `contrats ko nvm septembre` → Trouve les contrats KO de l'agence NVM en septembre
                - `agence 169 initial août` → Contrats initiaux de l'agence 169 en août
                - `avenant ok octobre` → Avenants validés en octobre
                - `001-NVM-173` → Trouve le contrat même avec fautes de frappe
                - `erreur nvm` → Tous les contrats KO de NVM
                
                ### 🎯 Mode Exact
                Recherche une correspondance exacte dans toutes les colonnes
                
                ### 🔤 Mode Flou
                Tolère les fautes de frappe et trouve des correspondances approximatives
                
                **Astuces :**
                - Utilisez des mots-clés simples
                - Combinez agence + statut + mois pour affiner
                - Le score indique la pertinence (100% = parfait)
                """)
            
            # Recherche avancée (ancien système en fallback)
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
        
                with col2:
                    st.metric(
                        "🔴 Pire",
                        df_agences.iloc[-1]['Agence'],
                        f"{df_agences.iloc[-1]['Taux (%)']}%"
                    )
                
                with col3:
                    st.metric(
                        "📊 Moyenne",
                        f"{taux_moyen:.1f}%"
                    )
                
                with col4:
                    agences_alerte = len(df_agences[df_agences['Taux (%)'] < 60])
                    st.metric(
                        "⚠️ En alerte",
                        agences_alerte,
                        delta=f"< 60%",
                        delta_color="inverse"
                    )
                
                with col5:
                    agences_ok = len(df_agences[df_agences['Taux (%)'] >= taux_moyen])
                    st.metric(
                        "✅ Au-dessus",
                        f"{agences_ok}/{len(df_agences)}"
                    )
                
                # === FILTRE INTERACTIF ===
                st.markdown("### 🔍 Filtre Interactif")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    filtre_agences = st.multiselect(
                        "Sélectionner des agences",
                        options=df_agences['Agence'].tolist(),
                        default=df_agences['Agence'].tolist()[:5]
                    )
                
                with col2:
                    filtre_seuil = st.slider(
                        "Taux minimum (%)",
                        0, 100, 0
                    )
                
                with col3:
                    tri_par = st.selectbox(
                        "Trier par",
                        ["Taux (%)", "KO", "Total", "Agence"]
                    )
                
                # Appliquer les filtres
                df_filtered = df_agences.copy()
                if filtre_agences:
                    df_filtered = df_filtered[df_filtered['Agence'].isin(filtre_agences)]
                df_filtered = df_filtered[df_filtered['Taux (%)'] >= filtre_seuil]
                df_filtered = df_filtered.sort_values(tri_par, ascending=False)
                
                # === GRAPHIQUES ===
                st.markdown("### 📊 Visualisations")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Graphique en barres horizontal
                    fig_bar = go.Figure()
                    
                    colors = ['#28a745' if x >= 80 else '#ffc107' if x >= 60 else '#dc3545' 
                             for x in df_filtered['Taux (%)']]
                    
                    fig_bar.add_trace(go.Bar(
                        y=df_filtered['Agence'],
                        x=df_filtered['Taux (%)'],
                        orientation='h',
                        marker=dict(color=colors),
                        text=df_filtered['Taux (%)'].apply(lambda x: f"{x:.1f}%"),
                        textposition='outside',
                        hovertemplate='<b>%{y}</b><br>Taux: %{x:.1f}%<extra></extra>'
                    ))
                    
                    fig_bar.update_layout(
                        title="Taux de réussite par agence",
                        xaxis_title="Taux de réussite (%)",
                        yaxis_title="Agence",
                        height=400,
                        showlegend=False
                    )
                    
                    st.plotly_chart(fig_bar, use_container_width=True)
                
                with col2:
                    # Scatter plot OK vs KO
                    fig_scatter = px.scatter(
                        df_filtered,
                        x='KO',
                        y='OK',
                        size='Total',
                        color='Taux (%)',
                        hover_name='Agence',
                        title="Répartition OK vs KO par agence",
                        labels={'KO': 'Nombre de KO', 'OK': 'Nombre de OK'},
                        color_continuous_scale='RdYlGn'
                    )
                    
                    fig_scatter.update_layout(height=400)
                    st.plotly_chart(fig_scatter, use_container_width=True)
                
                # === TABLEAU DÉTAILLÉ ===
                st.markdown("### 📋 Tableau Détaillé")
                
                # Ajouter une colonne de statut visuel
                def get_status_emoji(taux):
                    if taux >= 80:
                        return "🟢 Excellent"
                    elif taux >= 60:
                        return "🟡 Moyen"
                    else:
                        return "🔴 Critique"
                
                df_display = df_filtered.copy()
                df_display['Statut'] = df_display['Taux (%)'].apply(get_status_emoji)
                df_display['Écart vs Moyenne'] = df_display['Écart vs Moyenne'].apply(lambda x: f"{x:+.1f}%")
                
                # Réorganiser les colonnes
                df_display = df_display[['Agence', 'Total', 'OK', 'KO', 'Taux (%)', 'Écart vs Moyenne', 'Statut']]
                
                st.dataframe(
                    df_display,
                    width='stretch',
                    height=400,
                    hide_index=True
                )
                
                # === AGENCES À RISQUE ===
                agences_risque = df_agences[df_agences['Taux (%)'] < 60]
                if len(agences_risque) > 0:
                    st.markdown("### ⚠️ Agences à Risque (Taux < 60%)")
                    st.error(f"**{len(agences_risque)} agence(s)** nécessite(nt) une attention immédiate")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.dataframe(
                            agences_risque[['Agence', 'Taux (%)', 'KO']],
                            hide_index=True
                        )
                    
                    with col2:
                        st.markdown("""
                        **Actions recommandées :**
                        - 🔍 Audit approfondi des processus
                        - 📋 Plan d'action correctif
                        - 👥 Formation des équipes
                        - 📊 Suivi hebdomadaire renforcé
                        """)
                
                # === TOP PERFORMERS ===
                st.markdown("### 🌟 Top 5 Performers")
                top_5 = df_agences.head(5)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.dataframe(
                        top_5[['Agence', 'Taux (%)', 'Total']],
                        hide_index=True
                    )
                
                with col2:
                    st.markdown("""
                    **Bonnes pratiques à partager :**
                    - ✅ Processus efficaces
                    - 📚 Documentation de référence
                    - 🎓 Sessions de formation
                    - 🏆 Benchmark pour autres agences
                    """)
                
                # === COMPARAISON TEMPORELLE ===
                if 'Date_Integration' in df_clean.columns:
                    st.markdown("### 📈 Évolution Temporelle")
                    
                    agence_selectionnee = st.selectbox(
                        "Sélectionner une agence pour voir son évolution",
                        options=df_agences['Agence'].tolist()
                    )
                    
                    if agence_selectionnee:
                        df_agence_temp = df_clean[df_clean['Code_Unite'] == agence_selectionnee].copy()
                        df_agence_temp['Date_Integration'] = pd.to_datetime(df_agence_temp['Date_Integration'], errors='coerce')
                        df_agence_temp = df_agence_temp.dropna(subset=['Date_Integration'])
                        
                        if len(df_agence_temp) > 0:
                            df_agence_temp['Mois'] = df_agence_temp['Date_Integration'].dt.to_period('M').astype(str)
                            
                            # Calculer le taux par mois
                            monthly_stats = []
                            for mois in df_agence_temp['Mois'].unique():
                                df_mois = df_agence_temp[df_agence_temp['Mois'] == mois]
                                total_m = len(df_mois)
                                ok_m = (df_mois['Statut_Final'].str.upper() == 'OK').sum()
                                taux_m = (ok_m / total_m * 100) if total_m > 0 else 0
                                monthly_stats.append({
                                    'Mois': mois,
                                    'Total': total_m,
                                    'Taux (%)': taux_m
                                })
                            
                            df_monthly = pd.DataFrame(monthly_stats).sort_values('Mois')
                            
                            fig_evolution = go.Figure()
                            
                            fig_evolution.add_trace(go.Scatter(
                                x=df_monthly['Mois'],
                                y=df_monthly['Taux (%)'],
                                mode='lines+markers',
                                name='Taux de réussite',
                                line=dict(color='#4472C4', width=3),
                                marker=dict(size=10)
                            ))
                            
                            # Ajouter la ligne de moyenne
                            fig_evolution.add_hline(
                                y=taux_moyen,
                                line_dash="dash",
                                line_color="red",
                                annotation_text=f"Moyenne nationale: {taux_moyen:.1f}%"
                            )
                            
                            fig_evolution.update_layout(
                                title=f"Évolution du taux de réussite - {agence_selectionnee}",
                                xaxis_title="Mois",
                                yaxis_title="Taux de réussite (%)",
                                height=400,
                                hovermode='x unified'
                            )
                            
                            st.plotly_chart(fig_evolution, use_container_width=True)
                            
                            # Tendance
                            if len(df_monthly) >= 2:
                                tendance = df_monthly.iloc[-1]['Taux (%)'] - df_monthly.iloc[-2]['Taux (%)']
                                if tendance > 0:
                                    st.success(f"📈 Tendance positive : +{tendance:.1f}% par rapport au mois précédent")
                                elif tendance < 0:
                                    st.error(f"📉 Tendance négative : {tendance:.1f}% par rapport au mois précédent")
                                else:
                                    st.info("→ Stable par rapport au mois précédent")
                
                # === EXPORT DU DASHBOARD ===
                st.markdown("### 💾 Export")
                
                csv_agences = df_display.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Télécharger le tableau (CSV)",
                    data=csv_agences,
                    file_name=f"dashboard_agences_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            else:
                st.warning("⚠️ Colonnes 'Code_Unite' ou 'Statut_Final' manquantes pour l'analyse par agence")
        
        # TAB 3: Analyses détaillées (ancien tab3 devient tab4)
        with tab4:
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
        
        # TAB 4: Visualisations (ancien tab4 devient tab5)
        with tab5:
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
        
        # TAB 5: Export multi-onglets (ancien tab5 devient tab6)
        with tab6:
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
