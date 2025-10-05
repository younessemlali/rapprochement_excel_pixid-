import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# Configuration de la page
st.set_page_config(
    page_title="Excel Analyzer Pro",
    page_icon="📊",
    layout="wide"
)

# Titre principal
st.title("📊 Excel Analyzer Pro")
st.markdown("### Embellissez et analysez vos fichiers Excel en quelques clics")

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
        summary_data = {
            'Métrique': [
                'Nombre total de contrats',
                'Nombre de contrats OK',
                'Nombre de contrats KO',
                'Taux de réussite (%)',
                'Nombre de contrats initiaux',
                'Nombre d\'avenants',
                'Nombre d\'unités distinctes',
                'Période couverte'
            ],
            'Valeur': []
        }
        
        # Calculer les métriques
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
        
        summary_data['Valeur'] = [total, ok_count, ko_count, f"{taux_reussite}%", initiaux, avenants, unites, periode]
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, index=False, sheet_name='Vue d\'ensemble')
        style_worksheet(writer.sheets['Vue d\'ensemble'], df_summary)
        
        # ONGLET 3: Analyse des statuts OK
        if ok_count > 0:
            df_ok = df[df['Statut_Final'].str.upper() == 'OK'].copy()
            
            # Analyser par type
            ok_by_type = df_ok['Type (libellé)'].value_counts().reset_index()
            ok_by_type.columns = ['Type de contrat', 'Nombre']
            ok_by_type['Pourcentage'] = round((ok_by_type['Nombre'] / ok_count * 100), 2)
            
            # Analyser par unité
            ok_by_unite = df_ok['Code_Unite'].value_counts().reset_index()
            ok_by_unite.columns = ['Unité', 'Nombre']
            ok_by_unite['Pourcentage'] = round((ok_by_unite['Nombre'] / ok_count * 100), 2)
            
            # Créer un résumé
            ok_summary = pd.DataFrame({
                'Métrique': ['Total contrats OK', 'Nombre de types différents', 'Nombre d\'unités'],
                'Valeur': [ok_count, df_ok['Type (libellé)'].nunique(), df_ok['Code_Unite'].nunique()]
            })
            
            # Écrire dans l'onglet
            ok_summary.to_excel(writer, index=False, sheet_name='Contrats OK', startrow=0)
            ok_by_type.to_excel(writer, index=False, sheet_name='Contrats OK', startrow=len(ok_summary) + 3)
            ok_by_unite.to_excel(writer, index=False, sheet_name='Contrats OK', startrow=len(ok_summary) + len(ok_by_type) + 6)
            
            # Ajouter des titres
            ws_ok = writer.sheets['Contrats OK']
            ws_ok.insert_rows(len(ok_summary) + 2)
            ws_ok.cell(row=len(ok_summary) + 2, column=1).value = "Répartition par type de contrat:"
            ws_ok.cell(row=len(ok_summary) + 2, column=1).font = Font(bold=True, size=12)
            
            ws_ok.insert_rows(len(ok_summary) + len(ok_by_type) + 5)
            ws_ok.cell(row=len(ok_summary) + len(ok_by_type) + 5, column=1).value = "Répartition par unité:"
            ws_ok.cell(row=len(ok_summary) + len(ok_by_type) + 5, column=1).font = Font(bold=True, size=12)
            
            style_worksheet(ws_ok, ok_summary)
        
        # ONGLET 4: Analyse des statuts KO avec erreurs
        if ko_count > 0:
            df_ko = df[df['Statut_Final'].str.upper() != 'OK'].copy()
            
            # Résumé KO
            ko_summary = pd.DataFrame({
                'Métrique': ['Total contrats KO', 'Taux d\'échec (%)', 'Nombre de types d\'erreurs'],
                'Valeur': [
                    ko_count,
                    f"{round((ko_count / total * 100), 2)}%",
                    df_ko['Statut_Final'].nunique()
                ]
            })
            
            # Analyser les types d'erreurs (Statut_Final)
            ko_by_status = df_ko['Statut_Final'].value_counts().reset_index()
            ko_by_status.columns = ['Type d\'erreur', 'Nombre']
            ko_by_status['Pourcentage'] = round((ko_by_status['Nombre'] / ko_count * 100), 2)
            
            # Analyser les messages d'erreur
            error_messages = []
            if 'Message_Integration' in df_ko.columns:
                msg_int = df_ko[df_ko['Message_Integration'] != '']['Message_Integration'].value_counts()
                if len(msg_int) > 0:
                    error_messages.append(('Message_Integration', msg_int))
            
            if 'Message_Transfert' in df_ko.columns:
                msg_trans = df_ko[df_ko['Message_Transfert'] != '']['Message_Transfert'].value_counts()
                if len(msg_trans) > 0:
                    error_messages.append(('Message_Transfert', msg_trans))
            
            # Analyser KO par type de contrat
            ko_by_type = df_ko['Type (libellé)'].value_counts().reset_index()
            ko_by_type.columns = ['Type de contrat', 'Nombre KO']
            
            # Écrire dans l'onglet
            current_row = 0
            ko_summary.to_excel(writer, index=False, sheet_name='Contrats KO', startrow=current_row)
            current_row += len(ko_summary) + 3
            
            # Types d'erreurs
            ws_ko = writer.sheets['Contrats KO']
            ws_ko.cell(row=current_row, column=1).value = "Répartition des erreurs par statut:"
            ws_ko.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            ko_by_status.to_excel(writer, index=False, sheet_name='Contrats KO', startrow=current_row)
            current_row += len(ko_by_status) + 3
            
            # Messages d'erreur
            for msg_type, msg_counts in error_messages:
                ws_ko.cell(row=current_row, column=1).value = f"Messages d'erreur - {msg_type}:"
                ws_ko.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                current_row += 1
                
                msg_df = pd.DataFrame({
                    'Message': msg_counts.index,
                    'Occurrences': msg_counts.values
                }).head(10)  # Top 10 messages
                msg_df.to_excel(writer, index=False, sheet_name='Contrats KO', startrow=current_row)
                current_row += len(msg_df) + 3
            
            # KO par type
            ws_ko.cell(row=current_row, column=1).value = "Contrats KO par type:"
            ws_ko.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            ko_by_type.to_excel(writer, index=False, sheet_name='Contrats KO', startrow=current_row)
            
            style_worksheet(ws_ko, ko_summary)
        
        # ONGLET 5: Analyse des types d'avenants
        if 'Initial/Avenant' in df.columns and 'Type (libellé)' in df.columns:
            # Répartition Initial vs Avenant
            init_avenant = df['Initial/Avenant'].value_counts().reset_index()
            init_avenant.columns = ['Catégorie', 'Nombre']
            init_avenant['Pourcentage'] = round((init_avenant['Nombre'] / total * 100), 2)
            
            # Détail des types
            types_detail = df['Type (libellé)'].value_counts().reset_index()
            types_detail.columns = ['Type de contrat', 'Nombre']
            types_detail['Pourcentage'] = round((types_detail['Nombre'] / total * 100), 2)
            
            # Croisement type x statut
            if 'Statut_Final' in df.columns:
                cross_type_status = pd.crosstab(
                    df['Type (libellé)'],
                    df['Statut_Final'],
                    margins=True,
                    margins_name='Total'
                ).reset_index()
            
            # Écrire
            current_row = 0
            ws_types = writer.sheets.get('Types et Avenants')
            if ws_types is None:
                init_avenant.to_excel(writer, index=False, sheet_name='Types et Avenants', startrow=current_row)
                ws_types = writer.sheets['Types et Avenants']
                current_row += len(init_avenant) + 3
                
                ws_types.cell(row=current_row, column=1).value = "Détail par type de contrat:"
                ws_types.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                current_row += 1
                types_detail.to_excel(writer, index=False, sheet_name='Types et Avenants', startrow=current_row)
                current_row += len(types_detail) + 3
                
                if 'cross_type_status' in locals():
                    ws_types.cell(row=current_row, column=1).value = "Croisement Type × Statut:"
                    ws_types.cell(row=current_row, column=1).font = Font(bold=True, size=12)
                    current_row += 1
                    cross_type_status.to_excel(writer, index=False, sheet_name='Types et Avenants', startrow=current_row)
                
                style_worksheet(ws_types, init_avenant)
        
        # ONGLET 6: Analyse temporelle
        if 'Date_Integration' in df.columns:
            df_temp = df.copy()
            df_temp['Date_Integration'] = pd.to_datetime(df_temp['Date_Integration'], errors='coerce')
            df_temp = df_temp.dropna(subset=['Date_Integration'])
            
            if len(df_temp) > 0:
                # Par jour
                df_temp['Date'] = df_temp['Date_Integration'].dt.date
                timeline_day = df_temp.groupby('Date').size().reset_index(name='Nombre de contrats')
                
                # Par mois
                df_temp['Mois'] = df_temp['Date_Integration'].dt.to_period('M').astype(str)
                timeline_month = df_temp.groupby('Mois').size().reset_index(name='Nombre de contrats')
                
                # Par statut et date
                timeline_status = df_temp.groupby(['Date', 'Statut_Final']).size().reset_index(name='Nombre')
                timeline_status_pivot = timeline_status.pivot(index='Date', columns='Statut_Final', values='Nombre').fillna(0).reset_index()
                
                # Écrire
                current_row = 0
                ws_time = writer.sheets.get('Analyse temporelle')
                if ws_time is None:
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

# Fonction pour analyser les données
def analyze_data(df):
    """Analyse les données et retourne des statistiques"""
    analysis = {}
    
    analysis['total_rows'] = len(df)
    analysis['total_columns'] = len(df.columns)
    
    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
    analysis['numeric_stats'] = df[numeric_cols].describe() if numeric_cols else None
    
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    analysis['categorical_summary'] = {}
    for col in categorical_cols[:5]:
        value_counts = df[col].value_counts()
        if len(value_counts) > 0:
            analysis['categorical_summary'][col] = value_counts
    
    analysis['missing_values'] = df.isnull().sum()
    analysis['duplicates'] = df.duplicated().sum()
    
    return analysis

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
            st.dataframe(df.head(10), use_container_width=True)
        
        df_clean = clean_data(df)
        
        st.success(f"✅ Fichier chargé avec succès : {len(df_clean)} lignes, {len(df_clean.columns)} colonnes")
        
        # Créer des onglets
        tab1, tab2, tab3, tab4 = st.tabs(["📋 Données", "📊 Analyses détaillées", "📈 Visualisations", "💾 Export multi-onglets"])
        
        # TAB 1: Données nettoyées
        with tab1:
            st.subheader("Données nettoyées et formatées")
            st.dataframe(df_clean, use_container_width=True, height=400)
            
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
        with tab2:
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
                    st.markdown("#### Détail des erreurs")
                    df_ko = df_clean[df_clean['Statut_Final'].str.upper() != 'OK']
                    error_types = df_ko['Statut_Final'].value_counts().reset_index()
                    error_types.columns = ['Type d\'erreur', 'Nombre']
                    st.dataframe(error_types, use_container_width=True, hide_index=True)
            
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
                st.dataframe(types_count, use_container_width=True, hide_index=True)
        
        # TAB 3: Visualisations
        with tab3:
            st.subheader("📈 Visualisations interactives")
            
            date_cols = [col for col in df_clean.columns if 'date' in col.lower()]
            status_cols = [col for col in df_clean.columns if 'statut' in col.lower()]
            type_cols = [col for col in df_clean.columns if 'type' in col.lower()]
            
            col1, col2 = st.columns(2)
            
            # Graphique statuts OK/KO
            if status_cols and 'Statut_Final' in df_clean.columns:
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
                    st.plotly_chart(fig, use_container_width=True)
            
            # Graphique types
            if 'Type (libellé)' in df_clean.columns:
                with col2:
                    st.markdown("#### Types de contrats")
                    type_counts = df_clean['Type (libellé)'].value_counts()
                    fig = px.bar(
                        x=type_counts.index,
                        y=type_counts.values,
                        title="Nombre par type",
                        labels={'x': 'Type', 'y': 'Nombre'}
                    )
                    fig.update_layout(showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
            
            # Évolution temporelle
            if date_cols and 'Date_Integration' in df_clean.columns:
                st.markdown("#### Évolution temporelle")
                df_temp = df_clean.copy()
                df_temp['Date_Integration'] = pd.to_datetime(df_temp['Date_Integration'], errors='coerce')
                df_temp = df_temp.dropna(subset=['Date_Integration'])
                df_temp['Date'] = df_temp['Date_Integration'].dt.date
                timeline = df_temp.groupby('Date').size().reset_index(name='Nombre')
                
                fig = px.line(
                    timeline,
                    x='Date',
                    y='Nombre',
                    title="Volume de contrats par jour",
                    markers=True
                )
                st.plotly_chart(fig, use_container_width=True)
            
            # Analyse croisée Type × Statut
            if 'Type (libellé)' in df_clean.columns and 'Statut_Final' in df_clean.columns:
                st.markdown("#### Analyse croisée Type × Statut")
                cross_data = pd.crosstab(df_clean['Type (libellé)'], df_clean['Statut_Final'])
                fig = px.bar(
                    cross_data,
                    barmode='group',
                    title="Répartition des statuts par type de contrat"
                )
                st.plotly_chart(fig, use_container_width=True)
        
        # TAB 4: Export multi-onglets
        with tab4:
            st.subheader("💾 Télécharger l'analyse complète Excel")
            
            st.markdown("""
            ### 📑 Le fichier Excel généré contient les onglets suivants :
            
            1. **Données nettoyées** - Toutes vos données formatées et nettoyées
            2. **Vue d'ensemble** - Métriques clés et statistiques générales
            3. **Contrats OK** - Analyse détaillée des contrats réussis par type et unité
            4. **Contrats KO** - Analyse des erreurs avec messages détaillés
            5. **Types et Avenants** - Répartition des types de contrats et avenants
            6. **Analyse temporelle** - Évolution dans le temps avec volumes quotidiens et mensuels
            
            ### ✨ Caractéristiques :
            - 🎨 Mise en forme professionnelle sur tous les onglets
            - 📊 Tableaux de synthèse avec pourcentages
            - 🔍 Filtres automatiques activés
            - 📈 Analyses croisées (Type × Statut)
            - ⚠️ Messages d'erreur détaillés pour les contrats KO
            """)
            
            excel_file = create_comprehensive_excel(df_clean)
            
            st.download_button(
                label="⬇️ Télécharger l'analyse Excel complète (multi-onglets)",
                data=excel_file,
                file_name=f"analyse_complete_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.success("✅ Fichier Excel avec 6 onglets d'analyse prêt au téléchargement !")
    
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier : {str(e)}")
        st.info("Vérifiez que votre fichier Excel est valide et non corrompu.")

else:
    st.info("👆 Commencez par uploader un fichier Excel pour l'analyser")
    
    st.markdown("""
    ### 🚀 Fonctionnalités de l'application
    
    #### 1️⃣ Nettoyage automatique
    - Suppression des lignes et colonnes vides
    - Gestion intelligente des valeurs "nan"
    - Suppression des espaces superflus
    
    #### 2️⃣ Analyses multi-niveaux
    - Vue d'ensemble avec métriques clés (taux de réussite, volumes)
    - Analyse détaillée des contrats OK (par type, par unité)
    - Analyse approfondie des contrats KO avec messages d'erreur
    - Répartition Initial vs Avenants
    - Croisements Type × Statut
    - Analyse temporelle (évolution quotidienne et mensuelle)
    
    #### 3️⃣ Visualisations interactives
    - Graphiques de distribution OK/KO
    - Évolution temporelle
    - Analyse croisée Type × Statut
    - Tableaux de bord dynamiques
    
    #### 4️⃣ Export Excel multi-onglets
    - **6 onglets d'analyse** au lieu d'un seul fichier
    - Chaque onglet répond à une question spécifique
    - Mise en forme professionnelle automatique
    - Tableaux de synthèse avec pourcentages
    - Filtres et navigation optimisés
    
    ---
    
    ### 📊 Spécificités pour les contrats
    - Détection automatique des statuts OK/KO
    - Analyse des messages d'erreur (Integration et Transfert)
    - Classification Initial/Avenant
    - Suivi par unité et par type de contrat
    - Analyse temporelle des intégrations
    
    ### 📋 Formats supportés
    - `.xlsx` (Excel 2007 et plus récent)
    - `.xls` (Excel 97-2003)
    
    ### ⚡ Performance
    Optimisé pour traiter des fichiers avec **plusieurs dizaines de milliers de lignes**
    """)

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666;'>Excel Analyzer Pro - Analyse complète de contrats avec rapports multi-onglets</div>",
    unsafe_allow_html=True
)
