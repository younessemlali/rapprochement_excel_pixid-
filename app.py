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
    # Supprimer les lignes entièrement vides
    df = df.dropna(how='all')
    
    # Supprimer les colonnes entièrement vides
    df = df.dropna(axis=1, how='all')
    
    # Supprimer les espaces en début/fin dans les colonnes texte
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()
    
    # Remplacer les valeurs vides par des chaînes vides pour éviter "nan"
    df = df.fillna('')
    
    return df

# Fonction pour créer un fichier Excel embelli
def create_beautiful_excel(df, filename="fichier_embelli.xlsx"):
    """Crée un fichier Excel avec mise en forme professionnelle"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Données')
        
        workbook = writer.book
        worksheet = writer.sheets['Données']
        
        # Styles
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
        
        # Appliquer le style aux en-têtes
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = alignment_center
            cell.border = border_style
        
        # Appliquer le style aux lignes de données
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
        
        # Ajouter des filtres automatiques
        worksheet.auto_filter.ref = worksheet.dimensions
    
    output.seek(0)
    return output

# Fonction pour analyser les données
def analyze_data(df):
    """Analyse les données et retourne des statistiques"""
    analysis = {}
    
    # Statistiques générales
    analysis['total_rows'] = len(df)
    analysis['total_columns'] = len(df.columns)
    
    # Analyser les colonnes numériques
    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
    analysis['numeric_stats'] = df[numeric_cols].describe() if numeric_cols else None
    
    # Analyser les colonnes catégorielles
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    analysis['categorical_summary'] = {}
    for col in categorical_cols[:5]:  # Limiter aux 5 premières colonnes
        value_counts = df[col].value_counts()
        if len(value_counts) > 0:
            analysis['categorical_summary'][col] = value_counts
    
    # Détection de valeurs manquantes
    analysis['missing_values'] = df.isnull().sum()
    
    # Détection de doublons
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
        # Lire le fichier Excel
        df = pd.read_excel(uploaded_file)
        
        # Afficher un aperçu des données brutes
        with st.expander("👁️ Aperçu des données brutes", expanded=False):
            st.dataframe(df.head(10), use_container_width=True)
        
        # Nettoyer les données
        df_clean = clean_data(df)
        
        st.success(f"✅ Fichier chargé avec succès : {len(df_clean)} lignes, {len(df_clean.columns)} colonnes")
        
        # Créer des onglets
        tab1, tab2, tab3, tab4 = st.tabs(["📋 Données nettoyées", "📊 Analyses", "📈 Visualisations", "💾 Export"])
        
        # TAB 1: Données nettoyées
        with tab1:
            st.subheader("Données nettoyées et formatées")
            st.dataframe(df_clean, use_container_width=True, height=400)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Lignes", len(df_clean))
            with col2:
                st.metric("Colonnes", len(df_clean.columns))
            with col3:
                duplicates = df_clean.duplicated().sum()
                st.metric("Doublons détectés", duplicates)
        
        # TAB 2: Analyses
        with tab2:
            st.subheader("📊 Analyses statistiques")
            
            analysis = analyze_data(df_clean)
            
            # Colonnes numériques
            if analysis['numeric_stats'] is not None and not analysis['numeric_stats'].empty:
                st.markdown("#### 🔢 Statistiques des colonnes numériques")
                st.dataframe(analysis['numeric_stats'], use_container_width=True)
            
            # Colonnes catégorielles
            if analysis['categorical_summary']:
                st.markdown("#### 📝 Répartition des valeurs catégorielles")
                cols = st.columns(2)
                for idx, (col_name, value_counts) in enumerate(analysis['categorical_summary'].items()):
                    with cols[idx % 2]:
                        st.markdown(f"**{col_name}**")
                        st.dataframe(
                            pd.DataFrame({
                                'Valeur': value_counts.index,
                                'Nombre': value_counts.values
                            }),
                            hide_index=True
                        )
            
            # Valeurs manquantes
            missing = analysis['missing_values'][analysis['missing_values'] > 0]
            if not missing.empty:
                st.markdown("#### ⚠️ Valeurs manquantes")
                st.dataframe(
                    pd.DataFrame({
                        'Colonne': missing.index,
                        'Nombre': missing.values
                    }),
                    hide_index=True
                )
            else:
                st.success("✅ Aucune valeur manquante détectée")
        
        # TAB 3: Visualisations
        with tab3:
            st.subheader("📈 Visualisations interactives")
            
            # Identifier les colonnes pertinentes
            date_cols = [col for col in df_clean.columns if 'date' in col.lower()]
            status_cols = [col for col in df_clean.columns if 'statut' in col.lower()]
            type_cols = [col for col in df_clean.columns if 'type' in col.lower()]
            
            col1, col2 = st.columns(2)
            
            # Graphique 1: Distribution des statuts
            if status_cols:
                with col1:
                    st.markdown("#### Distribution des statuts")
                    status_col = status_cols[0]
                    status_counts = df_clean[status_col].value_counts()
                    fig = px.pie(
                        values=status_counts.values,
                        names=status_counts.index,
                        title=f"Répartition: {status_col}",
                        hole=0.4
                    )
                    st.plotly_chart(fig, use_container_width=True)
            
            # Graphique 2: Distribution des types
            if type_cols:
                with col2:
                    st.markdown("#### Distribution des types")
                    type_col = type_cols[0]
                    type_counts = df_clean[type_col].value_counts()
                    fig = px.bar(
                        x=type_counts.index,
                        y=type_counts.values,
                        title=f"Nombre par {type_col}",
                        labels={'x': type_col, 'y': 'Nombre'}
                    )
                    fig.update_layout(showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
            
            # Graphique 3: Évolution temporelle
            if date_cols:
                st.markdown("#### Évolution temporelle")
                date_col = date_cols[0]
                df_clean[date_col] = pd.to_datetime(df_clean[date_col], errors='coerce')
                df_temp = df_clean.dropna(subset=[date_col])
                df_temp['Date'] = df_temp[date_col].dt.date
                timeline = df_temp.groupby('Date').size().reset_index(name='Nombre')
                
                fig = px.line(
                    timeline,
                    x='Date',
                    y='Nombre',
                    title=f"Évolution du nombre d'entrées dans le temps",
                    markers=True
                )
                st.plotly_chart(fig, use_container_width=True)
            
            # Graphique 4: Matrice de corrélation pour colonnes numériques
            numeric_cols = df_clean.select_dtypes(include=['int64', 'float64']).columns.tolist()
            if len(numeric_cols) > 1:
                st.markdown("#### Corrélations entre colonnes numériques")
                corr = df_clean[numeric_cols].corr()
                fig = px.imshow(
                    corr,
                    text_auto=True,
                    aspect="auto",
                    title="Matrice de corrélation",
                    color_continuous_scale='RdBu_r'
                )
                st.plotly_chart(fig, use_container_width=True)
        
        # TAB 4: Export
        with tab4:
            st.subheader("💾 Télécharger le fichier embelli")
            
            st.markdown("""
            Le fichier Excel généré inclut :
            - ✨ Mise en forme professionnelle (en-têtes colorés, lignes alternées)
            - 🎨 Bordures et alignements optimisés
            - 📏 Largeurs de colonnes ajustées automatiquement
            - 🔒 Ligne d'en-tête figée
            - 🔍 Filtres automatiques activés
            - 🧹 Données nettoyées (espaces supprimés, doublons identifiés)
            """)
            
            # Créer le fichier Excel embelli
            excel_file = create_beautiful_excel(df_clean)
            
            # Bouton de téléchargement
            st.download_button(
                label="⬇️ Télécharger le fichier Excel embelli",
                data=excel_file,
                file_name=f"fichier_embelli_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.info("💡 Astuce : Le fichier téléchargé conserve toutes les données nettoyées avec une mise en forme professionnelle.")
    
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier : {str(e)}")
        st.info("Vérifiez que votre fichier Excel est valide et non corrompu.")

else:
    # Instructions d'utilisation
    st.info("👆 Commencez par uploader un fichier Excel pour l'analyser et l'embellir")
    
    st.markdown("""
    ### 🚀 Fonctionnalités de l'application
    
    #### 1️⃣ Nettoyage automatique
    - Suppression des lignes et colonnes vides
    - Suppression des espaces superflus
    - Détection des doublons
    
    #### 2️⃣ Analyses avancées
    - Statistiques descriptives (moyenne, médiane, écart-type)
    - Répartition des valeurs catégorielles
    - Détection des valeurs manquantes
    - Identification des anomalies
    
    #### 3️⃣ Visualisations interactives
    - Graphiques de distribution
    - Évolution temporelle
    - Matrices de corrélation
    - Tableaux de bord dynamiques
    
    #### 4️⃣ Export embelli
    - Mise en forme professionnelle
    - En-têtes colorés et en gras
    - Lignes alternées pour meilleure lisibilité
    - Colonnes ajustées automatiquement
    - Filtres et navigation optimisés
    
    ---
    
    ### 📋 Formats supportés
    - `.xlsx` (Excel 2007 et plus récent)
    - `.xls` (Excel 97-2003)
    
    ### ⚡ Performance
    Optimisé pour traiter des fichiers avec **plusieurs dizaines de milliers de lignes**
    """)

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666;'>Excel Analyzer Pro - Créé avec Streamlit</div>",
    unsafe_allow_html=True
)
