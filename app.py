# Excel Analyzer Pro - Version Complète
# Application Streamlit pour l'analyse intelligente de fichiers Excel

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from thefuzz import fuzz, process
import re

# ==================== CONFIGURATION ====================
st.set_page_config(
    page_title="Excel Analyzer Pro",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Analyzer Pro - Analyse intelligente de contrats")
st.markdown("### Embellissez, analysez et recherchez dans vos fichiers Excel")

# ==================== FONCTIONS UTILITAIRES ====================

def clean_data(df):
    """Nettoie les données du DataFrame"""
    df = df.dropna(how='all').dropna(axis=1, how='all')
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()
    df = df.replace('nan', '').fillna('')
    return df

def parse_natural_language_query(query, df):
    """Parse une requête en langage naturel"""
    filters = {}
    query_lower = query.lower()
    
    # Statuts
    if any(word in query_lower for word in ['ko', 'échec', 'erreur', 'rejet']):
        filters['statut'] = 'KO'
    elif any(word in query_lower for word in ['ok', 'réussi', 'succès']):
        filters['statut'] = 'OK'
    
    # Agences
    if 'Code_Unite' in df.columns:
        for agence in df['Code_Unite'].unique():
            if str(agence).lower() in query_lower:
                filters['agence'] = agence
                break
    
    # Mois
    mois_map = {'janvier':1,'jan':1,'février':2,'fev':2,'mars':3,'avril':4,'mai':5,'juin':6,
                'juillet':7,'août':8,'aout':8,'septembre':9,'sept':9,'octobre':10,'novembre':11,'décembre':12,'dec':12}
    for nom, num in mois_map.items():
        if nom in query_lower:
            filters['mois'] = num
            break
    
    # Initial/Avenant
    if 'initial' in query_lower:
        filters['init_avenant'] = 'Initial'
    elif 'avenant' in query_lower:
        filters['init_avenant'] = 'Avenant'
    
    return filters

def fuzzy_search(query, df, column, limit=10):
    """Recherche floue"""
    if column not in df.columns:
        return []
    values = [v for v in df[column].dropna().astype(str).unique() if v.strip()]
    if not values or not query.strip():
        return []
    matches = process.extract(query, values, limit=limit, scorer=fuzz.token_sort_ratio)
    return [(m[0], m[1]) for m in matches if m[1] > 50]

def calculate_relevance_score(row, query, filters):
    """Calcule le score de pertinence"""
    score = 0
    query_lower = query.lower()
    
    if 'Contrat' in row.index:
        if query_lower in str(row['Contrat']).lower():
            score += 100
        else:
            score += fuzz.partial_ratio(query_lower, str(row['Contrat']).lower()) * 0.5
    
    if filters.get('agence') and row.get('Code_Unite') == filters['agence']:
        score += 50
    if filters.get('statut'):
        if filters['statut'] == 'KO' and str(row.get('Statut_Final','')).upper() != 'OK':
            score += 50
        elif filters['statut'] == 'OK' and str(row.get('Statut_Final','')).upper() == 'OK':
            score += 50
    if filters.get('mois') and 'Date_Integration' in row.index:
        try:
            if pd.to_datetime(row['Date_Integration']).month == filters['mois']:
                score += 40
        except:
            pass
    
    return score

def get_smart_suggestions(partial_input, df, limit=5):
    """Génère des suggestions intelligentes"""
    suggestions = []
    if not partial_input or len(partial_input) < 2:
        return suggestions
    
    partial_lower = partial_input.lower()
    
    # Contrats
    if 'Contrat' in df.columns:
        matches = df['Contrat'].dropna().astype(str)
        matches = matches[matches.str.contains(partial_input, case=False, na=False)].head(limit)
        suggestions.extend([{'type':'📄 Contrat','value':c,'score':fuzz.partial_ratio(partial_lower,c.lower())} for c in matches])
    
    # Agences
    if 'Code_Unite' in df.columns:
        for agence in df['Code_Unite'].unique():
            if partial_lower in str(agence).lower():
                suggestions.append({'type':'🏢 Agence','value':agence,'score':fuzz.ratio(partial_lower,str(agence).lower())})
    
    # Statuts
    if 'ko' in partial_lower:
        suggestions.append({'type':'❌ Statut','value':'KO','score':100})
    if 'ok' in partial_lower:
        suggestions.append({'type':'✅ Statut','value':'OK','score':100})
    
    return sorted(suggestions, key=lambda x: x['score'], reverse=True)[:limit]

def style_worksheet(ws):
    """Style une feuille Excel"""
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
        fill = PatternFill(start_color="F2F2F2" if row_idx%2==0 else "FFFFFF", fill_type="solid")
        for cell in row:
            cell.fill = fill
            cell.border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)
    
    ws.freeze_panes = 'A2'

def create_comprehensive_excel(df):
    """Crée l'Excel complet avec tous les onglets"""
    from openpyxl import Workbook
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    # Onglet 1: Données
    ws_data = wb.create_sheet('Données nettoyées')
    for r in dataframe_to_rows(df, index=False, header=True):
        ws_data.append(r)
    style_worksheet(ws_data)
    
    # Onglet 2: Vue d'ensemble
    total = len(df)
    ok = len(df[df['Statut_Final'].str.upper() == 'OK'])
    ko = total - ok
    
    ws_vue = wb.create_sheet('Vue d ensemble')
    vue_data = [
        ['Métrique', 'Valeur'],
        ['Total contrats', total],
        ['Contrats OK', ok],
        ['Contrats KO', ko],
        ['Taux réussite', f"{round(ok/total*100,1) if total else 0}%"],
        ['Agences', df['Code_Unite'].nunique() if 'Code_Unite' in df.columns else 0]
    ]
    for row in vue_data:
        ws_vue.append(row)
    style_worksheet(ws_vue)
    
    # Onglet 3: Analyse Agences ENRICHIE
    if 'Code_Unite' in df.columns:
        ws_ag = wb.create_sheet('Analyse par agence')
        ws_ag.append(['ANALYSE COMPLÈTE PAR AGENCE'])
        ws_ag['A1'].font = Font(bold=True, size=14)
        ws_ag.append([])
        
        # Métriques par agence
        agences = []
        for ag in df['Code_Unite'].unique():
            df_ag = df[df['Code_Unite']==ag]
            t = len(df_ag)
            o = (df_ag['Statut_Final'].str.upper()=='OK').sum()
            agences.append({'Agence':ag,'Total':t,'OK':o,'KO':t-o,'Taux':round(o/t*100,1) if t else 0})
        
        df_ag = pd.DataFrame(agences)
        if len(df_ag) > 0:
            moy = df_ag['Taux'].mean()
            
            # Dashboard
            ws_ag.append(['🎯 DASHBOARD EXÉCUTIF'])
            ws_ag.append(['Meilleure agence', f"{df_ag.loc[df_ag['Taux'].idxmax(),'Agence']} ({df_ag['Taux'].max()}%)"])
            ws_ag.append(['Pire agence', f"{df_ag.loc[df_ag['Taux'].idxmin(),'Agence']} ({df_ag['Taux'].min()}%)"])
            ws_ag.append(['Taux moyen', f"{moy:.1f}%"])
            ws_ag.append(['Agences < 60%', len(df_ag[df_ag['Taux']<60])])
            ws_ag.append([])
            
            # Classement
            ws_ag.append(['CLASSEMENT GÉNÉRAL'])
            df_ag['Rang'] = df_ag['Taux'].rank(ascending=False, method='min').astype(int)
            df_ag['Écart'] = (df_ag['Taux'] - moy).round(1)
            df_ag['Statut'] = df_ag['Taux'].apply(lambda x: '🟢 Excellent' if x>=80 else '🟡 Moyen' if x>=60 else '🔴 Critique')
            df_ag = df_ag.sort_values('Rang')
            
            for r in dataframe_to_rows(df_ag[['Rang','Agence','Total','OK','KO','Taux','Écart','Statut']], index=False, header=True):
                ws_ag.append(r)
    
    wb.save(output)
    output.seek(0)
    return output

# ==================== INTERFACE STREAMLIT ====================

uploaded_file = st.file_uploader("📁 Fichier Excel", type=['xlsx','xls'])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        with st.expander("👁️ Aperçu", expanded=False):
            st.dataframe(df.head(10), width='stretch')
        
        df_clean = clean_data(df)
        st.success(f"✅ {len(df_clean)} lignes, {len(df_clean.columns)} colonnes")
        
        tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs(["🔍 Recherche","📋 Données","🏢 Agences","📊 Analyses","📈 Visualisations","💾 Export"])
        
        # TAB 1: RECHERCHE INTELLIGENTE
        with tab1:
            st.subheader("🔍 Recherche Hybride Intelligente")
            
            if 'history' not in st.session_state:
                st.session_state.history = []
            
            col1,col2 = st.columns([4,1])
            with col1:
                query = st.text_input("🔎 Recherche", placeholder="Ex: contrats ko nvm septembre")
            with col2:
                mode = st.selectbox("Mode", ["🧠 Hybride","🎯 Exact","🔤 Flou"])
            
            if query and len(query) >= 2:
                sugg = get_smart_suggestions(query, df_clean, 5)
                if sugg:
                    with st.expander("💡 Suggestions", expanded=True):
                        cols = st.columns(min(len(sugg),5))
                        for i,s in enumerate(sugg[:5]):
                            with cols[i]:
                                st.button(f"{s['type']}: {s['value']}", key=f"s{i}")
            
            if st.button("🔍 RECHERCHER", type="primary") and query:
                with st.spinner("Recherche..."):
                    results = df_clean.copy()
                    
                    if mode == "🧠 Hybride":
                        filters = parse_natural_language_query(query, df_clean)
                        if filters:
                            st.info(f"Filtres: {', '.join([f'{k}:{v}' for k,v in filters.items()])}")
                        
                        if filters.get('statut'):
                            if filters['statut'] == 'KO':
                                results = results[results['Statut_Final'].str.upper() != 'OK']
                            else:
                                results = results[results['Statut_Final'].str.upper() == 'OK']
                        if filters.get('agence'):
                            results = results[results['Code_Unite'] == filters['agence']]
                        if filters.get('mois'):
                            results['Date_Integration'] = pd.to_datetime(results['Date_Integration'], errors='coerce')
                            results = results[results['Date_Integration'].dt.month == filters['mois']]
                        
                        results['_score'] = results.apply(lambda r: calculate_relevance_score(r, query, filters), axis=1)
                        results = results[results['_score'] > 0].sort_values('_score', ascending=False)
                    
                    elif mode == "🎯 Exact":
                        mask = pd.Series([False] * len(results))
                        for col in results.columns:
                            mask |= results[col].astype(str).str.contains(query, case=False, na=False)
                        results = results[mask]
                    
                    else:  # Flou
                        if 'Contrat' in results.columns:
                            matches = fuzzy_search(query, results, 'Contrat', 50)
                            if matches:
                                results = results[results['Contrat'].isin([m[0] for m in matches])]
                                results['_score'] = results['Contrat'].map({m[0]:m[1] for m in matches})
                                results = results.sort_values('_score', ascending=False)
                    
                    if len(results) > 0:
                        st.success(f"✅ {len(results)} résultat(s)")
                        
                        if '_score' in results.columns:
                            col1,col2,col3 = st.columns(3)
                            col1.metric("Score moyen", f"{results['_score'].mean():.0f}%")
                            col2.metric("Meilleur", f"{results['_score'].max():.0f}%")
                            col3.metric("Minimum", f"{results['_score'].min():.0f}%")
                        
                        st.dataframe(results, width='stretch', height=400)
                        
                        csv = results.to_csv(index=False).encode()
                        st.download_button("📥 CSV", csv, f"recherche_{datetime.now():%Y%m%d_%H%M%S}.csv")
                    else:
                        st.warning(f"Aucun résultat pour '{query}'")
        
        # TAB 2: DONNÉES
        with tab2:
            st.subheader("📋 Données nettoyées")
            st.dataframe(df_clean, width='stretch', height=400)
            
            col1,col2,col3,col4 = st.columns(4)
            col1.metric("Lignes", len(df_clean))
            col2.metric("Colonnes", len(df_clean.columns))
            col3.metric("Doublons", df_clean.duplicated().sum())
            if 'Statut_Final' in df_clean.columns:
                col4.metric("OK", len(df_clean[df_clean['Statut_Final'].str.upper()=='OK']))
        
        # TAB 3: DASHBOARD AGENCES
        with tab3:
            st.subheader("🏢 Dashboard Agences")
            
            if 'Code_Unite' in df_clean.columns and 'Statut_Final' in df_clean.columns:
                ag_metrics = []
                for ag in df_clean['Code_Unite'].unique():
                    d = df_clean[df_clean['Code_Unite']==ag]
                    t = len(d)
                    o = (d['Statut_Final'].str.upper()=='OK').sum()
                    ag_metrics.append({'Agence':ag,'Total':t,'OK':o,'KO':t-o,'Taux':round(o/t*100,1) if t else 0})
                
                df_ag = pd.DataFrame(ag_metrics)
                moy = df_ag['Taux'].mean()
                df_ag['Écart'] = df_ag['Taux'] - moy
                df_ag = df_ag.sort_values('Taux', ascending=False)
                
                # Métriques
                st.markdown("### 🎯 Métriques Clés")
                col1,col2,col3,col4,col5 = st.columns(5)
                col1.metric("🏆 Meilleure", df_ag.iloc[0]['Agence'], f"{df_ag.iloc[0]['Taux']}%")
                col2.metric("🔴 Pire", df_ag.iloc[-1]['Agence'], f"{df_ag.iloc[-1]['Taux']}%")
                col3.metric("📊 Moyenne", f"{moy:.1f}%")
                col4.metric("⚠️ < 60%", len(df_ag[df_ag['Taux']<60]))
                col5.metric("✅ > Moy", f"{len(df_ag[df_ag['Taux']>=moy])}/{len(df_ag)}")
                
                # Filtres
                st.markdown("### 🔍 Filtres")
                col1,col2,col3 = st.columns(3)
                with col1:
                    filtre_ag = st.multiselect("Agences", df_ag['Agence'].tolist(), df_ag['Agence'].tolist()[:5])
                with col2:
                    seuil = st.slider("Taux min (%)", 0, 100, 0)
                with col3:
                    tri = st.selectbox("Trier par", ["Taux","KO","Total"])
                
                df_filt = df_ag[df_ag['Agence'].isin(filtre_ag)] if filtre_ag else df_ag
                df_filt = df_filt[df_filt['Taux'] >= seuil].sort_values(tri, ascending=False)
                
                # Graphiques
                st.markdown("### 📊 Visualisations")
                col1,col2 = st.columns(2)
                
                with col1:
                    colors = ['#28a745' if x>=80 else '#ffc107' if x>=60 else '#dc3545' for x in df_filt['Taux']]
                    fig = go.Figure(go.Bar(y=df_filt['Agence'], x=df_filt['Taux'], orientation='h',
                                          marker_color=colors, text=df_filt['Taux'].apply(lambda x:f"{x:.1f}%")))
                    fig.update_layout(title="Taux par agence", height=400, showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    fig = px.scatter(df_filt, x='KO', y='OK', size='Total', color='Taux',
                                    hover_name='Agence', title="OK vs KO", color_continuous_scale='RdYlGn')
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)
                
                # Tableau
                st.markdown("### 📋 Tableau")
                df_filt['Statut'] = df_filt['Taux'].apply(lambda x: '🟢 Excellent' if x>=80 else '🟡 Moyen' if x>=60 else '🔴 Critique')
                st.dataframe(df_filt[['Agence','Total','OK','KO','Taux','Écart','Statut']], width='stretch', hide_index=True)
                
                # Alertes
                risque = df_ag[df_ag['Taux']<60]
                if len(risque) > 0:
                    st.markdown("### ⚠️ Agences à Risque")
                    st.error(f"{len(risque)} agence(s) < 60%")
                    st.dataframe(risque[['Agence','Taux','KO']], hide_index=True)
            else:
                st.warning("Colonnes manquantes")
        
        # TAB 4: ANALYSES
        with tab4:
            st.subheader("📊 Analyses détaillées")
            
            if 'Statut_Final' in df_clean.columns:
                st.markdown("### 🎯 Statuts")
                total = len(df_clean)
                ok = len(df_clean[df_clean['Statut_Final'].str.upper()=='OK'])
                ko = total - ok
                
                col1,col2,col3 = st.columns(3)
                col1.metric("Total", total)
                col2.metric("✅ OK", ok, f"{round(ok/total*100,1)}%")
                col3.metric("❌ KO", ko, f"{round(ko/total*100,1)}%")
                
                if ko > 0:
                    st.markdown("#### Détail erreurs")
                    err = df_clean[df_clean['Statut_Final'].str.upper()!='OK']['Statut_Final'].value_counts().reset_index()
                    err.columns = ['Erreur','Nombre']
                    err['%'] = round(err['Nombre']/ko*100,1)
                    st.dataframe(err, hide_index=True)
        
        # TAB 5: VISUALISATIONS
        with tab5:
            st.subheader("📈 Visualisations")
            
            col1,col2 = st.columns(2)
            
            # Pie OK/KO
            if 'Statut_Final' in df_clean.columns:
                with col1:
                    ok = len(df_clean[df_clean['Statut_Final'].str.upper()=='OK'])
                    ko = len(df_clean) - ok
                    fig = px.pie(values=[ok,ko], names=['OK','KO'], title="Statuts",
                                hole=0.4, color_discrete_map={'OK':'#28a745','KO':'#dc3545'})
                    st.plotly_chart(fig, use_container_width=True)
            
            # Bar Types
            if 'Type (libellé)' in df_clean.columns:
                with col2:
                    types = df_clean['Type (libellé)'].value_counts()
                    fig = px.bar(x=types.index, y=types.values, title="Types", labels={'x':'Type','y':'Nombre'})
                    st.plotly_chart(fig, use_container_width=True)
            
            # Timeline
            if 'Date_Integration' in df_clean.columns:
                st.markdown("#### Évolution")
                df_t = df_clean.copy()
                df_t['Date_Integration'] = pd.to_datetime(df_t['Date_Integration'], errors='coerce')
                df_t = df_t.dropna(subset=['Date_Integration'])
                df_t['Date'] = df_t['Date_Integration'].dt.date
                timeline = df_t.groupby('Date').size().reset_index(name='Nombre')
                
                fig = px.line(timeline, x='Date', y='Nombre', title="Volume quotidien", markers=True)
                st.plotly_chart(fig, use_container_width=True)
        
        # TAB 6: EXPORT
        with tab6:
            st.subheader("💾 Export Excel")
            
            st.markdown("""
            ### 📑 Contenu du fichier Excel :
            1. **Données nettoyées**
            2. **Vue d'ensemble** - Métriques clés
            3. **🆕 Analyse par agence** - Dashboard exécutif + classement avec code couleur
            
            ### ✨ Fonctionnalités :
            - Mise en forme professionnelle
            - Dashboard exécutif avec indicateurs clés
            - Classement des agences avec statuts visuels
            - Codes couleur automatiques (🟢🟡🔴)
            """)
            
            excel = create_comprehensive_excel(df_clean)
            
            st.download_button(
                "⬇️ Télécharger Excel",
                excel,
                f"analyse_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.success("✅ Fichier prêt !")
    
    except Exception as e:
        st.error(f"❌ Erreur: {str(e)}")
        st.exception(e)

else:
    st.info("👆 Uploadez un fichier Excel")
    
    st.markdown("""
    ### 🚀 Fonctionnalités
    
    #### 🔍 Recherche Hybride Intelligente
    - 3 modes : Hybride, Exact, Flou
    - Compréhension du langage naturel
    - Suggestions en temps réel
    - Score de pertinence
    
    #### 🏢 Dashboard Agences
    - Métriques clés en temps réel
    - Filtres interactifs
    - Visualisations avec code couleur
    - Détection agences à risque
    
    #### 📊 Analyses Complètes
    - Statistiques détaillées
    - Visualisations interactives
    - Evolution temporelle
    
    #### 💾 Export Excel Enrichi
    - Dashboard exécutif
    - Classement avec code couleur
    - Mise en forme professionnelle
    """)

st.markdown("---")
st.markdown("<div style='text-align:center;color:#666;'>Excel Analyzer Pro v2.0</div>", unsafe_allow_html=True)
