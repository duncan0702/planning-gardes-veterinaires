import streamlit as st
from datetime import datetime, timedelta
import pandas as pd
from Planning_gardes import VetScheduler, VetSchedulerHistory, VetSchedulerConfig
import json
import os

# Configuration de la page
st.set_page_config(
    page_title="Planning de Gardes Vétérinaires",
    page_icon="🏥",
    layout="wide"
)

# Titre principal
st.title("🏥 Planning de Gardes Vétérinaires")
st.markdown("---")

# Initialiser l'historique
@st.cache_resource
def load_history():
    """Charge l'historique (mis en cache pour performance)"""
    history_file = 'historique_gardes.json'
    return VetSchedulerHistory(history_file)

history = load_history()

# Sidebar pour la configuration
st.sidebar.header("⚙️ Configuration")

# Section 1 : Dates
st.sidebar.subheader("📅 Période")
col1, col2 = st.sidebar.columns(2)

with col1:
    start_date = st.date_input(
        "Date de début",
        value=datetime(2026, 1, 5),
        help="Choisissez un lundi de préférence"
    )

with col2:
    # Par défaut : 5 semaines après la date de début
    default_end = start_date + timedelta(days=34)
    end_date = st.date_input(
        "Date de fin",
        value=default_end,
        help="Choisissez un dimanche de préférence"
    )

# Section 2 : Vétérinaires (avec possibilité de modifier)
st.sidebar.subheader("👨‍⚕️ Vétérinaires")

# Vétérinaires par défaut
default_vets = {
    'Dr. Julien': {'jour_repos': [1, 3], 'conges': []},
    'Dr. Maxime': {'jour_repos': 2, 'conges': []},
    'Dr. Isaure': {'jour_repos': 4, 'conges': []},
    'Dr. Mélanie': {'jour_repos': 2, 'conges': []},
    'Dr. Nicolas': {'jour_repos': 3, 'conges': []},
    'Dr. Timoty': {'jour_repos': [3, 4], 'conges': []},
    'Dr. Laura': {'jour_repos': 0, 'conges': []},
    'Dr. Lauranne': {'jour_repos': 4, 'conges': []},
    'Dr. Malaurie': {'jour_repos': 1, 'conges': []},
    'Dr. Sarah': {'jour_repos': 2, 'conges': []},
    'Dr. Olivier': {'jour_repos': [2], 'conges': []},
    'Dr. Dorra': {'jour_repos': [1, 2, 3], 'conges': []},
}

# Stocker les vétérinaires dans session_state
if 'veterinaires' not in st.session_state:
    st.session_state.veterinaires = default_vets.copy()

# Option pour ajouter des congés
with st.sidebar.expander("➕ Ajouter des congés"):
    vet_conge = st.selectbox("Vétérinaire", list(st.session_state.veterinaires.keys()))
    date_conge = st.date_input("Date de congé", key="date_conge_input")
    
    if st.button("Ajouter ce congé"):
        conge_str = date_conge.strftime('%Y-%m-%d')
        if conge_str not in st.session_state.veterinaires[vet_conge]['conges']:
            st.session_state.veterinaires[vet_conge]['conges'].append(conge_str)
            st.success(f"✅ Congé ajouté pour {vet_conge} le {conge_str}")
            st.rerun()

# Afficher les congés actuels
with st.sidebar.expander("📋 Congés enregistrés"):
    has_conges = False
    for vet_name, vet_info in st.session_state.veterinaires.items():
        if vet_info['conges']:
            has_conges = True
            st.write(f"**{vet_name}:**")
            for conge in vet_info['conges']:
                col_c1, col_c2 = st.columns([3, 1])
                with col_c1:
                    st.write(f"  • {conge}")
                with col_c2:
                    if st.button("❌", key=f"del_{vet_name}_{conge}"):
                        st.session_state.veterinaires[vet_name]['conges'].remove(conge)
                        st.rerun()
    
    if not has_conges:
        st.info("Aucun congé enregistré")

# Bouton pour réinitialiser les congés
if st.sidebar.button("🔄 Réinitialiser tous les congés"):
    for vet_name in st.session_state.veterinaires:
        st.session_state.veterinaires[vet_name]['conges'] = []
    st.success("✅ Tous les congés ont été effacés")
    st.rerun()

st.sidebar.markdown("---")

# Section 3 : Historique
st.sidebar.subheader("📚 Historique")
if st.sidebar.button("🗑️ Effacer l'historique"):
    if st.sidebar.checkbox("Confirmer la suppression"):
        history.clear()
        st.sidebar.success("✅ Historique effacé")
        st.rerun()

# Zone principale
tab1, tab2, tab3 = st.tabs(["🎯 Générer Planning", "📊 Historique", "ℹ️ Aide"])

# TAB 1 : Génération du planning
with tab1:
    st.header("Générer un nouveau planning")
    
    # Afficher un résumé
    col_info1, col_info2, col_info3 = st.columns(3)
    
    with col_info1:
        st.metric("Date de début", start_date.strftime('%d/%m/%Y'))
    
    with col_info2:
        st.metric("Date de fin", end_date.strftime('%d/%m/%Y'))
    
    with col_info3:
        nb_jours = (end_date - start_date).days + 1
        st.metric("Nombre de jours", nb_jours)
    
    # Vérifier les congés
    total_conges = sum(len(v['conges']) for v in st.session_state.veterinaires.values())
    if total_conges > 0:
        st.info(f"ℹ️ {total_conges} jour(s) de congé enregistré(s)")
    
    st.markdown("---")
    
    # Bouton de génération
    if st.button("🚀 Générer le planning", type="primary", use_container_width=True):
        
        with st.spinner("⏳ Génération du planning en cours... (cela peut prendre 10-30 secondes)"):
            try:
                # Créer le scheduler
                scheduler = VetScheduler(
                    start_date=start_date.strftime('%Y-%m-%d'),
                    end_date=end_date.strftime('%Y-%m-%d'),
                    veterinaires=st.session_state.veterinaires,
                    history=history
                )
                
                # Résoudre
                schedule = scheduler.solve(time_limit=60)
                
                if schedule:
                    st.success("✅ Planning généré avec succès !")
                    
                    # Stocker dans session_state
                    st.session_state.schedule = schedule
                    st.session_state.scheduler = scheduler
                    
                    # Ajouter à l'historique
                    period_name = f"{start_date.strftime('%B_%Y')}".lower()
                    history.add_schedule(schedule, period_name=period_name)
                    
                    st.rerun()
                    
                else:
                    st.error("❌ Impossible de générer un planning avec ces contraintes.")
                    st.warning("💡 Essayez de :")
                    st.write("- Réduire les congés")
                    st.write("- Augmenter la période")
                    st.write("- Vérifier la configuration des vétérinaires")
                    
            except Exception as e:
                st.error(f"❌ Erreur : {str(e)}")
                st.exception(e)
    
    # Afficher le planning si généré
    if 'schedule' in st.session_state and st.session_state.schedule:
        st.markdown("---")
        st.subheader("📅 Planning généré")
        
        # Convertir en DataFrame pour affichage
        schedule = st.session_state.schedule
        
        # Créer un DataFrame formaté
        data_display = []
        for entry in schedule:
            date_obj = datetime.strptime(entry['date'], '%Y-%m-%d')
            jour_fr = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'][date_obj.weekday()]
            is_weekend = date_obj.weekday() >= 5
            
            second_role = entry.get('deuxieme', '') if is_weekend else entry.get('rappelable', '')
            second_label = "2ème de garde" if is_weekend else "Rappelable"
            
            data_display.append({
                'Date': entry['date'],
                'Jour': jour_fr,
                'Premier de garde': entry.get('premier', ''),
                second_label: second_role
            })
        
        df = pd.DataFrame(data_display)
        
        # Afficher avec style
        st.dataframe(df, use_container_width=True, hide_index=True)
        
        # Boutons d'export
        col_exp1, col_exp2 = st.columns(2)
        
        with col_exp1:
            # Export Excel
            excel_filename = f"planning_{start_date.strftime('%Y%m%d')}.xlsx"
            
            if st.button("📥 Télécharger Excel", use_container_width=True):
                # Générer le fichier Excel
                st.session_state.scheduler.export_to_excel(schedule, excel_filename)
                
                # Lire le fichier pour le téléchargement
                with open(excel_filename, 'rb') as f:
                    st.download_button(
                        label="⬇️ Cliquez ici pour télécharger",
                        data=f,
                        file_name=excel_filename,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
        
        with col_exp2:
            # Export JSON
            json_data = json.dumps(schedule, indent=2, ensure_ascii=False)
            st.download_button(
                label="📥 Télécharger JSON",
                data=json_data,
                file_name=f"planning_{start_date.strftime('%Y%m%d')}.json",
                mime='application/json',
                use_container_width=True
            )
        
        # Diagnostic
        with st.expander("🔍 Voir le diagnostic du planning"):
            diagnosis = st.session_state.scheduler.diagnose_schedule(schedule)
            
            if diagnosis['violations']:
                st.error(f"⚠️ {len(diagnosis['violations'])} violation(s) détectée(s)")
                for v in diagnosis['violations'][:5]:
                    st.write(f"- {v}")
            else:
                st.success("✅ Aucune violation détectée")
            
            if diagnosis['warnings']:
                st.warning(f"⚠️ {len(diagnosis['warnings'])} avertissement(s)")
                for w in diagnosis['warnings']:
                    st.write(f"- {w}")

# TAB 2 : Historique
with tab2:
    st.header("📚 Historique des plannings")
    
    if not history.history:
        st.info("📋 Aucun historique disponible. Générez votre premier planning !")
    else:
        # Afficher l'historique
        for period_name, period_data in history.history.items():
            with st.expander(f"📅 {period_name} ({period_data['date_debut']} → {period_data['date_fin']})"):
                
                # Créer un DataFrame pour les stats
                stats_list = []
                for vet, stats in sorted(period_data['stats'].items()):
                    stats_list.append({
                        'Vétérinaire': vet,
                        '1er semaine': stats['premier_semaine'],
                        '1er WE': stats['premier_weekend'] // 2,
                        'Rappelable': stats['rappelable_semaine'],
                        '2ème WE': stats['deuxieme_weekend'] // 2
                    })
                
                df_stats = pd.DataFrame(stats_list)
                st.dataframe(df_stats, use_container_width=True, hide_index=True)
        
        # Statistiques cumulées
        st.markdown("---")
        st.subheader("📊 Statistiques cumulées (tous les plannings)")
        
        cumul_stats = history.get_cumulative_stats(list(default_vets.keys()))
        
        cumul_list = []
        for vet_name in sorted(cumul_stats.keys()):
            stats = cumul_stats[vet_name]
            total = (stats['premier_semaine'] + 
                    stats['premier_weekend'] // 2 + 
                    stats['rappelable_semaine'] + 
                    stats['deuxieme_weekend'] // 2)
            
            cumul_list.append({
                'Vétérinaire': vet_name,
                '1er semaine': stats['premier_semaine'],
                '1er WE': stats['premier_weekend'] // 2,
                'Rappelable': stats['rappelable_semaine'],
                '2ème WE': stats['deuxieme_weekend'] // 2,
                'Total': total
            })
        
        df_cumul = pd.DataFrame(cumul_list)
        st.dataframe(df_cumul, use_container_width=True, hide_index=True)
        
        # Graphique
        st.bar_chart(df_cumul.set_index('Vétérinaire')[['1er semaine', '1er WE', 'Rappelable', '2ème WE']])

# TAB 3 : Aide
with tab3:
    st.header("ℹ️ Guide d'utilisation")
    
    st.markdown("""
    ### 🎯 Comment utiliser cette application
    
    #### 1. Configurer la période
    - Choisissez une **date de début** (de préférence un lundi)
    - Choisissez une **date de fin** (de préférence un dimanche)
    - Idéalement : 4-5 semaines complètes
    
    #### 2. Ajouter des congés (optionnel)
    - Utilisez la section "➕ Ajouter des congés" dans la barre latérale
    - Sélectionnez le vétérinaire et la date
    - Les congés seront pris en compte lors de la génération
    
    #### 3. Générer le planning
    - Cliquez sur **"🚀 Générer le planning"**
    - Patientez 10-30 secondes (calcul complexe)
    - Le planning apparaît automatiquement
    
    #### 4. Exporter
    - **Excel** : Fichier formaté avec couleurs et statistiques
    - **JSON** : Format brut pour traitement informatique
    
    #### 5. Consulter l'historique
    - Onglet "📊 Historique"
    - Voir tous les plannings générés
    - Statistiques cumulées pour équité sur plusieurs mois
    
    ---
    
    ### 📋 Règles automatiques
    
    - ✅ 1 premier de garde + 1 rappelable par jour de semaine
    - ✅ 1 premier + 1 deuxième le week-end (même duo samedi-dimanche)
    - ✅ Maximum 1 garde premier par semaine civile
    - ✅ Maximum 2 rappelables par semaine civile
    - ✅ Repos obligatoire après une garde premier en semaine
    - ✅ Équilibrage automatique entre tous les vétérinaires
    - ✅ Respect des jours de repos hebdomadaires
    - ✅ Respect des congés
    - ✅ **Dr. Olivier** : uniquement rappelable en semaine (max 1 sur 2 semaines)
    - ✅ **Dr. Laura** : peut être de garde le week-end même si repos le lundi
    - ✅ **Dr. Julien** : peut être de garde la veille de ses jours de repos
    
    ---
    
    ### ⚠️ En cas de problème
    
    **"Impossible de générer un planning"** :
    - Vérifiez qu'il n'y a pas trop de congés simultanés
    - Essayez avec une période plus longue
    - Réduisez les contraintes si possible
    
    **Le planning semble déséquilibré** :
    - Consultez l'historique : l'équilibrage se fait sur plusieurs mois
    - Le diagnostic automatique vous alerte en cas de problème
    
    ---
    
    ### 🔒 Sécurité des données
    
    - L'historique est stocké localement sur le serveur
    - Aucune donnée n'est partagée avec des tiers
    - Sauvegarde automatique après chaque génération
    """)

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        <p>Planning de Gardes Vétérinaires v1.0 | Propulsé par Streamlit</p>
    </div>
    """,
    unsafe_allow_html=True
)
