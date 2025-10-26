import json
from datetime import datetime, timedelta
from Planning_gardes import VetScheduler, VetSchedulerHistory

def parse_conges(conges_list):
    """
    Parse les congés avec format mixte : dates individuelles ou périodes
    
    Formats acceptés :
    - "2026-01-15" : date unique
    - "2026-01-15:2026-01-28" : période du 15 au 28 janvier
    - "2026-01-15 : 2026-01-28" : période avec espaces (accepté aussi)
    
    Args:
        conges_list: Liste de chaînes représentant dates ou périodes
        
    Returns:
        Liste de dates individuelles au format 'YYYY-MM-DD'
    """
    dates_conges = []
    
    for conge in conges_list:
        conge = conge.strip()  # Enlever les espaces
        
        # Vérifier si c'est une période (contient ":")
        if ':' in conge:
            # Séparer la date de début et de fin
            parts = conge.split(':')
            if len(parts) != 2:
                raise ValueError(f"Format de période invalide: '{conge}'. Utilisez 'YYYY-MM-DD:YYYY-MM-DD'")
            
            date_debut_str = parts[0].strip()
            date_fin_str = parts[1].strip()
            
            try:
                date_debut = datetime.strptime(date_debut_str, '%Y-%m-%d')
                date_fin = datetime.strptime(date_fin_str, '%Y-%m-%d')
            except ValueError as e:
                raise ValueError(f"Format de date invalide dans '{conge}': {e}")
            
            if date_debut > date_fin:
                raise ValueError(f"La date de début doit être avant la date de fin: '{conge}'")
            
            # Générer toutes les dates entre début et fin (inclus)
            current = date_debut
            while current <= date_fin:
                dates_conges.append(current.strftime('%Y-%m-%d'))
                current += timedelta(days=1)
        else:
            # C'est une date unique
            try:
                datetime.strptime(conge, '%Y-%m-%d')  # Validation
                dates_conges.append(conge)
            except ValueError as e:
                raise ValueError(f"Format de date invalide: '{conge}'. Utilisez 'YYYY-MM-DD'")
    
    return dates_conges

def load_config(filepath='config_planning.json'):
    """
    Charge la configuration depuis un fichier JSON
    
    Args:
        filepath: Chemin vers le fichier de configuration
        
    Returns:
        dict: Configuration avec 'periode' et 'veterinaires'
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        # Validation basique
        if 'periode' not in config:
            raise ValueError("Le fichier doit contenir une section 'periode'")
        if 'veterinaires' not in config:
            raise ValueError("Le fichier doit contenir une section 'veterinaires'")
        
        periode = config['periode']
        if 'date_debut' not in periode or 'date_fin' not in periode:
            raise ValueError("La période doit contenir 'date_debut' et 'date_fin'")
        
        print("✅ Configuration chargée avec succès!")
        print(f"   Période: {periode['date_debut']} → {periode['date_fin']}")
        if 'description' in periode:
            print(f"   Description: {periode['description']}")
        print(f"   Vétérinaires: {len(config['veterinaires'])}")
        
        return config
        
    except FileNotFoundError:
        print(f"❌ Fichier non trouvé: {filepath}")
        print("💡 Créez un fichier config_planning.json avec la structure requise")
        raise
    except json.JSONDecodeError as e:
        print(f"❌ Erreur de format JSON: {e}")
        raise
    except Exception as e:
        print(f"❌ Erreur lors du chargement: {e}")
        raise


def generer_planning(config_file='config_planning.json', 
                     output_file=None,
                     time_limit=300):
    """
    Génère un planning de gardes à partir d'un fichier de configuration
    
    Args:
        config_file: Chemin vers le fichier de configuration JSON
        output_file: Nom du fichier Excel de sortie (optionnel)
        time_limit: Temps limite de recherche en secondes
    """
    # Charger la configuration
    config = load_config(config_file)
    
    # Extraire les données
    periode = config['periode']
    veterinaires = config['veterinaires']
    
    # Générer un nom de fichier par défaut si non spécifié
    if output_file is None:
        date_debut = periode['date_debut'].replace('-', '')
        date_fin = periode['date_fin'].replace('-', '')
        output_file = f'planning_{date_debut}_{date_fin}.xlsx'
    
    # Charger l'historique
    history = VetSchedulerHistory('historique_gardes.json')
    
    # Afficher l'historique existant
    if history.history:
        print("\n📚 HISTORIQUE CHARGÉ:")
        history.print_history()
    
    # Générer le planning
    print("\n" + "="*80)
    print("GÉNÉRATION DU PLANNING")
    print("="*80)
    
    scheduler = VetScheduler(
        start_date=periode['date_debut'],
        end_date=periode['date_fin'],
        veterinaires=veterinaires,
        history=history
    )
    
    schedule = scheduler.solve(time_limit=time_limit)
    
    if schedule:
        scheduler.print_schedule(schedule)
        scheduler.export_to_excel(schedule, output_file)
        
        # Ajouter à l'historique avec un nom basé sur la description ou les dates
        period_name = periode.get('description', f"{periode['date_debut']}_to_{periode['date_fin']}")
        history.add_schedule(schedule, period_name=period_name)
        
        print("\n" + "="*80)
        print("🎉 PLANNING GÉNÉRÉ AVEC SUCCÈS !")
        print("="*80)
        print(f"\n📄 Fichier Excel: {output_file}")
        print("💡 L'historique a été mis à jour pour les prochains plannings")
        
        return schedule
    else:
        print("\n❌ Impossible de générer le planning")
        print("💡 Vérifiez les contraintes et les congés dans config_planning.json")
        return None


def afficher_aide():
    """Affiche l'aide pour utiliser le système"""
    print("""
╔════════════════════════════════════════════════════════════════════════════╗
║           SYSTÈME DE GÉNÉRATION DE PLANNING DE GARDES                      ║
╚════════════════════════════════════════════════════════════════════════════╝

📋 UTILISATION SIMPLE:

1. Éditez le fichier 'config_planning.json' avec:
   - Les dates de début et fin de période
   - Les congés des vétérinaires (si besoin)
   - Les jours de repos (normalement fixes)

2. Lancez le script:
   python Planning_gardes_avec_config.py

3. Le planning sera généré automatiquement!

📝 FORMAT DES CONGÉS:
   "conges": ["2026-01-15", "2026-01-16", "2026-02-03"]

📅 JOURS DE LA SEMAINE:
   0 = Lundi    4 = Vendredi
   1 = Mardi    5 = Samedi
   2 = Mercredi 6 = Dimanche
   3 = Jeudi

💡 EXEMPLES:
   
   Ajouter des congés pour Dr. Laura:
   "Dr. Laura": {
     "jour_repos": 0,
     "conges": ["2026-01-20", "2026-01-21", "2026-01-22"]
   }
   
   Changer la période:
   "periode": {
     "date_debut": "2026-02-09",
     "date_fin": "2026-03-15",
     "description": "Planning février-mars 2026"
   }

🔧 FONCTIONS AVANCÉES:

   # Avec paramètres personnalisés
   from Planning_gardes_avec_config import generer_planning
   
   generer_planning(
       config_file='mon_config.json',
       output_file='mon_planning.xlsx',
       time_limit=600
   )
   
   # Avec historique
   from Planning_gardes_janvier import VetSchedulerHistory
   history = VetSchedulerHistory()
   history.print_history()
   history.clear()  # Effacer l'historique si besoin

╚════════════════════════════════════════════════════════════════════════════╝
    """)


if __name__ == "__main__":
    import sys
    
    # Si l'utilisateur demande de l'aide
    if len(sys.argv) > 1 and sys.argv[1] in ['-h', '--help', 'aide', 'help']:
        afficher_aide()
    else:
        # Génération normale du planning
        try:
            generer_planning()
        except Exception as e:
            print(f"\n❌ Erreur: {e}")
            print("\n💡 Utilisez 'python Planning_gardes_avec_config.py --help' pour l'aide")