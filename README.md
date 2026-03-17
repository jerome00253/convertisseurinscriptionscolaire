# Convertisseur Inscription Scolaire

Application Windows permettant de convertir des exports HTML/Excel d'inscriptions scolaires (format Illzach) vers un format de sortie standardisé (EMS).

## Fonctionnalités
- **Interface Graphique** : Facile d'utilisation avec le logo de la ville.
- **Sélection des Onglets** : Choisissez quels onglets de l'export traiter.
- **Sélection des Écoles (Nouveau)** : Choisissez spécifiquement les écoles à inclure dans l'export via une liste dynamique.
- **Onglet de Synthèse (Nouveau)** : Premier onglet récapitulatif avec statistiques (totaux, dérogations) et graphiques illustrant la répartition par école.
- **Filtrage par Date** : Filtrez les inscriptions par date de création (Optionnel).
- **Filtre Dérogation** : Filtrez les dossiers avec ou sans besoin de dérogation.
- **Mise en Forme Automatique** : Génère un fichier Excel pro avec en-têtes stylisés, filtres et alignement vertical "Haut".
- **Tri par École** : Crée un onglet par école dans le fichier de sortie.

## Installation (Développement)
1. Clonez le dépôt :
   ```bash
   git clone https://github.com/jerome00253/convertisseurinscriptionscolaire.git
   ```
2. Installez les dépendances :
   ```bash
   pip install -r requirements.txt
   ```
3. Lancez l'application :
   ```bash
   python converter_app.py
   ```

## Compilation en .exe
Pour générer l'exécutable Windows :
```bash
python -m PyInstaller --onefile --noconsole --add-data "1280px-LogoIllzach.jpg;." --collect-submodules tkcalendar --collect-submodules babel --name "ConvertisseurInscriptionScolaire" --icon "app_icon.ico" converter_app.py
```
