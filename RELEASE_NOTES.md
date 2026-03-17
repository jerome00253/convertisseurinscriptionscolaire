# Release v1.0.0 - Première version stable 🚀

Voici la première version officielle du **Convertisseur Inscription Scolaire**. Cet outil a été conçu pour automatiser et styliser la transformation des exports d'inscriptions (Illzach) vers le format standardisé EMS.

## 🌟 Points forts de la version 1.0.0

### Interface & Ergonomie
- **Application Windows native** : Interface graphique intuitive avec le logo de la ville.
- **Icône personnalisée** : Facile à identifier dans l'explorateur de fichiers.
- **Calendrier interactif** : Sélectionnez vos dates de filtrage via un widget calendrier (`tkcalendar`).

### Fonctionnalités de Traitement
- **Sélection des onglets** : Choisissez dynamiquement les feuilles de calcul à traiter depuis le fichier source.
- **Filtrage avancé** : 
    - Dates de création (optionnel, avec bouton d'activation).
    - État de dérogation (Oui, Non, ou Tous).
- **Organisation intelligente** : Création automatique d'un onglet par école dans le fichier Excel de sortie.

### Mise en forme professionnelle (Excel)
- **Alignement Vertical "Haut"** : Harmonisation de l'affichage pour toutes les cellules.
- **Lecture optimisée** : Élargissement des colonnes de raison et d'adresses.
- **Style Premium** : En-têtes en bleu ciel, texte en gras, filtres automatiques activés et volets figés.

---

## 📦 Comment installer / utiliser

- **Pour les utilisateurs** : Téléchargez le fichier `ConvertisseurInscriptionScolaire.exe` dans le dossier `dist`. Aucun Python requis.
- **Pour les développeurs** :
    1. Installez les dépendances : `pip install -r requirements.txt`
    2. Lancez via `python converter_app.py`

---
*Développé avec ❤️ pour la ville d'Illzach.*
