# TN-Montants V3.4 — Déploiement

## Objectif

Installer TN-Montants sur plusieurs PC tout en conservant les documents Word existants inchangés.

## Fichier à distribuer

Le composant principal est :

```text
TN-Montants-V3.4.dotm
```

Il contient :

- le moteur VBA V3.3 servant de baseline V3.4 ;
- les callbacks du Ruban ;
- le Ruban RibbonX ;
- les trois commandes de conversion.

## Installation manuelle

Copier le `.dotm` dans :

```text
%APPDATA%\\Microsoft\\Word\\STARTUP\\
```

Puis redémarrer Word.

## Installation automatisée

Le script `deployment/Install-TN-Montants.ps1` copie le modèle vers le dossier Startup de l'utilisateur courant.

Exemple :

```powershell
powershell -ExecutionPolicy Bypass -File .\\Install-TN-Montants.ps1 -DotmPath .\\TN-Montants-V3.4.dotm
```

Le script ne modifie pas `Normal.dotm` et ne touche pas aux documents de l'utilisateur.

## Désinstallation

Fermer Word puis exécuter :

```powershell
powershell -ExecutionPolicy Bypass -File .\\Uninstall-TN-Montants.ps1
```

## Évolution ultérieure

Une fois le déploiement stabilisé, une signature numérique du projet VBA peut être ajoutée pour renforcer la confiance et faciliter la gestion de la sécurité des macros.
