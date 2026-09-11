# TN-Montants V3.4 — Mise à jour du pack

Cette mise à jour adapte le pack V3.4 au mode de déploiement retenu : **modèle Word `.dotm` dédié chargé globalement**.

## Fichiers ajoutés/modifiés

```text
ribbon/customUI14.xml
 documentation/INSTALL.md
 documentation/DEPLOYMENT.md
 deployment/Install-TN-Montants.ps1
 deployment/Uninstall-TN-Montants.ps1
```

Le fichier `customUI14.xml` remplace `ribbon/customUI.xml` pour le modèle `.dotm` final.

## Principe

```text
Documents existants (.docx/.docm)
          │
          ▼
TN-Montants-V3.4.dotm
          ├── VBA
          └── RibbonX
```

Le modèle ne doit pas être fusionné avec `Normal.dotm`.

Le fichier `.exportedUI` n'est pas nécessaire dans cette architecture : l'interface Ruban est embarquée directement dans le `.dotm`. Une personnalisation `.exportedUI` peut rester une solution alternative pour une configuration individuelle, mais elle n'est pas retenue comme composant de distribution V3.4.
