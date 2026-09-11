# TN-Dinars2ArabicWords

Conversion des montants en dinars tunisiens, de chiffres vers l'arabe juridique classique, dans Microsoft Word.

**Auteur / Mainteneur :** Habib FARHAT  
**Version :** 3.4  
**Plateforme :** Microsoft Word / VBA

[English](README.md) · [العربية](README.ar.md)

## Présentation

TN-Dinars2ArabicWords est un module VBA pour Microsoft Word qui convertit les montants monétaires tunisiens exprimés en chiffres en leur forme littérale en arabe classique à usage juridique, tout en conservant le montant numérique original entre parenthèses.

Exemple :

```text
2523.551 د
```

devient :

```text
ألفان وخمسمائة وثلاثة وعشرون دينارا وخمسمائة وواحد وخمسون مليما (2523.551 د)
```

Le projet vise notamment les documents juridiques et administratifs tunisiens : contrats, conventions, requêtes, correspondances, procès-verbaux et documents similaires.

## Fonctionnalités

- Dinars tunisiens et millimes.
- Trois chiffres pour les millimes.
- Chiffres occidentaux, arabo-indiens et persans.
- Marqueurs courants : `د`, `د.ت`, `DT`, `TND`, `دينار`.
- Formes singulières, duales et plurielles.
- Formulation arabe classique/juridique.
- Conversion du montant sélectionné, d'une sélection ou du document.
- Protection contre les doubles conversions.
- Génération Unicode par `ChrW`.
- Tests de régression.
- Préfixe `TN33_` pour réduire les collisions dans `Normal.dotm`.

## Exemples

| Montant | Résultat |
|---|---|
| `0.000 د` | `صفر دينار (0.000 د)` |
| `1.000 د` | `دينار واحد (1.000 د)` |
| `2.000 د` | `ديناران (2.000 د)` |
| `6.009 د` | `ستة دنانير وتسعة مليمات (6.009 د)` |
| `33.104 د` | `ثلاثة وثلاثون دينارا ومائة وأربعة مليمات (33.104 د)` |
| `2523.551 د` | `ألفان وخمسمائة وثلاثة وعشرون دينارا وخمسمائة وواحد وخمسون مليما (2523.551 د)` |
| `10004.585 د` | `عشرة آلاف وأربعة دنانير وخمسمائة وخمسة وثمانون مليما (10004.585 د)` |
| `20002.003 د` | `عشرون ألفا وديناران وثلاثة مليمات (20002.003 د)` |
| `50845.801 د` | `خمسون ألفا وثمانمائة وخمسة وأربعون دينارا وثمانمائة ومليم واحد (50845.801 د)` |
| `70575.300 د` | `سبعون ألفا وخمسمائة وخمسة وسبعون دينارا وثلاثمائة مليم (70575.300 د)` |
| `160426.591 د` | `مائة وستون ألفا وأربعمائة وستة وعشرون دينارا وخمسمائة وواحد وتسعون مليما (160426.591 د)` |

## Installation

1. Ouvrir Word.
2. `Alt+F11`.
3. Sélectionner `Normal` / `Normal.dotm`.
4. **File → Import File...** et importer `TN_Montants_V3_3.bas`.
5. **Debug → Compile Normal**.
6. Enregistrer `Normal.dotm`.

N'activez les macros que pour des sources fiables et examinez le fichier `.bas` avant utilisation en production.

## Utilisation

### Montant sélectionné

```vb
TN33_ConvertirMontantSelectionne
```

### Tous les montants d'une sélection

```vb
TN33_ConvertirMontantsSelection
```

### Tous les montants du document

```vb
TN33_ConvertirMontantsDocument
```

### Tests

```vb
TN33_TesterV33
```

`Ctrl+G` affiche la fenêtre d'exécution immédiate du VBE.

## Architecture

Le module sépare :

1. l'intégration Word ;
2. la détection et la normalisation des montants ;
3. la génération des nombres arabes ;
4. la grammaire monétaire ;
5. la génération Unicode ;
6. les tests de régression.

Le préfixe `TN33_` évite autant que possible les collisions avec les éléments déjà présents dans `Normal.dotm`.

## Limites

Le module est conçu pour le dinar tunisien et le millime. Il ne s'agit pas d'un moteur général de conversion des nombres arabes et il ne couvre pas toutes les variantes grammaticales possibles.

Les montants générés dans des documents juridiquement importants doivent être vérifiés avant signature, dépôt ou publication.

## Développement et assistance IA

Le projet a été développé de manière itérative par **Habib FARHAT** avec l'assistance de **OpenAI ChatGPT (GPT-5.6 Luna)**.

L'assistance IA a notamment porté sur l'architecture, le débogage VBA, Unicode, les algorithmes, les règles grammaticales, les tests et la documentation. Les exigences fonctionnelles, les sorties de référence, les préférences linguistiques et les critères de validation ont été définis et vérifiés pendant le développement.

Le dépôt représente donc un processus de développement collaboratif humain–IA.

## Contributions

Les contributions sont les bienvenues, notamment pour les corrections grammaticales, les exemples juridiques tunisiens, les tests, la détection des montants, la compatibilité Word et la documentation.

Toute modification linguistique devrait être accompagnée de tests de régression.

## Licence

Licence MIT. Voir [`LICENSE`](LICENSE).

## Auteur

**Habib FARHAT**

Développé avec l'assistance de OpenAI ChatGPT (GPT-5.6 Luna).

Microsoft Word et VBA sont des marques de Microsoft Corporation. Ce projet n'est ni affilié à Microsoft ni approuvé par Microsoft ou OpenAI.
