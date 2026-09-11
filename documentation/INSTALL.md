# TN-Montants V3.4 — Installation Word

## Architecture recommandée

TN-Montants est distribué sous la forme d'un modèle Word dédié :

```text
TN-Montants-V3.4.dotm
```

Le modèle contient le moteur VBA et l'interface Ruban. Il est chargé globalement par Word, de sorte que les documents existants (`.docx` ou `.docm`) puissent utiliser TN-Montants sans être modifiés.

**Ne pas installer TN-Montants dans `Normal.dotm`.**

## 1. Construire le modèle `.dotm`

Cette étape est nécessaire une seule fois pour préparer le fichier distribuable.

1. Ouvrir Word.
2. Créer un document/modèle vierge.
3. `Alt+F11` → importer :
   - `src/TN_Montants_V33.bas`
   - `src/TN_Montants_Ribbon.bas`
4. Enregistrer le fichier comme :
   `TN-Montants-V3.4.dotm`
5. Fermer complètement Word avant d'insérer le Ruban.
6. Ouvrir `TN-Montants-V3.4.dotm` avec un éditeur RibbonX compatible Office 2010+.
7. Ajouter un **Office 2010+ Custom UI Part** contenant le fichier :
   `ribbon/customUI14.xml`
8. Enregistrer le `.dotm` et le rouvrir dans Word.

Le namespace utilisé par `customUI14.xml` est celui du schéma Office 2010+.

## 2. Installation globale sur un PC collaborateur

Copier `TN-Montants-V3.4.dotm` dans le dossier Startup de Word de l'utilisateur :

```text
%APPDATA%\\Microsoft\\Word\\STARTUP\\
```

Si le dossier `STARTUP` n'existe pas, le créer.

Redémarrer Word.

Le modèle est alors chargé automatiquement et ses macros sont disponibles dans les documents Word ouverts par l'utilisateur.

## 3. Vérification

Ouvrir un document Word existant contenant, par exemple, un montant :

```text
2523,551 d
```

Le Ruban doit afficher l'onglet :

```text
المبالغ التونسية
```

avec :

- `المبلغ المحدد`
- `مبالغ التحديد`
- `جميع مبالغ الوثيقة`

Tester d'abord sur une copie du document.

## 4. Sécurité des macros

Sur les PC des collaborateurs, Word doit autoriser l'exécution des macros du modèle. Si l'organisation applique une politique de sécurité restrictive, il peut être nécessaire d'utiliser un emplacement approuvé ou une signature numérique VBA.

Pour un déploiement professionnel, la signature numérique est recommandée avant diffusion large.

## 5. Mise à jour

Pour passer à une nouvelle version :

1. fermer Word ;
2. remplacer `TN-Montants-V3.4.dotm` dans le dossier Startup ;
3. redémarrer Word ;
4. exécuter la suite de régression avant de diffuser la nouvelle version.

Les documents utilisateurs ne doivent pas être modifiés lors de cette mise à jour.

## 6. Régression

Importer `tests/TN33_RegressionTests.bas` dans un environnement de test et exécuter `TN33_RunRegressionTests`.

`tests/expected-results.csv` est le Golden Master figé. Le testeur ne doit jamais produire les résultats attendus à partir du moteur testé.

## 7. Excel

La partie Excel reste indépendante de Word :

```excel
=TN_Montant(B2)
```

Voir `src/TN_Montants_Excel.bas` et `examples/Excel_Worksheet_Change.bas`.
