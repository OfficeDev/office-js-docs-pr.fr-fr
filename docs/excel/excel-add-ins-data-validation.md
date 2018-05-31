---
title: Ajouter la validation des données aux plages Excel
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: 8e5f09f1c566103f34ad584885769229c17ab1f7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437527"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a>Ajouter la validation des données aux plages Excel (préversion)

> [!NOTE]
> Tant que les API de validation des données sont en préversion, vous devez charger la version bêta de la bibliothèque JavaScript Office pour les utiliser. L'URL est https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Si vous utilisez TypeScript ou si votre éditeur de code utilise un fichier de définition de type TypeScript pour intelliSense, utilisez https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

La bibliothèque JavaScript Excel fournit des API pour permettre à votre complément d'ajouter une validation automatique des données aux tables, colonnes, lignes et autres plages d'un classeur. Pour comprendre les concepts et la terminologie de la validation des données, consultez les articles suivants qui portent sur la manière dont les utilisateurs peuvent ajouter la validation des données via l'IU Excel :

- [Appliquer la validation des données aux cellules](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Plus d'informations sur la validation des données](https://microsoft.sharepoint.com/:p:/r/teams/oext/_layouts/15/Doc.aspx?sourcedoc=%7B51143964-d52c-429d-bfac-c7495473d536%7D&action=edit)
- [Description et exemples de validation de données dans Excel](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Contrôle programmatique de la validation des données

La propriété `Range.dataValidation`, qui prend un objet de validation de données [,](https://dev.office.com/reference/add-ins/excel/datavalidation) est le point d'entrée pour le contrôle programmatique de la validation des données dans Excel. Il y a cinq propriétés à l'objet `DataValidation` :

- `rule` - Définit ce qui constitue des données valides pour la plage. Voir [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).
- `errorAlert` - Spécifie si une erreur apparaît lorsque l'utilisateur entre des données non valides et définit le texte, le titre et le style d'alerte ; par exemple, **Informatif**, **Avertissement**, et **Arrêter**. Voir [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).
- `prompt` - Indique si une invite s'affiche lorsque l'utilisateur survole la plage et définit le message d'assistance vocale. Voir [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).
- `ignoreBlanks` - Spécifie si la règle de validation des données s'applique aux cellules vides de la plage. Par défaut `true`.
- `type` - Une identification en lecture seule du type de validation, tel que WholeNumber, Date, TextLength, etc. Il est défini indirectement lorsque vous définissez la propriété `rule`.

> [!NOTE]
> La validation des données ajoutée de façon programmatique se comporte exactement comme la validation des données ajoutée manuellement. Surtout, notez que la validation des données est déclenchée uniquement si l'utilisateur entre directement une valeur dans une cellule ou copie et colle une cellule d'un autre emplacement dans le classeur et choisit l'option de collage, valeurs ****. Si l'utilisateur copie une cellule et la colle simplement dans une plage avec validation des données, la validation n'est pas déclenchée.

### <a name="creating-validation-rules"></a>Créer des règles de validation

Pour ajouter une validation de données à une plage, votre code doit définir la propriété `rule` de l'objet `DataValidation` dans `Range.dataValidation`. Cela prend un objet [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) qui a sept propriétés facultatives. *Une seule de ces propriétés peut être présente dans un objet `DataValidationRule`.* La propriété que vous incluez détermine le type de validation.

#### <a name="basic-and-datetime-validation-rule-types"></a>Types de règles de validation de base et DateTime

Les trois premières propriétés `DataValidationRule` (c.-à-d. les types de règles de validation) prennent un objet [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) comme leur valeur.

- `wholeNumber` – Nécessite un nombre entier en plus de toute autre validation spécifiée par l'objet `BasicDataValidation`.
- `decimal` - Nécessite un nombre décimal en plus de toute autre validation spécifiée par l'objet `BasicDataValidation`.
- `textLength` – Applique les détails de validation dans l'objet `BasicDataValidation` à la *longueur*  de la valeur de la cellule.

Voici un exemple de création d'une règle de validation. Tenez compte des informations suivantes concernant ce code :

- Le `operator` est l'opérateur binaire "supérieur à". Chaque fois que vous utilisez un opérateur binaire, la valeur que l'utilisateur essaie d'entrer dans la cellule est l'opérande de gauche et la valeur spécifiée dans `formula1` est l'opérande de droite. Donc, cette règle dit que seuls les nombres entiers supérieurs à 0 sont valides. 
- Le `formula1` est un nombre codé en dur. Si vous ne savez pas au moment du codage quelle devrait être la valeur, vous pouvez également utiliser une formule Excel (sous forme de chaîne) pour la valeur. Par exemple, « = A3 » et « = SOMME (A4:B5) » peuvent également être des valeurs de `formula1`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            wholeNumber: {
                formula1: 0,
                operator: "GreaterThan"
            }
        };

    return context.sync();
})
```

Voir [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) pour une liste des autres opérateurs binaires. 

Il existe aussi deux opérateurs ternaires : « Between » et « NotBetween ». Pour les utiliser, vous devez spécifier la propriété `formula2` facultative. Les valeurs `formula1` et `formula2` sont les opérandes de délimitation. La valeur que l'utilisateur essaie d'entrer dans la cellule est le troisième opérande (évaluée). Voici un exemple d'utilisation de l'opérateur « Between » :

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            decimal: {
                formula1: 0,
                formula2: 100,
                operator: "Between"
            }
        };

    return context.sync();
})
```

Les deux propriétés de règle suivantes prennent un objet [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) comme leur valeur.

- `date`
- `time`

L'objet `DateTimeDataValidation` est structuré de manière similaire à la `BasicDataValidation` : il a les propriétés `formula1`, `formula2`, et `operator`, et est utilisé de la même manière. La différence est que vous ne pouvez pas utiliser un nombre dans les propriétés de la formule, mais vous pouvez entrer une chaîne [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou une formule Excel). Voici un exemple qui définit des valeurs valides telles que des dates dans la première semaine d'avril 2018. 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            date: {
                formula1: "2018-04-01",
                formula2: "2018-04-08",
                operator: "Between"
            }
        };

    return context.sync();
})
```

#### <a name="list-validation-rule-type"></a>Type de règle de validation de liste

Utilisez la propriété `list` dans l'objet `DataValidationRule` pour spécifier que les seules valeurs valides sont celles d'une liste finie. Voir l'exemple suivant. Tenez compte des informations suivantes concernant ce code :

- Il suppose qu'il existe une feuille de calcul nommée "Noms" et que les valeurs de la plage "A1: A3" sont des noms.
- La propriété `source` spécifie la liste des valeurs valides. La plage avec les noms lui a été affectée. Vous pouvez également affecter une liste délimitée par des virgules, comme par exemple : « Sue, Ricky, Liz ». 
- La propriété `inCellDropDown` spécifie si un contrôle déroulant apparaîtra dans la cellule lorsque l'utilisateur le sélectionne. Si elle est définie sur `true`, une liste déroulante apparaît contenant la liste des valeurs du `source`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: nameSourceRange
        }
    };

    return context.sync();
})
```

#### <a name="custom-validation-rule-type"></a>Type de règle de validation personnalisée

Utilisez la propriété `custom` dans l'objet `DataValidationRule` pour spécifier une formule de validation personnalisée. Voir l'exemple suivant. Tenez compte des informations suivantes concernant ce code :

- Il suppose qu'il y a un tableau à deux colonnes avec des colonnes **Nom de l'athlète** et **Commentaires** dans les colonnes A et B de la feuille de calcul.
- Pour réduire la verbosité dans la colonne **Commentaires,** il rend invalides les données qui incluent le nom de l'athlète.
- `SEARCH(A2,B2)` renvoie la position de départ, de la chaîne dans B2, de la chaîne dans A2. Si A2 n'est pas contenu dans B2, il ne renvoie pas de nombre. `ISNUMBER()` retourne un booléen. La propriété `formula` indique donc que les données valides pour la colonne **Commentaire** sont les données qui n'incluent pas la chaîne présente dans la colonne **Nom de l'athlète**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var commentsRange = sheet.tables.getItem("AthletesTable").columns.getItem("Comments").getDataBodyRange();

    commentsRange.dataValidation.rule = {
            custom: {
                formula: "=NOT(ISNUMBER(SEARCH(A2,B2)))"
            }
        };

    return context.sync();
})
```

### <a name="create-validation-error-alerts"></a>Créer des alertes d'erreur de validation

Vous pouvez créer une alerte d'erreur personnalisée qui apparaît lorsqu'un utilisateur tente d'entrer des données non valides dans une cellule. Ce qui suit est un simple exemple. Tenez compte des informations suivantes concernant ce code :

- La propriété `style` détermine si l'utilisateur reçoit une alerte informative, un avertissement ou une alerte d' "arrêt". Seule `Stop` empêche réellement l'utilisateur d'ajouter des données invalides. La fenêtre contextuelle pour `Warning` et `Information` a des options qui permettent à l'utilisateur d'entrer les données invalides de toute façon.
- La propriété `showAlert` prend `true` par défaut. Cela signifie que l'hôte Excel affichera une alerte générique (de type `Stop`) sauf si vous créez une alerte personnalisée qui soit définit `showAlert` pour `false` ou définit un message, un titre et un style personnalisés. Ce code définit un message et un titre personnalisés.


```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.errorAlert = {
            message: "Sorry, only positive whole numbers are allowed",
            showAlert: true, // default is 'true'
            style: "Stop", // other possible values: Warning, Information
            title: "Negative or Decimal Number Entered"
        };
    
    // Set range.dataValidation.rule and optionally .prompt here.

    return context.sync();
})
```

Pour plus d'informations, voir [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).

### <a name="create-validation-prompts"></a>Créer des invites de validation

Vous pouvez créer une invite d'instruction qui apparaît lorsqu'un utilisateur survole ou sélectionne une cellule à laquelle la validation des données a été appliquée. Voici un exemple :

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.prompt = {
            message: "Please enter a positive whole number.",
            showPrompt: true, // default is 'false'
            title: "Positive Whole Numbers Only."
        };
    
    // Set range.dataValidation.rule and optionally .errorAlert here.

    return context.sync();
})
```

Pour plus d'informations, voir [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).

### <a name="remove-data-validation-from-a-range"></a>Supprimer la validation des données d'une plage

Pour supprimer la validation des données d'une plage, appelez la méthode [Range.dataValidation.clear ()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear).

```js
myrange.dataValidation.clear()
```

Il n'est pas nécessaire que la plage que vous effacez soit exactement la même plage que celle sur laquelle vous avez ajouté la validation des données. Si ce n'est pas le cas, seules les cellules des deux plages qui se chevauchent sont effacées, le cas échéant. 

> [!NOTE]
> L'effacement de la validation des données d'une plage efface également toute validation de données qu'un utilisateur a ajoutée manuellement à la plage.

## <a name="see-also"></a>Voir aussi

- [Concepts de base de l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Objet DataValidation (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/datavalidation)
- [Objet Range (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/range)



 
