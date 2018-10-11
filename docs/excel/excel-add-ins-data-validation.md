---
title: Ajouter de la validation des données aux plages Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 9e3aba8d87e84405bb3e1ae35a8d35d60ce8e2b6
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459153"
---
# <a name="add-data-validation-to-excel-ranges"></a>Ajouter de la validation des données aux plages Excel

La bibliothèque JavaScript Excel fournit des API pour permettre à votre complément d’ajouter de la validation automatique des données à des tableaux, des colonnes, des lignes et autres plages dans un classeur. Pour comprendre les concepts et la terminologie de la validation des données, consultez les articles suivants sur la façon dont les utilisateurs ajoutent de la validation des données par le biais de l’interface utilisateur d’Excel :

- [Appliquer de la validation des données aux cellules](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Plus d'informations sur la validation des données](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Description et exemples de validation de données dans Excel](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Contrôle programmatique de la validation des données

La propriété `Range.dataValidation`, qui accepte un objet [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) , est le point d’entrée du contrôle programmatique de la validation des données dans Excel. L’objet`DataValidation` a cinq propriétés :

- `rule` — définir ce qui constitue des données valides pour la plage. Voir [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).
- `errorAlert` — Spécifie si une erreur apparaît lorsque l’utilisateur entre des données non valides et définit le texte, le titre et le style d’alerte. Par exemple, **Informatif**, **Avertissement**, et **Arrêter**. Voir [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).
- `prompt` — Spécifie si une invite s’affiche lorsque l’utilisateur survole la plage et définir le message d’assistance vocale. Voir [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).
- `ignoreBlanks` — Spécifie si la règle de validation des données s’applique aux cellules vides de la plage. `true` par défaut.
- `type` — une identification en lecture seule du type de validation, tel que WholeNumber, Date, TextLength, etc. Elle est définie indirectement lorsque vous définissez la propriété `rule`.

> [!NOTE]
> La validation des données ajoutée par programme se comporte comme celle ajoutée manuellement. Notez en particulier que la validation des données est déclenchée uniquement si l’utilisateur entre directement une valeur dans une cellule ou copie et colle une cellule à partir d’un autre endroit du classeur et choisit l’option de collage **Valeurs**. Si l’utilisateur copie une cellule et effectue un collage brut dans une plage de validation des données, la validation n’est pas déclenchée.

## <a name="creating-validation-rules"></a>Création des règles de validation

Pour ajouter de la validation des données à une plage, votre code doit définir la propriété `rule` de l’objet `DataValidation` dans `Range.dataValidation`. Elle accepte un objet [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) qui a sept propriétés facultatives. *Seulement une de ces propriétés peut être présente dans un objet `DataValidationRule`.* La propriété que vous incluez détermine le type de validation.

### <a name="basic-and-datetime-validation-rule-types"></a>Types de règles de validation de base et DateTime

Les trois premières propriétés `DataValidationRule` (c.-à-d. les types de règles de validation) prennent un objet [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) comme leur valeur.

- `wholeNumber` – Nécessite un nombre entier en plus de toute autre validation spécifiée par l'objet `BasicDataValidation`.
- `decimal` – Nécessite un nombre décimal en plus de toute autre validation spécifiée par l'objet `BasicDataValidation`.
- `textLength` – Applique les détails de validation dans l'objet `BasicDataValidation` à la *longueur*  de la valeur de la cellule.

Voici un exemple de création d’une règle de validation. Notez ce qui suit concernant ce code :

- Le `operator` est l’opérateur binaire « Supérieur ». Chaque fois que vous utilisez un opérateur binaire, la valeur que l’utilisateur essaie d’entrer dans la cellule est l’opérande gauche et la valeur spécifiée dans `formula1` est l’opérande droit. Cette règle indique donc que seuls les nombres entiers qui sont supérieures à 0 sont valides. 
- Le `formula1` est un nombre codé en dur. Si vous ne savez quelle valeur utiliser au moment du codage, vous pouvez également utiliser une formule Excel (en tant que chaîne) pour la valeur. Par exemple, "=A3" et "=SUM(A4,B5)" peuvent également être des valeurs de `formula1`.

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

Voir [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) pour une liste des autres opérateurs binaires. 

Il y a aussi deux opérateurs ternaires : « Between » et « NotBetween ». Pour les utiliser, vous devez spécifier la propriété optionnelle `formula2`. Les valeurs `formula1` et `formula2` sont les opérandes limites. La valeur que l’utilisateur essaie d’entrer dans la cellule est le troisième opérande (évalué). Voici un exemple d’utilisation de l’opérateur « Between » :

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

Les deux propriétés de règle suivantes prennent un objet [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) comme valeur.

- `date`
- `time`

L’objet `DateTimeDataValidation` est structuré de la même façon que le `BasicDataValidation` : il possède les propriétés `formula1`, `formula2`, et `operator` et s’utilise de manière similaire. La différence est que vous ne pouvez pas utiliser un nombre dans les propriétés de formule, mais vous pouvez entrer une chaîne [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou une formule Excel). Voici un exemple qui définit les valeurs valides comme des dates dans la première semaine d’avril 2018. 

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

### <a name="list-validation-rule-type"></a>Type de règle de validation de liste

Utilisez la propriété `list` dans l’objet `DataValidationRule` pour spécifier que les seules valeurs valides sont celles d’une liste finie. Voici un exemple. Notez ce qui suit concernant ce code :

- Il suppose qu’il existe une feuille de calcul nommée « Names » et que les valeurs de la plage « A1:A3 » sont des noms.
- La propriété `source` définit la liste des valeurs valides. La plage avec les noms été lui a été affectée. Vous pouvez également affecter une liste délimitée par des virgules (par exemple : « Sue, Ricky, Liz »). 
- La propriété `inCellDropDown` spécifie si un contrôle de liste déroulante apparaît dans la cellule lorsque l’utilisateur la sélectionne. Si la valeur est définie sur `true`, alors la liste déroulante s’affiche avec la liste de valeurs à partir de `source`.

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

### <a name="custom-validation-rule-type"></a>Type de règle de validation personnalisée

Utilisez la propriété `custom` dans l’objet `DataValidationRule` pour spécifier une formule de validation personnalisée. En voici un exemple. Notez ce qui suit à propos de ce code :

- Il suppose qu’il y a un tableau à deux colonnes avec des colonnes **Athlete Name** et **Comments** dans les colonnes A et B de la feuille de calcul.
- Pour réduire la verbosité dans la colonne **Comments**, il rend invalides les données qui incluent le nom de l’athlète.
- `SEARCH(A2,B2)` Renvoie la position de départ, dans la chaîne dans B2, de la chaîne dans A2. Si A2 n’est pas incluse dans B2, elle ne renvoie pas un nombre. `ISNUMBER()` renvoie une valeur booléenne. La propriété `formula` indique donc que les données valides pour la colonne **Comment** sont les données qui n’incluent pas la chaîne dans la colonne **Athlete Name** .

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

## <a name="create-validation-error-alerts"></a>Création des alertes d’erreur de validation

Vous pouvez créer une alerte d’erreur personnalisée qui s’affiche lorsqu’un utilisateur essaie d’entrer des données non valides dans une cellule. Voici un exemple simple. Notez ce qui suit concernant ce code :

- La propriété `style`  détermine si l’utilisateur reçoit une alerte d’information, un avertissement ou une alerte « stop ». Seul `Stop` empêche réellement l’utilisateur d’ajouter des données non valides. La fenêtre contextuelle pour `Warning` et `Information` a des options qui permettent de saisir quand même les données non valides.
- La valeur par défaut de la propriété `showAlert` est `true`. Cela signifie que l’hôte Excel affichera une alerte générique (de type `Stop`) sauf si vous créez une alerte personnalisée qui définit `showAlert` sur `false` ou définit un message personnalisé, un titre et un style. Ce code définit un message personnalisé et un titre.


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

Pour en savoir plus, voir [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).

## <a name="create-validation-prompts"></a>Création des invites de validation

Vous pouvez créer une invite d’instruction qui apparaît lorsqu’un utilisateur survole ou sélectionne une cellule à laquelle la validation des données a été appliquée. Voici un exemple :

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

Pour en savoir plus, voir [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).

## <a name="remove-data-validation-from-a-range"></a>Suppression de la validation des données d’une plage

Pour supprimer la validation des données d'une plage, appelez la méthode [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--).

```js
myrange.dataValidation.clear()
```

Il n’est pas nécessaire que la plage que vous désactivez soit exactement la même plage que celle sur laquelle vous avez ajouté la validation des données. Si ce n’est pas le cas, seules les cellules qui se chevauchent dans les deux plages sont désactivées le cas échéant. 

> [!NOTE]
> La désactivation de la validation des données d’une plage efface également toute validation de données qu’un utilisateur a ajoutée manuellement à la plage.

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Objet DataValidation (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [Objet Range (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
