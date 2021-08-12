---
title: Ajout de validation des données à des plages Excel
description: Découvrez comment les EXCEL JavaScript permettent à votre add-in d’ajouter la validation automatique des données aux tableaux, colonnes, lignes et autres plages d’un workbook.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 1d05db463da8586e39e6a8e172529d9da46a8cec11ca36f5231cb9c3f1c49033
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084294"
---
# <a name="add-data-validation-to-excel-ranges"></a>Ajout de validation des données à des plages Excel

La bibliothèque JavaScript Excel fournit des API pour autoriser votre complément à ajouter la validation automatique des données aux tableaux, colonnes, lignes et autres plages dans un classeur. Pour comprendre les concepts et la terminologie de validation des données, consultez les articles suivants sur la façon dont les utilisateurs ajoutent la validation des données via l’interface Excel utilisateur.

- [Application d’une validation des données aux cellules](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Informations supplémentaires sur la validation des données](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Description et exemples de validation des données dans Excel](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Contrôle par programme de validation des données

La propriété `Range.dataValidation`, qui récupère un objet [DataValidation](/javascript/api/excel/excel.datavalidation), constitue le point d’entrée pour le contrôle par programmation de la validation des données dans Excel. Il existe cinq propriétés pour l’objet `DataValidation` :

- `rule` &#8212;Définit ce qui constitue des données valides pour la plage. Voir [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).
- `errorAlert` &#8212;Spécifie si une erreur s’affiche si l’utilisateur entre des données non valides et définit le texte de l’alerte, le titre et le style ; par exemple, **Information**, **Avertissement**, et **Stop**. Voir [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).
- `prompt` &#8212;Spécifie si une demande s’affiche lorsque l’utilisateur pointe sur la plage et définit le message. Voir [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).
- `ignoreBlanks` &#8212;Spécifie si la règle de validation des données s’applique à des cellules vides dans la plage. Par défaut `true`.
- `type` &#8212;Identification accessible en lecture seule du type de validation, par exemple, WholeNumber, Date, TextLength, etc. Elle est définie indirectement lorsque vous définissez la propriété `rule`.

> [!NOTE]
> La validation des données ajoutées par programme se comporte comme celle ajoutée manuellement. Notez que la validation des données est déclenchée uniquement si l’utilisateur entre une valeur dans une cellule ou copie directement et colle une cellule à partir d’un autre emplacement dans le classeur en choisissant l’option de collage **valeurs**. Si l’utilisateur copie une cellule et effectue un simple coller dans une plage avec validation des données, la validation n’est pas déclenchée.

## <a name="creating-validation-rules"></a>Créer les règles de validation

Pour ajouter la validation des données à une plage, votre code doit définir la propriété `rule` de l’objet `DataValidation` dans `Range.dataValidation`. Cela saisit un objet [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) contenant les sept propriétés facultatives. *Une seule de ces propriétés peut être présente dans un objet `DataValidationRule`.* La propriété que vous incluez détermine le type de validation.

### <a name="basic-and-datetime-validation-rule-types"></a>Règles de validation Basic et DateTime

Les trois premières propriétés `DataValidationRule` (c'est-à-dire les types de règles de validation) prennent un objet [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) comme valeur.

- `wholeNumber` &#8212;Nécessite un nombre entier en plus de toute autre validation spécifiée par l’objet `BasicDataValidation`.
- `decimal` &#8212;Nécessite un nombre décimal en plus de toute autre validation spécifiée par l’objet `BasicDataValidation`.
- `textLength` &#8212;Applique les détails de validation dans `BasicDataValidation` l’objet à la *longueur* de valeur de la cellule.

Voici un exemple de création d’une règle de validation. Notez ce qui suit à propos de ce code.

- Le `operator` est l’opérateur binaire « GreaterThan ». Chaque fois que vous utilisez un opérateur binaire, la valeur que l’utilisateur essaie d’entrer dans la cellule est l’opérande gauche et la valeur spécifiée dans `formula1` est l’opérande droite. Par conséquent cette règle indique qu’uniquement les nombres entiers supérieurs à 0 sont valides.
- Le `formula1` est un nombre codé en dur. Lors de la création du code, si vous ne savez pas quelle valeur indiquer, vous pouvez également utiliser une formule Excel (comme chaîne) pour la valeur. Par exemple, « = A3 » et « = SUM(A4,B5) » peuvent également être des valeurs de `formula1`.

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

Voir [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) pour obtenir la liste des autres opérateurs binaires. 

Il existe également deux opérateurs ternaires : « Between » et « NotBetween ». Pour les utiliser, vous devez spécifier la propriété `formula2` facultative. Les valeurs`formula1` et `formula2` sont les opérandes englobantes. La valeur que l’utilisateur essaie d’entrer dans la cellule est la troisième opérande (évaluée). Voici un exemple d’utilisation de l’opérateur « Between ».

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

Les deux propriétés de règle suivantes prennent un objet[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) comme valeur.

- `date`
- `time`

La structure de l’objet `DateTimeDataValidation` est similaire à celle de `BasicDataValidation` : ce dernier a les propriétés `formula1`, `formula2`, et `operator`. Il est aussi utilisé de la même façon. La différence est que vous ne pouvez pas utiliser un nombre dans les propriétés de formule, mais vous pouvez entrer une chaîne [8606 ISO datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou une formule Excel). Voici un exemple qui définit les valeurs valides comme des dates dans la première semaine d’avril 2018. 

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

Utilisez la propriété `list` dans l’objet `DataValidationRule` pour spécifier que les seules valeurs valides sont celles d’une liste de remise. Voici un exemple. Notez ce qui suit à propos de ce code.

- Il part du principe qu’il existe une feuille de calcul nommée « Noms » et que les valeurs dans la plage « A1:A3 » sont des noms.
- La propriété `source` indique la liste des valeurs valides. L’argument de chaîne fait référence à une plage de cellules contenant les noms. Vous pouvez également affecter une liste délimitée par des virgules ; par exemple : « Sue, Ricky, Florence ».
- La propriété `inCellDropDown` indique si un contrôle de liste déroulante s’affiche dans la cellule lorsque l’utilisateur la sélectionne. Si elle est définie sur `true`, alors la flèche déroulante s’affiche avec la liste des valeurs de `source`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: "=Names!$A$1:$A$3"
        }
    };

    return context.sync();
})
```

### <a name="custom-validation-rule-type"></a>Type de règle de validation personnalisée

Utilisez la propriété `custom` dans l’objet `DataValidationRule` pour spécifier une formule de validation personnalisée. Voici un exemple. Notez ce qui suit à propos de ce code.

- Il part du principe qu’il existe un tableau de deux colonnes avec des colonnes **nom athlète** et **commentaires** dans les colonnes A et B de la feuille de calcul.
- Pour réduire le niveau de détail dans la colonne **commentaires**, il rend les données qui incluent le nom de l’athlète invalides.
- `SEARCH(A2,B2)` renvoie la position de départ, dans la chaîne dans B2, de la chaîne dans A2. Si A2 n’est pas contenue dans B2, elle ne renvoie pas de nombre. `ISNUMBER()` renvoie une valeur booléenne. La propriété `formula` indique que les données valides pour la colonne **commentaires** sont des données qui n’incluent pas la chaîne dans la colonne **nom athlète**.

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

## <a name="create-validation-error-alerts"></a>Créer des alertes d’erreur de validation

Vous pouvez créer une alerte d’erreur personnalisée qui s’affiche lorsqu’un utilisateur tente d’entrer des données non valides dans une cellule. Voici un exemple simple. Notez ce qui suit à propos de ce code.

- La propriété `style` détermine si l’utilisateur reçoit une alerte d’information, un avertissement ou une alerte « Stop ». Seule l'alerte `Stop` empêche l’utilisateur d’ajouter des données non valides. La fenêtre contextuelle pour `Warning` et `Information` affiche des options qui autorisent l’utilisateur à entrer tout de même les données non valides.
- La propriété `showAlert` est définie par défaut sur `true`. Cela signifie que Excel une alerte générique (de type) s’ouvre, sauf si vous créez une alerte personnalisée qui définit ou définit un message, un titre et un `Stop` `showAlert` style `false` personnalisés. Ce code définit un message et un titre personnalisés.

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

Pour plus d’informations, voir [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).

## <a name="create-validation-prompts"></a>Créer des demandes de validation

Vous pouvez créer une invite de commandes instructive qui s’affiche lorsqu’un utilisateur survole ou sélectionne une cellule à laquelle la validation des données a été appliquée. Voici un exemple.

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

Pour plus d’informations, voir [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).

## <a name="remove-data-validation-from-a-range"></a>Supprimer la validation des données d’une plage

Pour supprimer la validation des données d’une plage, appelez la méthode [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear__).

```js
myrange.dataValidation.clear()
```

La plage que vous désactivez ne sera pas nécessairement exactement la même plage qu’une plage dans laquelle vous avez ajouté la validation des données. Si ce n’est pas le cas, uniquement les cellules des deux plages qui se chevauchent, le cas échéant, sont effacées. 

> [!NOTE]
> La désactivation de la validation des données à partir d’une plage efface également une validation des données qu’un utilisateur a ajoutée manuellement à la plage.

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Objet DataValidation (API JavaScript pour Excel)](/javascript/api/excel/excel.datavalidation)
- [Objet de plage (API JavaScript pour Excel)](/javascript/api/excel/excel.range)
