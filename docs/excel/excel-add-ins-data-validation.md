---
title: Ajouter la validation des donn?es aux plages Excel
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: 8e5f09f1c566103f34ad584885769229c17ab1f7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="add-data-validation-to-excel-ranges-preview"></a>Ajouter la validation des donn?es aux plages Excel (pr?version)

> [!NOTE]
> Tant que les API de validation des donn?es sont en pr?version, vous devez charger la version b?ta de la biblioth?que JavaScript Office pour les utiliser. L'URL est https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Si vous utilisez TypeScript ou si votre ?diteur de code utilise un fichier de d?finition de type TypeScript pour intelliSense, utilisez https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

La biblioth?que JavaScript Excel fournit des API pour permettre ? votre compl?ment d'ajouter une validation automatique des donn?es aux tables, colonnes, lignes et autres plages d'un classeur. Pour comprendre les concepts et la terminologie de la validation des donn?es, consultez les articles suivants qui portent sur la mani?re dont les utilisateurs peuvent ajouter la validation des donn?es via l'IU Excel?:

- [Appliquer la validation des donn?es aux cellules](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Plus d'informations sur la validation des donn?es](https://microsoft.sharepoint.com/:p:/r/teams/oext/_layouts/15/Doc.aspx?sourcedoc=%7B51143964-d52c-429d-bfac-c7495473d536%7D&action=edit)
- [Description et exemples de validation de donn?es dans Excel](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Contr?le programmatique de la validation des donn?es

La propri?t? `Range.dataValidation`, qui prend un objet de validation de donn?es [,](https://dev.office.com/reference/add-ins/excel/datavalidation) est le point d'entr?e pour le contr?le programmatique de la validation des donn?es dans Excel. Il y a cinq propri?t?s ? l'objet `DataValidation`?:

- `rule` - D?finit ce qui constitue des donn?es valides pour la plage. Voir [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).
- `errorAlert` - Sp?cifie si une erreur appara?t lorsque l'utilisateur entre des donn?es non valides et d?finit le texte, le titre et le style d'alerte?; par exemple, **Informatif**, **Avertissement**, et **Arr?ter**. Voir [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).
- `prompt` - Indique si une invite s'affiche lorsque l'utilisateur survole la plage et d?finit le message d'assistance vocale. Voir [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).
- `ignoreBlanks` - Sp?cifie si la r?gle de validation des donn?es s'applique aux cellules vides de la plage. Par d?faut `true`.
- `type` - Une identification en lecture seule du type de validation, tel que WholeNumber, Date, TextLength, etc. Il est d?fini indirectement lorsque vous d?finissez la propri?t? `rule`.

> [!NOTE]
> La validation des donn?es ajout?e de fa?on programmatique se comporte exactement comme la validation des donn?es ajout?e manuellement. Surtout, notez que la validation des donn?es est d?clench?e uniquement si l'utilisateur entre directement une valeur dans une cellule ou copie et colle une cellule d'un autre emplacement dans le classeur et choisit l'option de collage, valeurs ****. Si l'utilisateur copie une cellule et la colle simplement dans une plage avec validation des donn?es, la validation n'est pas d?clench?e.

### <a name="creating-validation-rules"></a>Cr?er des r?gles de validation

Pour ajouter une validation de donn?es ? une plage, votre code doit d?finir la propri?t? `rule` de l'objet `DataValidation` dans `Range.dataValidation`. Cela prend un objet [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) qui a sept propri?t?s facultatives. *Une seule de ces propri?t?s peut ?tre pr?sente dans un objet `DataValidationRule`.* La propri?t? que vous incluez d?termine le type de validation.

#### <a name="basic-and-datetime-validation-rule-types"></a>Types de r?gles de validation de base et DateTime

Les trois premi?res propri?t?s `DataValidationRule` (c.-?-d. les types de r?gles de validation) prennent un objet [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) comme leur valeur.

- `wholeNumber` ? N?cessite un nombre entier en plus de toute autre validation sp?cifi?e par l'objet `BasicDataValidation`.
- `decimal` - N?cessite un nombre d?cimal en plus de toute autre validation sp?cifi?e par l'objet `BasicDataValidation`.
- `textLength` ? Applique les d?tails de validation dans l'objet `BasicDataValidation` ? la *longueur*  de la valeur de la cellule.

Voici un exemple de cr?ation d'une r?gle de validation. Tenez compte des informations suivantes concernant ce code?:

- Le `operator` est l'op?rateur binaire "sup?rieur ?". Chaque fois que vous utilisez un op?rateur binaire, la valeur que l'utilisateur essaie d'entrer dans la cellule est l'op?rande de gauche et la valeur sp?cifi?e dans `formula1` est l'op?rande de droite. Donc, cette r?gle dit que seuls les nombres entiers sup?rieurs ? 0 sont valides. 
- Le `formula1` est un nombre cod? en dur. Si vous ne savez pas au moment du codage quelle devrait ?tre la valeur, vous pouvez ?galement utiliser une formule Excel (sous forme de cha?ne) pour la valeur. Par exemple, ??= A3?? et ??= SOMME (A4:B5)?? peuvent ?galement ?tre des valeurs de `formula1`.

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

Voir [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) pour une liste des autres op?rateurs binaires. 

Il existe aussi deux op?rateurs ternaires?: ??Between?? et ??NotBetween??. Pour les utiliser, vous devez sp?cifier la propri?t? `formula2` facultative. Les valeurs `formula1` et `formula2` sont les op?randes de d?limitation. La valeur que l'utilisateur essaie d'entrer dans la cellule est le troisi?me op?rande (?valu?e). Voici un exemple d'utilisation de l'op?rateur ??Between???:

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

Les deux propri?t?s de r?gle suivantes prennent un objet [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) comme leur valeur.

- `date`
- `time`

L'objet `DateTimeDataValidation` est structur? de mani?re similaire ? la `BasicDataValidation`?: il a les propri?t?s `formula1`, `formula2`, et `operator`, et est utilis? de la m?me mani?re. La diff?rence est que vous ne pouvez pas utiliser un nombre dans les propri?t?s de la formule, mais vous pouvez entrer une cha?ne [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou une formule Excel). Voici un exemple qui d?finit des valeurs valides telles que des dates dans la premi?re semaine d'avril 2018. 

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

#### <a name="list-validation-rule-type"></a>Type de r?gle de validation de liste

Utilisez la propri?t? `list` dans l'objet `DataValidationRule` pour sp?cifier que les seules valeurs valides sont celles d'une liste finie. Voir l'exemple suivant. Tenez compte des informations suivantes concernant ce code?:

- Il suppose qu'il existe une feuille de calcul nomm?e "Noms" et que les valeurs de la plage "A1: A3" sont des noms.
- La propri?t? `source` sp?cifie la liste des valeurs valides. La plage avec les noms lui a ?t? affect?e. Vous pouvez ?galement affecter une liste d?limit?e par des virgules, comme par exemple?: ??Sue, Ricky, Liz??. 
- La propri?t? `inCellDropDown` sp?cifie si un contr?le d?roulant appara?tra dans la cellule lorsque l'utilisateur le s?lectionne. Si elle est d?finie sur `true`, une liste d?roulante appara?t contenant la liste des valeurs du `source`.

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

#### <a name="custom-validation-rule-type"></a>Type de r?gle de validation personnalis?e

Utilisez la propri?t? `custom` dans l'objet `DataValidationRule` pour sp?cifier une formule de validation personnalis?e. Voir l'exemple suivant. Tenez compte des informations suivantes concernant ce code?:

- Il suppose qu'il y a un tableau ? deux colonnes avec des colonnes **Nom de l'athl?te** et **Commentaires** dans les colonnes A et B de la feuille de calcul.
- Pour r?duire la verbosit? dans la colonne **Commentaires,** il rend invalides les donn?es qui incluent le nom de l'athl?te.
- `SEARCH(A2,B2)` renvoie la position de d?part, de la cha?ne dans B2, de la cha?ne dans A2. Si A2 n'est pas contenu dans B2, il ne renvoie pas de nombre. `ISNUMBER()` retourne un bool?en. La propri?t? `formula` indique donc que les donn?es valides pour la colonne **Commentaire** sont les donn?es qui n'incluent pas la cha?ne pr?sente dans la colonne **Nom de l'athl?te**.

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

### <a name="create-validation-error-alerts"></a>Cr?er des alertes d'erreur de validation

Vous pouvez cr?er une alerte d'erreur personnalis?e qui appara?t lorsqu'un utilisateur tente d'entrer des donn?es non valides dans une cellule. Ce qui suit est un simple exemple. Tenez compte des informations suivantes concernant ce code?:

- La propri?t? `style` d?termine si l'utilisateur re?oit une alerte informative, un avertissement ou une alerte d' "arr?t". Seule `Stop` emp?che r?ellement l'utilisateur d'ajouter des donn?es invalides. La fen?tre contextuelle pour `Warning` et `Information` a des options qui permettent ? l'utilisateur d'entrer les donn?es invalides de toute fa?on.
- La propri?t? `showAlert` prend `true` par d?faut. Cela signifie que l'h?te Excel affichera une alerte g?n?rique (de type `Stop`) sauf si vous cr?ez une alerte personnalis?e qui soit d?finit `showAlert` pour `false` ou d?finit un message, un titre et un style personnalis?s. Ce code d?finit un message et un titre personnalis?s.


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

### <a name="create-validation-prompts"></a>Cr?er des invites de validation

Vous pouvez cr?er une invite d'instruction qui appara?t lorsqu'un utilisateur survole ou s?lectionne une cellule ? laquelle la validation des donn?es a ?t? appliqu?e. Voici un exemple?:

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

### <a name="remove-data-validation-from-a-range"></a>Supprimer la validation des donn?es d'une plage

Pour supprimer la validation des donn?es d'une plage, appelez la m?thode [Range.dataValidation.clear ()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear).

```js
myrange.dataValidation.clear()
```

Il n'est pas n?cessaire que la plage que vous effacez soit exactement la m?me plage que celle sur laquelle vous avez ajout? la validation des donn?es. Si ce n'est pas le cas, seules les cellules des deux plages qui se chevauchent sont effac?es, le cas ?ch?ant. 

> [!NOTE]
> L'effacement de la validation des donn?es d'une plage efface ?galement toute validation de donn?es qu'un utilisateur a ajout?e manuellement ? la plage.

## <a name="see-also"></a>Voir aussi

- [Concepts de base de l?API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Objet DataValidation (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/datavalidation)
- [Objet Range (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/range)



 
