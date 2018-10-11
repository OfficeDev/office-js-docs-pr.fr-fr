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
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="ee231-102">Ajouter de la validation des données aux plages Excel</span><span class="sxs-lookup"><span data-stu-id="ee231-102">Add data validation to Excel ranges</span></span>

<span data-ttu-id="ee231-p101">La bibliothèque JavaScript Excel fournit des API pour permettre à votre complément d’ajouter de la validation automatique des données à des tableaux, des colonnes, des lignes et autres plages dans un classeur. Pour comprendre les concepts et la terminologie de la validation des données, consultez les articles suivants sur la façon dont les utilisateurs ajoutent de la validation des données par le biais de l’interface utilisateur d’Excel :</span><span class="sxs-lookup"><span data-stu-id="ee231-p101">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook. To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="ee231-105">Appliquer de la validation des données aux cellules</span><span class="sxs-lookup"><span data-stu-id="ee231-105">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="ee231-106">Plus d'informations sur la validation des données</span><span class="sxs-lookup"><span data-stu-id="ee231-106">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="ee231-107">Description et exemples de validation de données dans Excel</span><span class="sxs-lookup"><span data-stu-id="ee231-107">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="ee231-108">Contrôle programmatique de la validation des données</span><span class="sxs-lookup"><span data-stu-id="ee231-108">Programmatic control of data validation</span></span>

<span data-ttu-id="ee231-p102">La propriété `Range.dataValidation`, qui accepte un objet [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) , est le point d’entrée du contrôle programmatique de la validation des données dans Excel. L’objet`DataValidation` a cinq propriétés :</span><span class="sxs-lookup"><span data-stu-id="ee231-p102">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel. There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="ee231-p103">`rule` — définir ce qui constitue des données valides pour la plage. Voir [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="ee231-p103">`rule` &#8212; Defines what constitutes valid data for the range. See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="ee231-p104">`errorAlert` — Spécifie si une erreur apparaît lorsque l’utilisateur entre des données non valides et définit le texte, le titre et le style d’alerte. Par exemple, **Informatif**, **Avertissement**, et **Arrêter**. Voir [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="ee231-p104">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**. See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="ee231-p105">`prompt` — Spécifie si une invite s’affiche lorsque l’utilisateur survole la plage et définir le message d’assistance vocale. Voir [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="ee231-p105">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message. See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="ee231-p106">`ignoreBlanks` — Spécifie si la règle de validation des données s’applique aux cellules vides de la plage. `true` par défaut.</span><span class="sxs-lookup"><span data-stu-id="ee231-p106">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range. Defaults to `true`.</span></span>
- <span data-ttu-id="ee231-119">`type` — une identification en lecture seule du type de validation, tel que WholeNumber, Date, TextLength, etc. Elle est définie indirectement lorsque vous définissez la propriété `rule`.</span><span class="sxs-lookup"><span data-stu-id="ee231-119">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="ee231-p107">La validation des données ajoutée par programme se comporte comme celle ajoutée manuellement. Notez en particulier que la validation des données est déclenchée uniquement si l’utilisateur entre directement une valeur dans une cellule ou copie et colle une cellule à partir d’un autre endroit du classeur et choisit l’option de collage **Valeurs**. Si l’utilisateur copie une cellule et effectue un collage brut dans une plage de validation des données, la validation n’est pas déclenchée.</span><span class="sxs-lookup"><span data-stu-id="ee231-p107">Data validation added programmatically behaves just like manually added data validation. In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option. If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="ee231-123">Création des règles de validation</span><span class="sxs-lookup"><span data-stu-id="ee231-123">Creating validation rules</span></span>

<span data-ttu-id="ee231-p108">Pour ajouter de la validation des données à une plage, votre code doit définir la propriété `rule` de l’objet `DataValidation` dans `Range.dataValidation`. Elle accepte un objet [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) qui a sept propriétés facultatives. *Seulement une de ces propriétés peut être présente dans un objet `DataValidationRule`.* La propriété que vous incluez détermine le type de validation.</span><span class="sxs-lookup"><span data-stu-id="ee231-p108">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`. This takes a [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties. *No more than one of these properties may be present in any `DataValidationRule` object.* The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="ee231-128">Types de règles de validation de base et DateTime</span><span class="sxs-lookup"><span data-stu-id="ee231-128">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="ee231-129">Les trois premières propriétés `DataValidationRule` (c.-à-d. les types de règles de validation) prennent un objet [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) comme leur valeur.</span><span class="sxs-lookup"><span data-stu-id="ee231-129">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="ee231-130">`wholeNumber` – Nécessite un nombre entier en plus de toute autre validation spécifiée par l'objet `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="ee231-130">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="ee231-131">`decimal` – Nécessite un nombre décimal en plus de toute autre validation spécifiée par l'objet `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="ee231-131">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="ee231-132">`textLength` – Applique les détails de validation dans l'objet `BasicDataValidation` à la *longueur*  de la valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="ee231-132">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="ee231-p109">Voici un exemple de création d’une règle de validation. Notez ce qui suit concernant ce code :</span><span class="sxs-lookup"><span data-stu-id="ee231-p109">Here is an example of creating a validation rule. Note the following about this code:</span></span>

- <span data-ttu-id="ee231-p110">Le `operator` est l’opérateur binaire « Supérieur ». Chaque fois que vous utilisez un opérateur binaire, la valeur que l’utilisateur essaie d’entrer dans la cellule est l’opérande gauche et la valeur spécifiée dans `formula1` est l’opérande droit. Cette règle indique donc que seuls les nombres entiers qui sont supérieures à 0 sont valides.</span><span class="sxs-lookup"><span data-stu-id="ee231-p110">The `operator` is the binary operator "GreaterThan". Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand. So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="ee231-p111">Le `formula1` est un nombre codé en dur. Si vous ne savez quelle valeur utiliser au moment du codage, vous pouvez également utiliser une formule Excel (en tant que chaîne) pour la valeur. Par exemple, "=A3" et "=SUM(A4,B5)" peuvent également être des valeurs de `formula1`.</span><span class="sxs-lookup"><span data-stu-id="ee231-p111">The `formula1` is a hard-coded number. If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value. For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="ee231-141">Voir [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) pour une liste des autres opérateurs binaires.</span><span class="sxs-lookup"><span data-stu-id="ee231-141">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="ee231-p112">Il y a aussi deux opérateurs ternaires : « Between » et « NotBetween ». Pour les utiliser, vous devez spécifier la propriété optionnelle `formula2`. Les valeurs `formula1` et `formula2` sont les opérandes limites. La valeur que l’utilisateur essaie d’entrer dans la cellule est le troisième opérande (évalué). Voici un exemple d’utilisation de l’opérateur « Between » :</span><span class="sxs-lookup"><span data-stu-id="ee231-p112">There are also two ternary operators: "Between" and "NotBetween". To use these, you must specify the optional `formula2` property. The `formula1` and `formula2` values are the bounding operands. The value that the user tries to enter in the cell is the third (evaluated) operand. The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="ee231-147">Les deux propriétés de règle suivantes prennent un objet [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) comme valeur.</span><span class="sxs-lookup"><span data-stu-id="ee231-147">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="ee231-p113">L’objet `DateTimeDataValidation` est structuré de la même façon que le `BasicDataValidation` : il possède les propriétés `formula1`, `formula2`, et `operator` et s’utilise de manière similaire. La différence est que vous ne pouvez pas utiliser un nombre dans les propriétés de formule, mais vous pouvez entrer une chaîne [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou une formule Excel). Voici un exemple qui définit les valeurs valides comme des dates dans la première semaine d’avril 2018.</span><span class="sxs-lookup"><span data-stu-id="ee231-p113">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way. The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula). The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="ee231-151">Type de règle de validation de liste</span><span class="sxs-lookup"><span data-stu-id="ee231-151">List validation rule type</span></span>

<span data-ttu-id="ee231-p114">Utilisez la propriété `list` dans l’objet `DataValidationRule` pour spécifier que les seules valeurs valides sont celles d’une liste finie. Voici un exemple. Notez ce qui suit concernant ce code :</span><span class="sxs-lookup"><span data-stu-id="ee231-p114">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="ee231-155">Il suppose qu’il existe une feuille de calcul nommée « Names » et que les valeurs de la plage « A1:A3 » sont des noms.</span><span class="sxs-lookup"><span data-stu-id="ee231-155">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="ee231-p115">La propriété `source` définit la liste des valeurs valides. La plage avec les noms été lui a été affectée. Vous pouvez également affecter une liste délimitée par des virgules (par exemple : « Sue, Ricky, Liz »).</span><span class="sxs-lookup"><span data-stu-id="ee231-p115">The `source` property specifies the list of valid values. The range with the names has been assigned to it. You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="ee231-p116">La propriété `inCellDropDown` spécifie si un contrôle de liste déroulante apparaît dans la cellule lorsque l’utilisateur la sélectionne. Si la valeur est définie sur `true`, alors la liste déroulante s’affiche avec la liste de valeurs à partir de `source`.</span><span class="sxs-lookup"><span data-stu-id="ee231-p116">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it. If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

### <a name="custom-validation-rule-type"></a><span data-ttu-id="ee231-161">Type de règle de validation personnalisée</span><span class="sxs-lookup"><span data-stu-id="ee231-161">Custom validation rule type</span></span>

<span data-ttu-id="ee231-p117">Utilisez la propriété `custom` dans l’objet `DataValidationRule` pour spécifier une formule de validation personnalisée. En voici un exemple. Notez ce qui suit à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="ee231-p117">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="ee231-165">Il suppose qu’il y a un tableau à deux colonnes avec des colonnes **Athlete Name** et **Comments** dans les colonnes A et B de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="ee231-165">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="ee231-166">Pour réduire la verbosité dans la colonne **Comments**, il rend invalides les données qui incluent le nom de l’athlète.</span><span class="sxs-lookup"><span data-stu-id="ee231-166">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="ee231-p118">`SEARCH(A2,B2)` Renvoie la position de départ, dans la chaîne dans B2, de la chaîne dans A2. Si A2 n’est pas incluse dans B2, elle ne renvoie pas un nombre. `ISNUMBER()` renvoie une valeur booléenne. La propriété `formula` indique donc que les données valides pour la colonne **Comment** sont les données qui n’incluent pas la chaîne dans la colonne **Athlete Name** .</span><span class="sxs-lookup"><span data-stu-id="ee231-p118">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2. If A2 is not contained in B2, it does not return a number. `ISNUMBER()` returns a boolean. So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="ee231-171">Création des alertes d’erreur de validation</span><span class="sxs-lookup"><span data-stu-id="ee231-171">Create validation error alerts</span></span>

<span data-ttu-id="ee231-p119">Vous pouvez créer une alerte d’erreur personnalisée qui s’affiche lorsqu’un utilisateur essaie d’entrer des données non valides dans une cellule. Voici un exemple simple. Notez ce qui suit concernant ce code :</span><span class="sxs-lookup"><span data-stu-id="ee231-p119">You can a create custom error alert that appears when a user tries to enter invalid data in a cell. The following is a simple example. Note the following about this code:</span></span>

- <span data-ttu-id="ee231-p120">La propriété `style`  détermine si l’utilisateur reçoit une alerte d’information, un avertissement ou une alerte « stop ». Seul `Stop` empêche réellement l’utilisateur d’ajouter des données non valides. La fenêtre contextuelle pour `Warning` et `Information` a des options qui permettent de saisir quand même les données non valides.</span><span class="sxs-lookup"><span data-stu-id="ee231-p120">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert. Only `Stop` actually prevents the user from adding invalid data. The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="ee231-p121">La valeur par défaut de la propriété `showAlert` est `true`. Cela signifie que l’hôte Excel affichera une alerte générique (de type `Stop`) sauf si vous créez une alerte personnalisée qui définit `showAlert` sur `false` ou définit un message personnalisé, un titre et un style. Ce code définit un message personnalisé et un titre.</span><span class="sxs-lookup"><span data-stu-id="ee231-p121">The `showAlert` property defaults to `true`. This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style. This code sets a custom message and title.</span></span>


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

<span data-ttu-id="ee231-181">Pour en savoir plus, voir [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="ee231-181">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="ee231-182">Création des invites de validation</span><span class="sxs-lookup"><span data-stu-id="ee231-182">Create validation prompts</span></span>

<span data-ttu-id="ee231-p122">Vous pouvez créer une invite d’instruction qui apparaît lorsqu’un utilisateur survole ou sélectionne une cellule à laquelle la validation des données a été appliquée. Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="ee231-p122">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied. The following is an example:</span></span>

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

<span data-ttu-id="ee231-185">Pour en savoir plus, voir [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="ee231-185">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="ee231-186">Suppression de la validation des données d’une plage</span><span class="sxs-lookup"><span data-stu-id="ee231-186">Remove data validation from a range</span></span>

<span data-ttu-id="ee231-187">Pour supprimer la validation des données d'une plage, appelez la méthode [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--).</span><span class="sxs-lookup"><span data-stu-id="ee231-187">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="ee231-p123">Il n’est pas nécessaire que la plage que vous désactivez soit exactement la même plage que celle sur laquelle vous avez ajouté la validation des données. Si ce n’est pas le cas, seules les cellules qui se chevauchent dans les deux plages sont désactivées le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="ee231-p123">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation. If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="ee231-190">La désactivation de la validation des données d’une plage efface également toute validation de données qu’un utilisateur a ajoutée manuellement à la plage.</span><span class="sxs-lookup"><span data-stu-id="ee231-190">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="ee231-191">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ee231-191">See also</span></span>

- [<span data-ttu-id="ee231-192">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="ee231-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ee231-193">Objet DataValidation (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="ee231-193">Chart Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="ee231-194">Objet Range (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="ee231-194">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
