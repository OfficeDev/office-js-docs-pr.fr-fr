---
title: Ajout de validation des données à des plages Excel
description: Découvrez comment les API JavaScript pour Excel permettent à votre complément d’ajouter une validation automatique des données aux tableaux, colonnes, lignes et autres plages d’un classeur.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 54ac86def46a130b8b95876a3c42ef8704f9549c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609614"
---
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="d3acd-103">Ajout de validation des données à des plages Excel</span><span class="sxs-lookup"><span data-stu-id="d3acd-103">Add data validation to Excel ranges</span></span>

<span data-ttu-id="d3acd-104">La bibliothèque JavaScript Excel fournit des API pour autoriser votre complément à ajouter la validation automatique des données aux tableaux, colonnes, lignes et autres plages dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="d3acd-104">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="d3acd-105">Pour mieux comprendre les concepts et la terminologie de validation des données, consultez les articles suivants sur la manière dont les utilisateurs ajoutent la validation des données via l’interface utilisateur Excel :</span><span class="sxs-lookup"><span data-stu-id="d3acd-105">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="d3acd-106">Application d’une validation des données aux cellules</span><span class="sxs-lookup"><span data-stu-id="d3acd-106">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="d3acd-107">Informations supplémentaires sur la validation des données</span><span class="sxs-lookup"><span data-stu-id="d3acd-107">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="d3acd-108">Description et exemples de validation des données dans Excel</span><span class="sxs-lookup"><span data-stu-id="d3acd-108">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="d3acd-109">Contrôle par programme de validation des données</span><span class="sxs-lookup"><span data-stu-id="d3acd-109">Programmatic control of data validation</span></span>

<span data-ttu-id="d3acd-110">La propriété `Range.dataValidation`, qui récupère un objet [DataValidation](/javascript/api/excel/excel.datavalidation), constitue le point d’entrée pour le contrôle par programmation de la validation des données dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d3acd-110">The `Range.dataValidation` property, which takes a [DataValidation](/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="d3acd-111">Il existe cinq propriétés pour l’objet `DataValidation` :</span><span class="sxs-lookup"><span data-stu-id="d3acd-111">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="d3acd-112">`rule` &#8212;Définit ce qui constitue des données valides pour la plage.</span><span class="sxs-lookup"><span data-stu-id="d3acd-112">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="d3acd-113">Voir [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="d3acd-113">See [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="d3acd-114">`errorAlert` &#8212;Spécifie si une erreur s’affiche si l’utilisateur entre des données non valides et définit le texte de l’alerte, le titre et le style ; par exemple, **Information**, **Avertissement**, et **Stop**.</span><span class="sxs-lookup"><span data-stu-id="d3acd-114">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="d3acd-115">Voir [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="d3acd-115">See [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="d3acd-116">`prompt` &#8212;Spécifie si une demande s’affiche lorsque l’utilisateur pointe sur la plage et définit le message.</span><span class="sxs-lookup"><span data-stu-id="d3acd-116">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="d3acd-117">Voir [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="d3acd-117">See [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="d3acd-118">`ignoreBlanks` &#8212;Spécifie si la règle de validation des données s’applique à des cellules vides dans la plage.</span><span class="sxs-lookup"><span data-stu-id="d3acd-118">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="d3acd-119">Par défaut `true`.</span><span class="sxs-lookup"><span data-stu-id="d3acd-119">Defaults to `true`.</span></span>
- <span data-ttu-id="d3acd-120">`type` &#8212;Identification accessible en lecture seule du type de validation, par exemple, WholeNumber, Date, TextLength, etc. Elle est définie indirectement lorsque vous définissez la propriété `rule`.</span><span class="sxs-lookup"><span data-stu-id="d3acd-120">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="d3acd-121">La validation des données ajoutées par programme se comporte comme celle ajoutée manuellement.</span><span class="sxs-lookup"><span data-stu-id="d3acd-121">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="d3acd-122">Notez que la validation des données est déclenchée uniquement si l’utilisateur entre une valeur dans une cellule ou copie directement et colle une cellule à partir d’un autre emplacement dans le classeur en choisissant l’option de collage **valeurs**.</span><span class="sxs-lookup"><span data-stu-id="d3acd-122">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="d3acd-123">Si l’utilisateur copie une cellule et effectue un simple coller dans une plage avec validation des données, la validation n’est pas déclenchée.</span><span class="sxs-lookup"><span data-stu-id="d3acd-123">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="d3acd-124">Créer les règles de validation</span><span class="sxs-lookup"><span data-stu-id="d3acd-124">Creating validation rules</span></span>

<span data-ttu-id="d3acd-125">Pour ajouter la validation des données à une plage, votre code doit définir la propriété `rule` de l’objet `DataValidation` dans `Range.dataValidation`.</span><span class="sxs-lookup"><span data-stu-id="d3acd-125">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="d3acd-126">Cela saisit un objet [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) contenant les sept propriétés facultatives.</span><span class="sxs-lookup"><span data-stu-id="d3acd-126">This takes a [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="d3acd-127">*Une seule de ces propriétés peut être présente dans un objet `DataValidationRule`.*</span><span class="sxs-lookup"><span data-stu-id="d3acd-127">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="d3acd-128">La propriété que vous incluez détermine le type de validation.</span><span class="sxs-lookup"><span data-stu-id="d3acd-128">The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="d3acd-129">Règles de validation Basic et DateTime</span><span class="sxs-lookup"><span data-stu-id="d3acd-129">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="d3acd-130">Les trois premières propriétés `DataValidationRule` (c'est-à-dire les types de règles de validation) prennent un objet [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) comme valeur.</span><span class="sxs-lookup"><span data-stu-id="d3acd-130">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="d3acd-131">`wholeNumber` &#8212;Nécessite un nombre entier en plus de toute autre validation spécifiée par l’objet `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="d3acd-131">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="d3acd-132">`decimal` &#8212;Nécessite un nombre décimal en plus de toute autre validation spécifiée par l’objet `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="d3acd-132">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="d3acd-133">`textLength` &#8212;Applique les détails de validation dans `BasicDataValidation` l’objet à la *longueur* de valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="d3acd-133">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="d3acd-134">Voici un exemple de création d’une règle de validation.</span><span class="sxs-lookup"><span data-stu-id="d3acd-134">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="d3acd-135">Voici quelques caractéristiques notables de ce code :</span><span class="sxs-lookup"><span data-stu-id="d3acd-135">Note the following about this code:</span></span>

- <span data-ttu-id="d3acd-136">Le `operator` est l’opérateur binaire « GreaterThan ».</span><span class="sxs-lookup"><span data-stu-id="d3acd-136">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="d3acd-137">Chaque fois que vous utilisez un opérateur binaire, la valeur que l’utilisateur essaie d’entrer dans la cellule est l’opérande gauche et la valeur spécifiée dans `formula1` est l’opérande droite.</span><span class="sxs-lookup"><span data-stu-id="d3acd-137">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="d3acd-138">Par conséquent cette règle indique qu’uniquement les nombres entiers supérieurs à 0 sont valides.</span><span class="sxs-lookup"><span data-stu-id="d3acd-138">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="d3acd-139">Le `formula1` est un nombre codé en dur.</span><span class="sxs-lookup"><span data-stu-id="d3acd-139">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="d3acd-140">Lors de la création du code, si vous ne savez pas quelle valeur indiquer, vous pouvez également utiliser une formule Excel (comme chaîne) pour la valeur.</span><span class="sxs-lookup"><span data-stu-id="d3acd-140">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="d3acd-141">Par exemple, « = A3 » et « = SUM(A4,B5) » peuvent également être des valeurs de `formula1`.</span><span class="sxs-lookup"><span data-stu-id="d3acd-141">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="d3acd-142">Voir [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) pour obtenir la liste des autres opérateurs binaires.</span><span class="sxs-lookup"><span data-stu-id="d3acd-142">See [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="d3acd-143">Il existe également deux opérateurs ternaires : « Between » et « NotBetween ».</span><span class="sxs-lookup"><span data-stu-id="d3acd-143">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="d3acd-144">Pour les utiliser, vous devez spécifier la propriété `formula2` facultative.</span><span class="sxs-lookup"><span data-stu-id="d3acd-144">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="d3acd-145">Les valeurs`formula1` et `formula2` sont les opérandes englobantes.</span><span class="sxs-lookup"><span data-stu-id="d3acd-145">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="d3acd-146">La valeur que l’utilisateur essaie d’entrer dans la cellule est la troisième opérande (évaluée).</span><span class="sxs-lookup"><span data-stu-id="d3acd-146">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="d3acd-147">Voici un exemple d’utilisation de l’opérateur « Between » :</span><span class="sxs-lookup"><span data-stu-id="d3acd-147">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="d3acd-148">Les deux propriétés de règle suivantes prennent un objet[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) comme valeur.</span><span class="sxs-lookup"><span data-stu-id="d3acd-148">The next two rule properties take a [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="d3acd-149">La structure de l’objet `DateTimeDataValidation` est similaire à celle de `BasicDataValidation` : ce dernier a les propriétés `formula1`, `formula2`, et `operator`. Il est aussi utilisé de la même façon.</span><span class="sxs-lookup"><span data-stu-id="d3acd-149">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="d3acd-150">La différence est que vous ne pouvez pas utiliser un nombre dans les propriétés de formule, mais vous pouvez entrer une chaîne [8606 ISO datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou une formule Excel).</span><span class="sxs-lookup"><span data-stu-id="d3acd-150">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="d3acd-151">Voici un exemple qui définit les valeurs valides comme des dates dans la première semaine d’avril 2018.</span><span class="sxs-lookup"><span data-stu-id="d3acd-151">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="d3acd-152">Type de règle de validation de liste</span><span class="sxs-lookup"><span data-stu-id="d3acd-152">List validation rule type</span></span>

<span data-ttu-id="d3acd-153">Utilisez la propriété `list` dans l’objet `DataValidationRule` pour spécifier que les seules valeurs valides sont celles d’une liste de remise.</span><span class="sxs-lookup"><span data-stu-id="d3acd-153">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="d3acd-154">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="d3acd-154">The following is an example.</span></span> <span data-ttu-id="d3acd-155">Voici quelques caractéristiques notables de ce code :</span><span class="sxs-lookup"><span data-stu-id="d3acd-155">Note the following about this code:</span></span>

- <span data-ttu-id="d3acd-156">Il part du principe qu’il existe une feuille de calcul nommée « Noms » et que les valeurs dans la plage « A1:A3 » sont des noms.</span><span class="sxs-lookup"><span data-stu-id="d3acd-156">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="d3acd-157">La propriété `source` indique la liste des valeurs valides.</span><span class="sxs-lookup"><span data-stu-id="d3acd-157">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="d3acd-158">L’argument de chaîne fait référence à une plage de cellules contenant les noms.</span><span class="sxs-lookup"><span data-stu-id="d3acd-158">The string argument refers to a range containing the names.</span></span> <span data-ttu-id="d3acd-159">Vous pouvez également affecter une liste délimitée par des virgules ; par exemple : « Sue, Ricky, Florence ».</span><span class="sxs-lookup"><span data-stu-id="d3acd-159">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="d3acd-160">La propriété `inCellDropDown` indique si un contrôle de liste déroulante s’affiche dans la cellule lorsque l’utilisateur la sélectionne.</span><span class="sxs-lookup"><span data-stu-id="d3acd-160">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="d3acd-161">Si elle est définie sur `true`, alors la flèche déroulante s’affiche avec la liste des valeurs de `source`.</span><span class="sxs-lookup"><span data-stu-id="d3acd-161">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

### <a name="custom-validation-rule-type"></a><span data-ttu-id="d3acd-162">Type de règle de validation personnalisée</span><span class="sxs-lookup"><span data-stu-id="d3acd-162">Custom validation rule type</span></span>

<span data-ttu-id="d3acd-163">Utilisez la propriété `custom` dans l’objet `DataValidationRule` pour spécifier une formule de validation personnalisée.</span><span class="sxs-lookup"><span data-stu-id="d3acd-163">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="d3acd-164">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="d3acd-164">The following is an example.</span></span> <span data-ttu-id="d3acd-165">Voici quelques caractéristiques notables de ce code :</span><span class="sxs-lookup"><span data-stu-id="d3acd-165">Note the following about this code:</span></span>

- <span data-ttu-id="d3acd-166">Il part du principe qu’il existe un tableau de deux colonnes avec des colonnes **nom athlète** et **commentaires** dans les colonnes A et B de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="d3acd-166">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="d3acd-167">Pour réduire le niveau de détail dans la colonne**commentaires**, il rend les données qui incluent le nom de l’athlète invalides.</span><span class="sxs-lookup"><span data-stu-id="d3acd-167">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="d3acd-168">`SEARCH(A2,B2)` renvoie la position de départ, dans la chaîne dans B2, de la chaîne dans A2.</span><span class="sxs-lookup"><span data-stu-id="d3acd-168">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="d3acd-169">Si A2 n’est pas contenue dans B2, elle ne renvoie pas de nombre.</span><span class="sxs-lookup"><span data-stu-id="d3acd-169">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="d3acd-170">`ISNUMBER()` renvoie une valeur booléenne.</span><span class="sxs-lookup"><span data-stu-id="d3acd-170">`ISNUMBER()` returns a boolean.</span></span> <span data-ttu-id="d3acd-171">La propriété `formula` indique que les données valides pour la colonne **commentaires** sont des données qui n’incluent pas la chaîne dans la colonne **nom athlète**.</span><span class="sxs-lookup"><span data-stu-id="d3acd-171">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="d3acd-172">Créer des alertes d’erreur de validation</span><span class="sxs-lookup"><span data-stu-id="d3acd-172">Create validation error alerts</span></span>

<span data-ttu-id="d3acd-173">Vous pouvez créer une alerte d’erreur personnalisée qui s’affiche lorsqu’un utilisateur tente d’entrer des données non valides dans une cellule.</span><span class="sxs-lookup"><span data-stu-id="d3acd-173">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="d3acd-174">Voici un exemple simple.</span><span class="sxs-lookup"><span data-stu-id="d3acd-174">The following is a simple example.</span></span> <span data-ttu-id="d3acd-175">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="d3acd-175">Note the following about this code:</span></span>

- <span data-ttu-id="d3acd-176">La propriété `style` détermine si l’utilisateur reçoit une alerte d’information, un avertissement ou une alerte « Stop ».</span><span class="sxs-lookup"><span data-stu-id="d3acd-176">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="d3acd-177">Seule l'alerte `Stop` empêche l’utilisateur d’ajouter des données non valides.</span><span class="sxs-lookup"><span data-stu-id="d3acd-177">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="d3acd-178">La fenêtre contextuelle pour `Warning` et `Information` affiche des options qui autorisent l’utilisateur à entrer tout de même les données non valides.</span><span class="sxs-lookup"><span data-stu-id="d3acd-178">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="d3acd-179">La propriété `showAlert` est définie par défaut sur `true`.</span><span class="sxs-lookup"><span data-stu-id="d3acd-179">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="d3acd-180">Cela signifie que l’hôte Excel affichera une fenêtre contextuelle d’alerte générique (de type `Stop`), sauf si vous créez une alerte personnalisée qui définit `showAlert` à `false` ou un message, titre et style personnalisés.</span><span class="sxs-lookup"><span data-stu-id="d3acd-180">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="d3acd-181">Ce code définit un message et un titre personnalisés.</span><span class="sxs-lookup"><span data-stu-id="d3acd-181">This code sets a custom message and title.</span></span>

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

<span data-ttu-id="d3acd-182">Pour plus d’informations, voir [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="d3acd-182">For more information, see [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="d3acd-183">Créer des demandes de validation</span><span class="sxs-lookup"><span data-stu-id="d3acd-183">Create validation prompts</span></span>

<span data-ttu-id="d3acd-184">Vous pouvez créer une invite de commandes instructive qui s’affiche lorsqu’un utilisateur survole ou sélectionne une cellule à laquelle la validation des données a été appliquée.</span><span class="sxs-lookup"><span data-stu-id="d3acd-184">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="d3acd-185">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="d3acd-185">The following is an example:</span></span>

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

<span data-ttu-id="d3acd-186">Pour plus d’informations, voir [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="d3acd-186">For more information, see [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="d3acd-187">Supprimer la validation des données d’une plage</span><span class="sxs-lookup"><span data-stu-id="d3acd-187">Remove data validation from a range</span></span>

<span data-ttu-id="d3acd-188">Pour supprimer la validation des données d’une plage, appelez la méthode [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--).</span><span class="sxs-lookup"><span data-stu-id="d3acd-188">To remove data validation from a range, call the  [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="d3acd-189">La plage que vous désactivez ne sera pas nécessairement exactement la même plage qu’une plage dans laquelle vous avez ajouté la validation des données.</span><span class="sxs-lookup"><span data-stu-id="d3acd-189">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="d3acd-190">Si ce n’est pas le cas, uniquement les cellules des deux plages qui se chevauchent, le cas échéant, sont effacées.</span><span class="sxs-lookup"><span data-stu-id="d3acd-190">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="d3acd-191">La désactivation de la validation des données à partir d’une plage efface également une validation des données qu’un utilisateur a ajoutée manuellement à la plage.</span><span class="sxs-lookup"><span data-stu-id="d3acd-191">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="d3acd-192">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d3acd-192">See also</span></span>

- [<span data-ttu-id="d3acd-193">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="d3acd-193">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d3acd-194">Objet DataValidation (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="d3acd-194">DataValidation Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="d3acd-195">Objet de plage (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="d3acd-195">Range Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.range)
