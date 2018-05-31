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
# <a name="add-data-validation-to-excel-ranges-preview"></a><span data-ttu-id="766a8-102">Ajouter la validation des données aux plages Excel (préversion)</span><span class="sxs-lookup"><span data-stu-id="766a8-102">Add data validation to Excel ranges (Preview)</span></span>

> [!NOTE]
> <span data-ttu-id="766a8-103">Tant que les API de validation des données sont en préversion, vous devez charger la version bêta de la bibliothèque JavaScript Office pour les utiliser.</span><span class="sxs-lookup"><span data-stu-id="766a8-103">While the data validation APIs are in preview, you must load the beta version of the Office JavaScript library to use them.</span></span> <span data-ttu-id="766a8-104">L'URL est https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="766a8-104">The full URL is https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span> <span data-ttu-id="766a8-105">Si vous utilisez TypeScript ou si votre éditeur de code utilise un fichier de définition de type TypeScript pour intelliSense, utilisez https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="766a8-105">If you are using TypeScript or your code editor uses a TypeScript type definition file for intellisense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="766a8-106">La bibliothèque JavaScript Excel fournit des API pour permettre à votre complément d'ajouter une validation automatique des données aux tables, colonnes, lignes et autres plages d'un classeur.</span><span class="sxs-lookup"><span data-stu-id="766a8-106">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="766a8-107">Pour comprendre les concepts et la terminologie de la validation des données, consultez les articles suivants qui portent sur la manière dont les utilisateurs peuvent ajouter la validation des données via l'IU Excel :</span><span class="sxs-lookup"><span data-stu-id="766a8-107">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="766a8-108">Appliquer la validation des données aux cellules</span><span class="sxs-lookup"><span data-stu-id="766a8-108">Apply data validation to cells</span></span>](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="766a8-109">Plus d'informations sur la validation des données</span><span class="sxs-lookup"><span data-stu-id="766a8-109">More on data validation</span></span>](https://microsoft.sharepoint.com/:p:/r/teams/oext/_layouts/15/Doc.aspx?sourcedoc=%7B51143964-d52c-429d-bfac-c7495473d536%7D&action=edit)
- [<span data-ttu-id="766a8-110">Description et exemples de validation de données dans Excel</span><span class="sxs-lookup"><span data-stu-id="766a8-110">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="766a8-111">Contrôle programmatique de la validation des données</span><span class="sxs-lookup"><span data-stu-id="766a8-111">Programmatic control of data validation</span></span>

<span data-ttu-id="766a8-112">La propriété `Range.dataValidation`, qui prend un objet de validation de données [,](https://dev.office.com/reference/add-ins/excel/datavalidation) est le point d'entrée pour le contrôle programmatique de la validation des données dans Excel.</span><span class="sxs-lookup"><span data-stu-id="766a8-112">The `Range.dataValidation` property, which takes a [DataValidation](https://dev.office.com/reference/add-ins/excel/datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="766a8-113">Il y a cinq propriétés à l'objet `DataValidation` :</span><span class="sxs-lookup"><span data-stu-id="766a8-113">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="766a8-114">`rule` - Définit ce qui constitue des données valides pour la plage.</span><span class="sxs-lookup"><span data-stu-id="766a8-114">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="766a8-115">Voir [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="766a8-115">See [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).</span></span>
- <span data-ttu-id="766a8-116">`errorAlert` - Spécifie si une erreur apparaît lorsque l'utilisateur entre des données non valides et définit le texte, le titre et le style d'alerte ; par exemple, **Informatif**, **Avertissement**, et **Arrêter**.</span><span class="sxs-lookup"><span data-stu-id="766a8-116">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="766a8-117">Voir [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="766a8-117">See [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span></span>
- <span data-ttu-id="766a8-118">`prompt` - Indique si une invite s'affiche lorsque l'utilisateur survole la plage et définit le message d'assistance vocale.</span><span class="sxs-lookup"><span data-stu-id="766a8-118">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="766a8-119">Voir [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="766a8-119">See [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span></span>
- <span data-ttu-id="766a8-120">`ignoreBlanks` - Spécifie si la règle de validation des données s'applique aux cellules vides de la plage.</span><span class="sxs-lookup"><span data-stu-id="766a8-120">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="766a8-121">Par défaut `true`.</span><span class="sxs-lookup"><span data-stu-id="766a8-121">Defaults to `true`.</span></span>
- <span data-ttu-id="766a8-122">`type` - Une identification en lecture seule du type de validation, tel que WholeNumber, Date, TextLength, etc. Il est défini indirectement lorsque vous définissez la propriété `rule`.</span><span class="sxs-lookup"><span data-stu-id="766a8-122">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="766a8-123">La validation des données ajoutée de façon programmatique se comporte exactement comme la validation des données ajoutée manuellement.</span><span class="sxs-lookup"><span data-stu-id="766a8-123">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="766a8-124">Surtout, notez que la validation des données est déclenchée uniquement si l'utilisateur entre directement une valeur dans une cellule ou copie et colle une cellule d'un autre emplacement dans le classeur et choisit l'option de collage, valeurs ****.</span><span class="sxs-lookup"><span data-stu-id="766a8-124">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="766a8-125">Si l'utilisateur copie une cellule et la colle simplement dans une plage avec validation des données, la validation n'est pas déclenchée.</span><span class="sxs-lookup"><span data-stu-id="766a8-125">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

### <a name="creating-validation-rules"></a><span data-ttu-id="766a8-126">Créer des règles de validation</span><span class="sxs-lookup"><span data-stu-id="766a8-126">Creating validation rules</span></span>

<span data-ttu-id="766a8-127">Pour ajouter une validation de données à une plage, votre code doit définir la propriété `rule` de l'objet `DataValidation` dans `Range.dataValidation`.</span><span class="sxs-lookup"><span data-stu-id="766a8-127">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="766a8-128">Cela prend un objet [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) qui a sept propriétés facultatives.</span><span class="sxs-lookup"><span data-stu-id="766a8-128">This takes a [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="766a8-129">*Une seule de ces propriétés peut être présente dans un objet `DataValidationRule`.*</span><span class="sxs-lookup"><span data-stu-id="766a8-129">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="766a8-130">La propriété que vous incluez détermine le type de validation.</span><span class="sxs-lookup"><span data-stu-id="766a8-130">The property that you include determines the type of validation.</span></span>

#### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="766a8-131">Types de règles de validation de base et DateTime</span><span class="sxs-lookup"><span data-stu-id="766a8-131">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="766a8-132">Les trois premières propriétés `DataValidationRule` (c.-à-d. les types de règles de validation) prennent un objet [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) comme leur valeur.</span><span class="sxs-lookup"><span data-stu-id="766a8-132">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="766a8-133">`wholeNumber` – Nécessite un nombre entier en plus de toute autre validation spécifiée par l'objet `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="766a8-133">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="766a8-134">`decimal` - Nécessite un nombre décimal en plus de toute autre validation spécifiée par l'objet `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="766a8-134">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="766a8-135">`textLength` – Applique les détails de validation dans l'objet `BasicDataValidation` à la *longueur*  de la valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="766a8-135">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="766a8-136">Voici un exemple de création d'une règle de validation.</span><span class="sxs-lookup"><span data-stu-id="766a8-136">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="766a8-137">Tenez compte des informations suivantes concernant ce code :</span><span class="sxs-lookup"><span data-stu-id="766a8-137">Note the following about this code:</span></span>

- <span data-ttu-id="766a8-138">Le `operator` est l'opérateur binaire "supérieur à".</span><span class="sxs-lookup"><span data-stu-id="766a8-138">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="766a8-139">Chaque fois que vous utilisez un opérateur binaire, la valeur que l'utilisateur essaie d'entrer dans la cellule est l'opérande de gauche et la valeur spécifiée dans `formula1` est l'opérande de droite.</span><span class="sxs-lookup"><span data-stu-id="766a8-139">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="766a8-140">Donc, cette règle dit que seuls les nombres entiers supérieurs à 0 sont valides.</span><span class="sxs-lookup"><span data-stu-id="766a8-140">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="766a8-141">Le `formula1` est un nombre codé en dur.</span><span class="sxs-lookup"><span data-stu-id="766a8-141">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="766a8-142">Si vous ne savez pas au moment du codage quelle devrait être la valeur, vous pouvez également utiliser une formule Excel (sous forme de chaîne) pour la valeur.</span><span class="sxs-lookup"><span data-stu-id="766a8-142">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="766a8-143">Par exemple, « = A3 » et « = SOMME (A4:B5) » peuvent également être des valeurs de `formula1`.</span><span class="sxs-lookup"><span data-stu-id="766a8-143">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="766a8-144">Voir [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) pour une liste des autres opérateurs binaires.</span><span class="sxs-lookup"><span data-stu-id="766a8-144">See [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="766a8-145">Il existe aussi deux opérateurs ternaires : « Between » et « NotBetween ».</span><span class="sxs-lookup"><span data-stu-id="766a8-145">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="766a8-146">Pour les utiliser, vous devez spécifier la propriété `formula2` facultative.</span><span class="sxs-lookup"><span data-stu-id="766a8-146">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="766a8-147">Les valeurs `formula1` et `formula2` sont les opérandes de délimitation.</span><span class="sxs-lookup"><span data-stu-id="766a8-147">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="766a8-148">La valeur que l'utilisateur essaie d'entrer dans la cellule est le troisième opérande (évaluée).</span><span class="sxs-lookup"><span data-stu-id="766a8-148">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="766a8-149">Voici un exemple d'utilisation de l'opérateur « Between » :</span><span class="sxs-lookup"><span data-stu-id="766a8-149">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="766a8-150">Les deux propriétés de règle suivantes prennent un objet [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) comme leur valeur.</span><span class="sxs-lookup"><span data-stu-id="766a8-150">The next two rule properties take a [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="766a8-151">L'objet `DateTimeDataValidation` est structuré de manière similaire à la `BasicDataValidation` : il a les propriétés `formula1`, `formula2`, et `operator`, et est utilisé de la même manière.</span><span class="sxs-lookup"><span data-stu-id="766a8-151">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="766a8-152">La différence est que vous ne pouvez pas utiliser un nombre dans les propriétés de la formule, mais vous pouvez entrer une chaîne [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou une formule Excel).</span><span class="sxs-lookup"><span data-stu-id="766a8-152">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="766a8-153">Voici un exemple qui définit des valeurs valides telles que des dates dans la première semaine d'avril 2018.</span><span class="sxs-lookup"><span data-stu-id="766a8-153">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

#### <a name="list-validation-rule-type"></a><span data-ttu-id="766a8-154">Type de règle de validation de liste</span><span class="sxs-lookup"><span data-stu-id="766a8-154">List validation rule type</span></span>

<span data-ttu-id="766a8-155">Utilisez la propriété `list` dans l'objet `DataValidationRule` pour spécifier que les seules valeurs valides sont celles d'une liste finie.</span><span class="sxs-lookup"><span data-stu-id="766a8-155">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="766a8-156">Voir l'exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="766a8-156">The following is an example.</span></span> <span data-ttu-id="766a8-157">Tenez compte des informations suivantes concernant ce code :</span><span class="sxs-lookup"><span data-stu-id="766a8-157">Note the following about this code:</span></span>

- <span data-ttu-id="766a8-158">Il suppose qu'il existe une feuille de calcul nommée "Noms" et que les valeurs de la plage "A1: A3" sont des noms.</span><span class="sxs-lookup"><span data-stu-id="766a8-158">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="766a8-159">La propriété `source` spécifie la liste des valeurs valides.</span><span class="sxs-lookup"><span data-stu-id="766a8-159">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="766a8-160">La plage avec les noms lui a été affectée.</span><span class="sxs-lookup"><span data-stu-id="766a8-160">The range with the names has been assigned to it.</span></span> <span data-ttu-id="766a8-161">Vous pouvez également affecter une liste délimitée par des virgules, comme par exemple : « Sue, Ricky, Liz ».</span><span class="sxs-lookup"><span data-stu-id="766a8-161">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="766a8-162">La propriété `inCellDropDown` spécifie si un contrôle déroulant apparaîtra dans la cellule lorsque l'utilisateur le sélectionne.</span><span class="sxs-lookup"><span data-stu-id="766a8-162">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="766a8-163">Si elle est définie sur `true`, une liste déroulante apparaît contenant la liste des valeurs du `source`.</span><span class="sxs-lookup"><span data-stu-id="766a8-163">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

#### <a name="custom-validation-rule-type"></a><span data-ttu-id="766a8-164">Type de règle de validation personnalisée</span><span class="sxs-lookup"><span data-stu-id="766a8-164">Custom validation rule type</span></span>

<span data-ttu-id="766a8-165">Utilisez la propriété `custom` dans l'objet `DataValidationRule` pour spécifier une formule de validation personnalisée.</span><span class="sxs-lookup"><span data-stu-id="766a8-165">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="766a8-166">Voir l'exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="766a8-166">The following is an example.</span></span> <span data-ttu-id="766a8-167">Tenez compte des informations suivantes concernant ce code :</span><span class="sxs-lookup"><span data-stu-id="766a8-167">Note the following about this code:</span></span>

- <span data-ttu-id="766a8-168">Il suppose qu'il y a un tableau à deux colonnes avec des colonnes **Nom de l'athlète** et **Commentaires** dans les colonnes A et B de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="766a8-168">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="766a8-169">Pour réduire la verbosité dans la colonne **Commentaires,** il rend invalides les données qui incluent le nom de l'athlète.</span><span class="sxs-lookup"><span data-stu-id="766a8-169">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="766a8-170">`SEARCH(A2,B2)` renvoie la position de départ, de la chaîne dans B2, de la chaîne dans A2.</span><span class="sxs-lookup"><span data-stu-id="766a8-170">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="766a8-171">Si A2 n'est pas contenu dans B2, il ne renvoie pas de nombre.</span><span class="sxs-lookup"><span data-stu-id="766a8-171">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="766a8-172">`ISNUMBER()` retourne un booléen.</span><span class="sxs-lookup"><span data-stu-id="766a8-172">Returns a `ISNUMBER()`.</span></span> <span data-ttu-id="766a8-173">La propriété `formula` indique donc que les données valides pour la colonne **Commentaire** sont les données qui n'incluent pas la chaîne présente dans la colonne **Nom de l'athlète**.</span><span class="sxs-lookup"><span data-stu-id="766a8-173">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

### <a name="create-validation-error-alerts"></a><span data-ttu-id="766a8-174">Créer des alertes d'erreur de validation</span><span class="sxs-lookup"><span data-stu-id="766a8-174">Create validation error alerts</span></span>

<span data-ttu-id="766a8-175">Vous pouvez créer une alerte d'erreur personnalisée qui apparaît lorsqu'un utilisateur tente d'entrer des données non valides dans une cellule.</span><span class="sxs-lookup"><span data-stu-id="766a8-175">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="766a8-176">Ce qui suit est un simple exemple.</span><span class="sxs-lookup"><span data-stu-id="766a8-176">The following is a simple example:</span></span> <span data-ttu-id="766a8-177">Tenez compte des informations suivantes concernant ce code :</span><span class="sxs-lookup"><span data-stu-id="766a8-177">Note the following about this code:</span></span>

- <span data-ttu-id="766a8-178">La propriété `style` détermine si l'utilisateur reçoit une alerte informative, un avertissement ou une alerte d' "arrêt".</span><span class="sxs-lookup"><span data-stu-id="766a8-178">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="766a8-179">Seule `Stop` empêche réellement l'utilisateur d'ajouter des données invalides.</span><span class="sxs-lookup"><span data-stu-id="766a8-179">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="766a8-180">La fenêtre contextuelle pour `Warning` et `Information` a des options qui permettent à l'utilisateur d'entrer les données invalides de toute façon.</span><span class="sxs-lookup"><span data-stu-id="766a8-180">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="766a8-181">La propriété `showAlert` prend `true` par défaut.</span><span class="sxs-lookup"><span data-stu-id="766a8-181">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="766a8-182">Cela signifie que l'hôte Excel affichera une alerte générique (de type `Stop`) sauf si vous créez une alerte personnalisée qui soit définit `showAlert` pour `false` ou définit un message, un titre et un style personnalisés.</span><span class="sxs-lookup"><span data-stu-id="766a8-182">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="766a8-183">Ce code définit un message et un titre personnalisés.</span><span class="sxs-lookup"><span data-stu-id="766a8-183">This code sets a custom message and title.</span></span>


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

<span data-ttu-id="766a8-184">Pour plus d'informations, voir [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="766a8-184">For more information see   </span></span>

### <a name="create-validation-prompts"></a><span data-ttu-id="766a8-185">Créer des invites de validation</span><span class="sxs-lookup"><span data-stu-id="766a8-185">Create validation prompts</span></span>

<span data-ttu-id="766a8-186">Vous pouvez créer une invite d'instruction qui apparaît lorsqu'un utilisateur survole ou sélectionne une cellule à laquelle la validation des données a été appliquée.</span><span class="sxs-lookup"><span data-stu-id="766a8-186">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="766a8-187">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="766a8-187">The following is an example:</span></span>

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

<span data-ttu-id="766a8-188">Pour plus d'informations, voir [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="766a8-188">For more information see   </span></span>

### <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="766a8-189">Supprimer la validation des données d'une plage</span><span class="sxs-lookup"><span data-stu-id="766a8-189">Remove data validation from a range</span></span>

<span data-ttu-id="766a8-190">Pour supprimer la validation des données d'une plage, appelez la méthode [Range.dataValidation.clear ()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear).</span><span class="sxs-lookup"><span data-stu-id="766a8-190">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="766a8-191">Il n'est pas nécessaire que la plage que vous effacez soit exactement la même plage que celle sur laquelle vous avez ajouté la validation des données.</span><span class="sxs-lookup"><span data-stu-id="766a8-191">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="766a8-192">Si ce n'est pas le cas, seules les cellules des deux plages qui se chevauchent sont effacées, le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="766a8-192">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="766a8-193">L'effacement de la validation des données d'une plage efface également toute validation de données qu'un utilisateur a ajoutée manuellement à la plage.</span><span class="sxs-lookup"><span data-stu-id="766a8-193">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="766a8-194">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="766a8-194">See also</span></span>

- [<span data-ttu-id="766a8-195">Concepts de base de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="766a8-195">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="766a8-196">Objet DataValidation (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="766a8-196">Worksheet Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/datavalidation)
- [<span data-ttu-id="766a8-197">Objet Range (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="766a8-197">Range Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/range)



 
