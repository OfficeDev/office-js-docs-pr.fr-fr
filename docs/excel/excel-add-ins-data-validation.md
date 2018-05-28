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
# <a name="add-data-validation-to-excel-ranges-preview"></a><span data-ttu-id="b6ba0-102">Ajouter la validation des donn?es aux plages Excel (pr?version)</span><span class="sxs-lookup"><span data-stu-id="b6ba0-102">Add data validation to Excel ranges (Preview)</span></span>

> [!NOTE]
> <span data-ttu-id="b6ba0-103">Tant que les API de validation des donn?es sont en pr?version, vous devez charger la version b?ta de la biblioth?que JavaScript Office pour les utiliser.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-103">While the data validation APIs are in preview, you must load the beta version of the Office JavaScript library to use them.</span></span> <span data-ttu-id="b6ba0-104">L'URL est https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-104">The full URL is https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span> <span data-ttu-id="b6ba0-105">Si vous utilisez TypeScript ou si votre ?diteur de code utilise un fichier de d?finition de type TypeScript pour intelliSense, utilisez https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-105">If you are using TypeScript or your code editor uses a TypeScript type definition file for intellisense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="b6ba0-106">La biblioth?que JavaScript Excel fournit des API pour permettre ? votre compl?ment d'ajouter une validation automatique des donn?es aux tables, colonnes, lignes et autres plages d'un classeur.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-106">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="b6ba0-107">Pour comprendre les concepts et la terminologie de la validation des donn?es, consultez les articles suivants qui portent sur la mani?re dont les utilisateurs peuvent ajouter la validation des donn?es via l'IU Excel?:</span><span class="sxs-lookup"><span data-stu-id="b6ba0-107">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="b6ba0-108">Appliquer la validation des donn?es aux cellules</span><span class="sxs-lookup"><span data-stu-id="b6ba0-108">Apply data validation to cells</span></span>](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="b6ba0-109">Plus d'informations sur la validation des donn?es</span><span class="sxs-lookup"><span data-stu-id="b6ba0-109">More on data validation</span></span>](https://microsoft.sharepoint.com/:p:/r/teams/oext/_layouts/15/Doc.aspx?sourcedoc=%7B51143964-d52c-429d-bfac-c7495473d536%7D&action=edit)
- [<span data-ttu-id="b6ba0-110">Description et exemples de validation de donn?es dans Excel</span><span class="sxs-lookup"><span data-stu-id="b6ba0-110">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="b6ba0-111">Contr?le programmatique de la validation des donn?es</span><span class="sxs-lookup"><span data-stu-id="b6ba0-111">Programmatic control of data validation</span></span>

<span data-ttu-id="b6ba0-112">La propri?t? `Range.dataValidation`, qui prend un objet de validation de donn?es [,](https://dev.office.com/reference/add-ins/excel/datavalidation) est le point d'entr?e pour le contr?le programmatique de la validation des donn?es dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-112">The `Range.dataValidation` property, which takes a [DataValidation](https://dev.office.com/reference/add-ins/excel/datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="b6ba0-113">Il y a cinq propri?t?s ? l'objet `DataValidation`?:</span><span class="sxs-lookup"><span data-stu-id="b6ba0-113">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="b6ba0-114">`rule` - D?finit ce qui constitue des donn?es valides pour la plage.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-114">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="b6ba0-115">Voir [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="b6ba0-115">See [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).</span></span>
- <span data-ttu-id="b6ba0-116">`errorAlert` - Sp?cifie si une erreur appara?t lorsque l'utilisateur entre des donn?es non valides et d?finit le texte, le titre et le style d'alerte?; par exemple, **Informatif**, **Avertissement**, et **Arr?ter**.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-116">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="b6ba0-117">Voir [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="b6ba0-117">See [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span></span>
- <span data-ttu-id="b6ba0-118">`prompt` - Indique si une invite s'affiche lorsque l'utilisateur survole la plage et d?finit le message d'assistance vocale.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-118">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="b6ba0-119">Voir [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="b6ba0-119">See [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span></span>
- <span data-ttu-id="b6ba0-120">`ignoreBlanks` - Sp?cifie si la r?gle de validation des donn?es s'applique aux cellules vides de la plage.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-120">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="b6ba0-121">Par d?faut `true`.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-121">Defaults to `true`.</span></span>
- <span data-ttu-id="b6ba0-122">`type` - Une identification en lecture seule du type de validation, tel que WholeNumber, Date, TextLength, etc. Il est d?fini indirectement lorsque vous d?finissez la propri?t? `rule`.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-122">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="b6ba0-123">La validation des donn?es ajout?e de fa?on programmatique se comporte exactement comme la validation des donn?es ajout?e manuellement.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-123">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="b6ba0-124">Surtout, notez que la validation des donn?es est d?clench?e uniquement si l'utilisateur entre directement une valeur dans une cellule ou copie et colle une cellule d'un autre emplacement dans le classeur et choisit l'option de collage, valeurs ****.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-124">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="b6ba0-125">Si l'utilisateur copie une cellule et la colle simplement dans une plage avec validation des donn?es, la validation n'est pas d?clench?e.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-125">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

### <a name="creating-validation-rules"></a><span data-ttu-id="b6ba0-126">Cr?er des r?gles de validation</span><span class="sxs-lookup"><span data-stu-id="b6ba0-126">Creating validation rules</span></span>

<span data-ttu-id="b6ba0-127">Pour ajouter une validation de donn?es ? une plage, votre code doit d?finir la propri?t? `rule` de l'objet `DataValidation` dans `Range.dataValidation`.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-127">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="b6ba0-128">Cela prend un objet [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) qui a sept propri?t?s facultatives.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-128">This takes a [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="b6ba0-129">*Une seule de ces propri?t?s peut ?tre pr?sente dans un objet `DataValidationRule`.*</span><span class="sxs-lookup"><span data-stu-id="b6ba0-129">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="b6ba0-130">La propri?t? que vous incluez d?termine le type de validation.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-130">The property that you include determines the type of validation.</span></span>

#### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="b6ba0-131">Types de r?gles de validation de base et DateTime</span><span class="sxs-lookup"><span data-stu-id="b6ba0-131">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="b6ba0-132">Les trois premi?res propri?t?s `DataValidationRule` (c.-?-d. les types de r?gles de validation) prennent un objet [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) comme leur valeur.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-132">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="b6ba0-133">`wholeNumber` ? N?cessite un nombre entier en plus de toute autre validation sp?cifi?e par l'objet `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-133">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="b6ba0-134">`decimal` - N?cessite un nombre d?cimal en plus de toute autre validation sp?cifi?e par l'objet `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-134">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="b6ba0-135">`textLength` ? Applique les d?tails de validation dans l'objet `BasicDataValidation` ? la *longueur*  de la valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-135">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="b6ba0-136">Voici un exemple de cr?ation d'une r?gle de validation.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-136">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="b6ba0-137">Tenez compte des informations suivantes concernant ce code?:</span><span class="sxs-lookup"><span data-stu-id="b6ba0-137">Note the following about this code:</span></span>

- <span data-ttu-id="b6ba0-138">Le `operator` est l'op?rateur binaire "sup?rieur ?".</span><span class="sxs-lookup"><span data-stu-id="b6ba0-138">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="b6ba0-139">Chaque fois que vous utilisez un op?rateur binaire, la valeur que l'utilisateur essaie d'entrer dans la cellule est l'op?rande de gauche et la valeur sp?cifi?e dans `formula1` est l'op?rande de droite.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-139">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="b6ba0-140">Donc, cette r?gle dit que seuls les nombres entiers sup?rieurs ? 0 sont valides.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-140">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="b6ba0-141">Le `formula1` est un nombre cod? en dur.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-141">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="b6ba0-142">Si vous ne savez pas au moment du codage quelle devrait ?tre la valeur, vous pouvez ?galement utiliser une formule Excel (sous forme de cha?ne) pour la valeur.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-142">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="b6ba0-143">Par exemple, ??= A3?? et ??= SOMME (A4:B5)?? peuvent ?galement ?tre des valeurs de `formula1`.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-143">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="b6ba0-144">Voir [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) pour une liste des autres op?rateurs binaires.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-144">See [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="b6ba0-145">Il existe aussi deux op?rateurs ternaires?: ??Between?? et ??NotBetween??.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-145">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="b6ba0-146">Pour les utiliser, vous devez sp?cifier la propri?t? `formula2` facultative.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-146">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="b6ba0-147">Les valeurs `formula1` et `formula2` sont les op?randes de d?limitation.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-147">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="b6ba0-148">La valeur que l'utilisateur essaie d'entrer dans la cellule est le troisi?me op?rande (?valu?e).</span><span class="sxs-lookup"><span data-stu-id="b6ba0-148">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="b6ba0-149">Voici un exemple d'utilisation de l'op?rateur ??Between???:</span><span class="sxs-lookup"><span data-stu-id="b6ba0-149">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="b6ba0-150">Les deux propri?t?s de r?gle suivantes prennent un objet [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) comme leur valeur.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-150">The next two rule properties take a [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="b6ba0-151">L'objet `DateTimeDataValidation` est structur? de mani?re similaire ? la `BasicDataValidation`?: il a les propri?t?s `formula1`, `formula2`, et `operator`, et est utilis? de la m?me mani?re.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-151">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="b6ba0-152">La diff?rence est que vous ne pouvez pas utiliser un nombre dans les propri?t?s de la formule, mais vous pouvez entrer une cha?ne [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou une formule Excel).</span><span class="sxs-lookup"><span data-stu-id="b6ba0-152">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="b6ba0-153">Voici un exemple qui d?finit des valeurs valides telles que des dates dans la premi?re semaine d'avril 2018.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-153">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

#### <a name="list-validation-rule-type"></a><span data-ttu-id="b6ba0-154">Type de r?gle de validation de liste</span><span class="sxs-lookup"><span data-stu-id="b6ba0-154">List validation rule type</span></span>

<span data-ttu-id="b6ba0-155">Utilisez la propri?t? `list` dans l'objet `DataValidationRule` pour sp?cifier que les seules valeurs valides sont celles d'une liste finie.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-155">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="b6ba0-156">Voir l'exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-156">The following is an example.</span></span> <span data-ttu-id="b6ba0-157">Tenez compte des informations suivantes concernant ce code?:</span><span class="sxs-lookup"><span data-stu-id="b6ba0-157">Note the following about this code:</span></span>

- <span data-ttu-id="b6ba0-158">Il suppose qu'il existe une feuille de calcul nomm?e "Noms" et que les valeurs de la plage "A1: A3" sont des noms.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-158">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="b6ba0-159">La propri?t? `source` sp?cifie la liste des valeurs valides.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-159">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="b6ba0-160">La plage avec les noms lui a ?t? affect?e.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-160">The range with the names has been assigned to it.</span></span> <span data-ttu-id="b6ba0-161">Vous pouvez ?galement affecter une liste d?limit?e par des virgules, comme par exemple?: ??Sue, Ricky, Liz??.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-161">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="b6ba0-162">La propri?t? `inCellDropDown` sp?cifie si un contr?le d?roulant appara?tra dans la cellule lorsque l'utilisateur le s?lectionne.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-162">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="b6ba0-163">Si elle est d?finie sur `true`, une liste d?roulante appara?t contenant la liste des valeurs du `source`.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-163">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

#### <a name="custom-validation-rule-type"></a><span data-ttu-id="b6ba0-164">Type de r?gle de validation personnalis?e</span><span class="sxs-lookup"><span data-stu-id="b6ba0-164">Custom validation rule type</span></span>

<span data-ttu-id="b6ba0-165">Utilisez la propri?t? `custom` dans l'objet `DataValidationRule` pour sp?cifier une formule de validation personnalis?e.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-165">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="b6ba0-166">Voir l'exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-166">The following is an example.</span></span> <span data-ttu-id="b6ba0-167">Tenez compte des informations suivantes concernant ce code?:</span><span class="sxs-lookup"><span data-stu-id="b6ba0-167">Note the following about this code:</span></span>

- <span data-ttu-id="b6ba0-168">Il suppose qu'il y a un tableau ? deux colonnes avec des colonnes **Nom de l'athl?te** et **Commentaires** dans les colonnes A et B de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-168">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="b6ba0-169">Pour r?duire la verbosit? dans la colonne **Commentaires,** il rend invalides les donn?es qui incluent le nom de l'athl?te.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-169">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="b6ba0-170">`SEARCH(A2,B2)` renvoie la position de d?part, de la cha?ne dans B2, de la cha?ne dans A2.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-170">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="b6ba0-171">Si A2 n'est pas contenu dans B2, il ne renvoie pas de nombre.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-171">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="b6ba0-172">`ISNUMBER()` retourne un bool?en.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-172">Returns a `ISNUMBER()`.</span></span> <span data-ttu-id="b6ba0-173">La propri?t? `formula` indique donc que les donn?es valides pour la colonne **Commentaire** sont les donn?es qui n'incluent pas la cha?ne pr?sente dans la colonne **Nom de l'athl?te**.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-173">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

### <a name="create-validation-error-alerts"></a><span data-ttu-id="b6ba0-174">Cr?er des alertes d'erreur de validation</span><span class="sxs-lookup"><span data-stu-id="b6ba0-174">Create validation error alerts</span></span>

<span data-ttu-id="b6ba0-175">Vous pouvez cr?er une alerte d'erreur personnalis?e qui appara?t lorsqu'un utilisateur tente d'entrer des donn?es non valides dans une cellule.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-175">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="b6ba0-176">Ce qui suit est un simple exemple.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-176">The following is a simple example:</span></span> <span data-ttu-id="b6ba0-177">Tenez compte des informations suivantes concernant ce code?:</span><span class="sxs-lookup"><span data-stu-id="b6ba0-177">Note the following about this code:</span></span>

- <span data-ttu-id="b6ba0-178">La propri?t? `style` d?termine si l'utilisateur re?oit une alerte informative, un avertissement ou une alerte d' "arr?t".</span><span class="sxs-lookup"><span data-stu-id="b6ba0-178">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="b6ba0-179">Seule `Stop` emp?che r?ellement l'utilisateur d'ajouter des donn?es invalides.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-179">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="b6ba0-180">La fen?tre contextuelle pour `Warning` et `Information` a des options qui permettent ? l'utilisateur d'entrer les donn?es invalides de toute fa?on.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-180">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="b6ba0-181">La propri?t? `showAlert` prend `true` par d?faut.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-181">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="b6ba0-182">Cela signifie que l'h?te Excel affichera une alerte g?n?rique (de type `Stop`) sauf si vous cr?ez une alerte personnalis?e qui soit d?finit `showAlert` pour `false` ou d?finit un message, un titre et un style personnalis?s.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-182">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="b6ba0-183">Ce code d?finit un message et un titre personnalis?s.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-183">This code sets a custom message and title.</span></span>


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

<span data-ttu-id="b6ba0-184">Pour plus d'informations, voir [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="b6ba0-184">For more information see   </span></span>

### <a name="create-validation-prompts"></a><span data-ttu-id="b6ba0-185">Cr?er des invites de validation</span><span class="sxs-lookup"><span data-stu-id="b6ba0-185">Create validation prompts</span></span>

<span data-ttu-id="b6ba0-186">Vous pouvez cr?er une invite d'instruction qui appara?t lorsqu'un utilisateur survole ou s?lectionne une cellule ? laquelle la validation des donn?es a ?t? appliqu?e.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-186">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="b6ba0-187">Voici un exemple?:</span><span class="sxs-lookup"><span data-stu-id="b6ba0-187">The following is an example:</span></span>

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

<span data-ttu-id="b6ba0-188">Pour plus d'informations, voir [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="b6ba0-188">For more information see   </span></span>

### <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="b6ba0-189">Supprimer la validation des donn?es d'une plage</span><span class="sxs-lookup"><span data-stu-id="b6ba0-189">Remove data validation from a range</span></span>

<span data-ttu-id="b6ba0-190">Pour supprimer la validation des donn?es d'une plage, appelez la m?thode [Range.dataValidation.clear ()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear).</span><span class="sxs-lookup"><span data-stu-id="b6ba0-190">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="b6ba0-191">Il n'est pas n?cessaire que la plage que vous effacez soit exactement la m?me plage que celle sur laquelle vous avez ajout? la validation des donn?es.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-191">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="b6ba0-192">Si ce n'est pas le cas, seules les cellules des deux plages qui se chevauchent sont effac?es, le cas ?ch?ant.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-192">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="b6ba0-193">L'effacement de la validation des donn?es d'une plage efface ?galement toute validation de donn?es qu'un utilisateur a ajout?e manuellement ? la plage.</span><span class="sxs-lookup"><span data-stu-id="b6ba0-193">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="b6ba0-194">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b6ba0-194">See also</span></span>

- [<span data-ttu-id="b6ba0-195">Concepts de base de l?API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="b6ba0-195">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b6ba0-196">Objet DataValidation (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="b6ba0-196">Worksheet Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/datavalidation)
- [<span data-ttu-id="b6ba0-197">Objet Range (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="b6ba0-197">Range Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/range)



 
