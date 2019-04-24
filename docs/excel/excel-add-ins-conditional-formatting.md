---
title: Appliquer la mise en forme conditionnelle aux plages avec l’API Excel JavaScript
description: ''
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 7c6b5b5433e2dc59259eb937ef553ff265443f75
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449443"
---
# <a name="apply-conditional-formatting-to-excel-ranges"></a><span data-ttu-id="dcf3c-102">Appliquer une mise en forme conditionnelle à des plages Excel</span><span class="sxs-lookup"><span data-stu-id="dcf3c-102">Apply conditional formatting to Excel ranges</span></span>

<span data-ttu-id="dcf3c-103">La bibliothèque JavaScript Excel fournit des API pour appliquer une mise en forme conditionnelle aux plages de données dans vos feuilles de calcul.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-103">The Excel JavaScript Library provides APIs to apply conditional formatting to data ranges in your worksheets.</span></span> <span data-ttu-id="dcf3c-104">Cette fonctionnalité simplifie l’analyse visuelle de grands ensembles de données.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-104">This functionality makes large sets of data easy to visually parse.</span></span> <span data-ttu-id="dcf3c-105">La mise en forme effectue également des mises à jour dynamiques en fonction des changements dans la plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-105">The formatting also dynamically updates based on changes within the range.</span></span> 

> [!NOTE]
> <span data-ttu-id="dcf3c-106">Cet article décrit la mise en forme conditionnelle dans le contexte de compléments Excel JavaScript. Les articles suivants offrent des informations détaillées sur les fonctionnalités de mise en forme conditionnelles complètes dans Excel.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-106">This article covers conditional formatting in the context of Excel JavaScript add-ins. The following articles provide detailed information about the full conditional formatting capabilities within Excel.</span></span>
> -  [<span data-ttu-id="dcf3c-107">Ajouter, modifier ou effacer des formats conditionnels</span><span class="sxs-lookup"><span data-stu-id="dcf3c-107">Add, change, or clear conditional formats</span></span>](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
> -  [<span data-ttu-id="dcf3c-108">Utilisez des formules avec mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="dcf3c-108">Use formulas with conditional formatting</span></span>](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)

## <a name="programmatic-control-of-conditional-formatting"></a><span data-ttu-id="dcf3c-109">Contrôle par programme de mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="dcf3c-109">Programmatic control of conditional formatting</span></span>

<span data-ttu-id="dcf3c-110">La `Range.conditionalFormats` propriété est un ensemble d’objets [ConditionalFormat](/javascript/api/excel/excel.conditionalformat)qui s’appliquent à la plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-110">The `Range.conditionalFormats` property is a collection of [ConditionalFormat](/javascript/api/excel/excel.conditionalformat) objects that apply to the range.</span></span>  <span data-ttu-id="dcf3c-111">L’`ConditionalFormat` objet contient plusieurs propriétés qui définissent le format à appliquer en fonction du [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).</span><span class="sxs-lookup"><span data-stu-id="dcf3c-111">The `ConditionalFormat` object contains several properties that define the format to be applied based on the [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).</span></span> 

-   `cellValue`
-   `colorScale`
-   `custom`
-   `dataBar`
-   `iconSet`
-   `preset`
-   `textComparison`
-   `topBottom`

> [!NOTE]
> <span data-ttu-id="dcf3c-112">Chacune de ces propriétés de mise en forme a une variante correspondante`*OrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-112">Each of these formatting properties has a corresponding `*OrNullObject` variant.</span></span> <span data-ttu-id="dcf3c-113">En savoir plus sur ce modèle dans la section[\* OrNullObject méthodes](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="dcf3c-113">Learn more about that pattern in the [\*OrNullObject methods](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods) section.</span></span>

<span data-ttu-id="dcf3c-114">Un seul type de format peut être défini pour l’objet ConditionalFormat.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-114">Only one format type can be set for the ConditionalFormat object.</span></span> <span data-ttu-id="dcf3c-115">Cela est déterminé par la `type` propriété, c'est-à-dire une [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype) valeur enum.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-115">This is determined by the `type` property, which is a [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype) enum value.</span></span> <span data-ttu-id="dcf3c-116">`type` est défini lorsque vous ajoutez une mise en forme conditionnelle à une plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-116">`type` is set when adding a conditional format to a range.</span></span> 

## <a name="creating-conditional-formatting-rules"></a><span data-ttu-id="dcf3c-117">Modification des règles de mise en forme conditionnelles</span><span class="sxs-lookup"><span data-stu-id="dcf3c-117">Creating conditional formatting rules</span></span>

<span data-ttu-id="dcf3c-118">Les mises en forme conditionnelles sont ajoutées à une plage à l’aide de `conditionalFormats.add`.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-118">Conditional formats are added to a range by using `conditionalFormats.add`.</span></span> <span data-ttu-id="dcf3c-119">Une fois ajoutées, vous pouvez définir les propriétés spécifiques à la mise en forme conditionnelle.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-119">Once added, the properties specific to the conditional format can be set.</span></span> <span data-ttu-id="dcf3c-120">Les exemples ci-dessous montrent la création de différents types de mise en forme.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-120">The following examples show the creation of different formatting types.</span></span>

### <a name="cell-valuejavascriptapiexcelexcelcellvalueconditionalformat"></a>[<span data-ttu-id="dcf3c-121">Valeur de la cellule</span><span class="sxs-lookup"><span data-stu-id="dcf3c-121">Cell value</span></span>](/javascript/api/excel/excel.cellvalueconditionalformat)

<span data-ttu-id="dcf3c-122">La mise en forme conditionnelle de valeur de la cellule applique un format défini par l’utilisateur en fonction des résultats d’une ou deux formules dans la [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule).</span><span class="sxs-lookup"><span data-stu-id="dcf3c-122">Cell value conditional formatting applies a user-defined format based on the results of one or two formulas in the [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule).</span></span> <span data-ttu-id="dcf3c-123">La`operator` propriété est un [ConditionalCellValueOperator](/javascript/api/excel/excel.conditionalcellvalueoperator) définissant comment les expressions qui en résultent sont liées à la mise en forme.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-123">The `operator` property is a [ConditionalCellValueOperator](/javascript/api/excel/excel.conditionalcellvalueoperator) defining how the resulting expressions relate to the formatting.</span></span>

<span data-ttu-id="dcf3c-124">L’exemple suivant montre une coloration de la police en rouge appliquée à une valeur dans la plage inférieure à zéro.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-124">The following example shows red font coloring applied to any value in the range less than zero.</span></span>

![Une plage avec des nombres négatifs en rouge.](../images/excel-conditional-format-cell-value.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B21:E23");
const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.cellValue
);

// set the font of negative numbers to red
conditionalFormat.cellValue.format.font.color = "red";
conditionalFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };

await context.sync();
```

### <a name="color-scalejavascriptapiexcelexcelcolorscaleconditionalformat"></a>[<span data-ttu-id="dcf3c-126">Échelle de couleur</span><span class="sxs-lookup"><span data-stu-id="dcf3c-126">Color scale</span></span>](/javascript/api/excel/excel.colorscaleconditionalformat)

<span data-ttu-id="dcf3c-127">La mise en forme conditionnelle de l’échelle de couleur applique un dégradé de couleur au sein de la plage de données.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-127">Color scale conditional formatting applies a color gradient across the data range.</span></span> <span data-ttu-id="dcf3c-128">La`criteria` propriété sur le `ColorScaleConditionalFormat` définit trois[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`et éventuellement, `midpoint`.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-128">The `criteria` property on the `ColorScaleConditionalFormat` defines three [ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`, and, optionally, `midpoint`.</span></span> <span data-ttu-id="dcf3c-129">Les critères des points d’échelle ont trois propriétés :</span><span class="sxs-lookup"><span data-stu-id="dcf3c-129">Each of the criterion scale points have three properties:</span></span>

-   <span data-ttu-id="dcf3c-130">`color` -Le code de couleur HTML pour le point de terminaison.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-130">`color` - The HTML color code for the endpoint.</span></span>
-   <span data-ttu-id="dcf3c-131">`formula` -Un nombre ou une formule représentant le point de terminaison.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-131">`formula` - A number or formula representing the endpoint.</span></span> <span data-ttu-id="dcf3c-132">Il s’agit de `null` si `type` est `lowestValue` ou `highestValue`.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-132">This will be `null` if `type` is `lowestValue` or `highestValue`.</span></span>
-   <span data-ttu-id="dcf3c-133">`type` -Comment la formule doit être évaluée.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-133">`type` - How the formula should be evaluated.</span></span> <span data-ttu-id="dcf3c-134">`highestValue` et `lowestValue` font référence à des valeurs dans la plage en cours de mise en forme.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-134">`highestValue` and `lowestValue` refer to values in the range being formatted.</span></span>

<span data-ttu-id="dcf3c-135">L’exemple suivant montre une plage colorée de bleue à jaune à rouge.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-135">The following example shows a range being colored blue to yellow to red.</span></span> <span data-ttu-id="dcf3c-136">Notez que `minimum` et `maximum` sont les valeurs inférieures et supérieures respectivement et utilisent les `null` formules.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-136">Note that `minimum` and `maximum` are the lowest and highest values respectively and use `null` formulas.</span></span> <span data-ttu-id="dcf3c-137">`midpoint` utilise le `percentage` type avec une formule de `”=50”` donc la cellule jaune est la valeur moyenne.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-137">`midpoint` is using the `percentage` type with a formula of `”=50”` so the yellowest cell is the mean value.</span></span>

![Une plage avec un petit nombre en bleu, nombre moyen en jaune et élevé en rouge, avec des dégradés entre les valeurs.](../images/excel-conditional-format-color-scale.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
      Excel.ConditionalFormatType.colorScale
);

// color the backgrounds of the cells from blue to yellow to red based on value
const criteria = {
      minimum: {
           formula: null,
           type: Excel.ConditionalFormatColorCriterionType.lowestValue,
           color: "blue"
      },
      midpoint: {
           formula: "50",
           type: Excel.ConditionalFormatColorCriterionType.percent,
           color: "yellow"
      },
      maximum: {
           formula: null,
           type: Excel.ConditionalFormatColorCriterionType.highestValue,
           color: "red"
      }
};
conditionalFormat.colorScale.criteria = criteria;

await context.sync();
```

### <a name="customjavascriptapiexcelexcelcustomconditionalformat"></a>[<span data-ttu-id="dcf3c-139">Personnalisé</span><span class="sxs-lookup"><span data-stu-id="dcf3c-139">Custom</span></span>](/javascript/api/excel/excel.customconditionalformat)

<span data-ttu-id="dcf3c-140">La mise en forme conditionnelle personnalisée applique un format défini par l’utilisateur aux cellules en fonction d’une formule de complexité arbitraire.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-140">Custom conditional formatting applies a user-defined format to the cells based on a formula of arbitrary complexity.</span></span> <span data-ttu-id="dcf3c-141">L’objet [ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule) vous permet de définir la formule dans des notations différentes :</span><span class="sxs-lookup"><span data-stu-id="dcf3c-141">The [ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule) object lets you define the formula in different notations:</span></span>

-   <span data-ttu-id="dcf3c-142">`formula` -Notation standard.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-142">`formula` - Standard notation.</span></span>
-   <span data-ttu-id="dcf3c-143">`formulaLocal` -Localisé en fonction de langue de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-143">`formulaLocal` - Localized based on the user’s language.</span></span>
-   <span data-ttu-id="dcf3c-144">`formulaR1C1` -Notation type L1C1.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-144">`formulaR1C1` - R1C1-style notation.</span></span>

<span data-ttu-id="dcf3c-145">L’exemple suivant colore les polices de cellules avec des valeurs supérieures à la cellule située à leur gauche en vert.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-145">The following example colors the fonts green of cells with higher values than the cell to their left.</span></span>

![Une plage avec des nombres verts place la valeur de la colonne précédente dans cette ligne comme inférieure.](../images/excel-conditional-format-custom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.custom
);

// if a cell has a higher value than the one to its left, set that cell’s font to green
conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT("RC[-1]",0),TRUE)';
conditionalFormat.custom.format.font.color = "green";

await context.sync();

```
### <a name="data-barjavascriptapiexcelexceldatabarconditionalformat"></a>[<span data-ttu-id="dcf3c-147">Barre de données</span><span class="sxs-lookup"><span data-stu-id="dcf3c-147">Data bar</span></span>](/javascript/api/excel/excel.databarconditionalformat)

<span data-ttu-id="dcf3c-148">La mise en forme conditionnelle de la barre de données ajoute des barres de données aux cellules.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-148">Data bar conditional formatting adds data bars to the cells.</span></span> <span data-ttu-id="dcf3c-149">Par défaut, les valeurs minimales et maximales dans la plage forment les limites et les tailles proportionnelles des barres de données.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-149">By default, the minimum and maximum values in the Range form the bounds and proportional sizes of the data bars.</span></span> <span data-ttu-id="dcf3c-150">L’objet `DataBarConditionalFormat` a plusieurs propriétés pour contrôler l’apparence de la barre.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-150">The `DataBarConditionalFormat` object has several properties to control the bar’s appearance.</span></span> 

<span data-ttu-id="dcf3c-151">L’exemple suivant met en forme la plage contenant des barres de données remplissant de gauche à droite.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-151">The following example formats the range with data bars filling left-to-right.</span></span>

![Une plage avec barre de données derrière les valeurs dans les cellules.](../images/excel-conditional-format-databar.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.dataBar
);

// give left-to-right, default-appearance data bars to all the cells
conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
await context.sync();
```

### <a name="icon-setjavascriptapiexcelexceliconsetconditionalformat"></a>[<span data-ttu-id="dcf3c-153">Jeu d’icônes</span><span class="sxs-lookup"><span data-stu-id="dcf3c-153">Icon set</span></span>](/javascript/api/excel/excel.iconsetconditionalformat)

<span data-ttu-id="dcf3c-154">La mise en forme conditionnelle du jeu d’icônes utilise Excel [icônes](/javascript/api/excel/excel.icon) pour mettre en surbrillance les cellules.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-154">Icon set conditional formatting uses Excel [Icons](/javascript/api/excel/excel.icon) to highlight cells.</span></span> <span data-ttu-id="dcf3c-155">La `criteria` propriété est une matrice de [ConditionalIconCriterion](/javascript/api/excel/excel.ConditionalIconCriterion), qui définit le symbole à insérer et la condition sous laquelle celui-ci est inséré.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-155">The `criteria` property is an array of [ConditionalIconCriterion](/javascript/api/excel/excel.ConditionalIconCriterion), which define the symbol to be inserted and the condition under which it is inserted.</span></span> <span data-ttu-id="dcf3c-156">Ce tableau est automatiquement pré-rempli avec éléments critères avec les propriétés par défaut.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-156">This array is automatically prepopulated with criterion elements with default properties.</span></span> <span data-ttu-id="dcf3c-157">Les propriétés individuelles ne peut pas être remplacées.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-157">Individual properties cannot be overwritten.</span></span> <span data-ttu-id="dcf3c-158">Au lieu de cela, l’ensemble de l’objet de critères doit être remplacé.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-158">Instead, the whole criteria object must be replaced.</span></span> 

<span data-ttu-id="dcf3c-159">L’exemple suivant montre un jeu d’icônes trois triangles utilisé dans la plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-159">The following example shows a three-triangle icon set applied across the range.</span></span>

![Plage avec des triangles verts vers le haut pour les valeurs supérieures à 1000, lignes jaunes pour les valeurs comprises entre 700 et 1000, et triangles rouges vers le bas pour les valeurs inférieures.](../images/excel-conditional-format-iconset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.iconSet
);

const iconSetCF = conditionalFormat.iconSet;
iconSetCF.style = Excel.IconSet.threeTriangles;

/*
   With a "three*” icon set style, such as "threeTriangles", the third
    element in the criteria array (criteria[2]) defines the "top" icon;
    e.g., a green triangle. The second (criteria[1]) defines the "middle"
    icon, The first (criteria[0]) defines the "low" icon, but it can often 
    be left empty as this method does below, because every cell that
   does not match the other two criteria always gets the low icon.
*/
iconSetCF.criteria = [
    {} as any,
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=700"
      },
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=1000"
      }
];

await context.sync();
```

### <a name="preset-criteriajavascriptapiexcelexcelpresetcriteriaconditionalformat"></a>[<span data-ttu-id="dcf3c-161">Critères prédéfinis</span><span class="sxs-lookup"><span data-stu-id="dcf3c-161">Preset criteria</span></span>](/javascript/api/excel/excel.presetcriteriaconditionalformat)

<span data-ttu-id="dcf3c-162">La mise en forme conditionnelle prédéfinie applique un format défini par l’utilisateur pour la plage basée sur une règle standard sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-162">Preset conditional formatting applies a user-defined format to the range based on a selected standard rule.</span></span> <span data-ttu-id="dcf3c-163">Ces règles sont définies par le [ConditionalFormatPresetCriterion](/javascript/api/excel/excel.ConditionalFormatPresetCriterion) dans le [ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule).</span><span class="sxs-lookup"><span data-stu-id="dcf3c-163">These rules are defined by the [ConditionalFormatPresetCriterion](/javascript/api/excel/excel.ConditionalFormatPresetCriterion) in the [ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule).</span></span> 

<span data-ttu-id="dcf3c-164">L’exemple suivant colore la police en blanc où la valeur d’une cellule est au moins un écart-type standard au-dessus de la moyenne de la plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-164">The following example colors the font white wherever a cell’s value is at least one standard deviation above the range’s average.</span></span>

![Une plage de cellules avec police en blanc où les valeurs sont au moins un écart-type standard au-dessus de la moyenne.](../images/excel-conditional-format-preset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.presetCriteria
);

// color every cell’s font white that is one standard deviation above average relative to the range
conditionalFormat.preset.format.font.color = "white";
conditionalFormat.preset.rule = {
     criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage
};

await context.sync();
```

### <a name="text-comparisonjavascriptapiexcelexceltextconditionalformat"></a>[<span data-ttu-id="dcf3c-166">Comparaison de texte</span><span class="sxs-lookup"><span data-stu-id="dcf3c-166">Text comparison</span></span>](/javascript/api/excel/excel.textconditionalformat)

<span data-ttu-id="dcf3c-167">La mise en forme conditionnelle de comparaison de texte utilise des comparaisons de chaînes comme condition.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-167">Text comparison conditional formatting uses string comparisons as the condition.</span></span> <span data-ttu-id="dcf3c-168">La `rule` propriété est un [ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule) définissant une chaîne à comparer avec la cellule et un opérateur pour spécifier le type de comparaison.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-168">The `rule` property is a [ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule) defining a string to compare with the cell and an operator to specify the type of comparison.</span></span> 

<span data-ttu-id="dcf3c-169">L’exemple suivant colore la police en rouge lorsque le texte d’une cellule contient « Différé ».</span><span class="sxs-lookup"><span data-stu-id="dcf3c-169">The following example formats the font color red when a cell’s text contains “Delayed”.</span></span>

![Une plage de cellules contenant « Différé » en rouge.](../images/excel-conditional-format-text.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B16:D18");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.containsText
);

// color the font of every cell containing “Delayed”
conditionalFormat.textComparison.format.font.color = "red";
conditionalFormat.textComparison.rule = {
     operator: Excel.ConditionalTextOperator.contains,
     text: "Delayed"
};

await context.sync();
```

### <a name="topbottomjavascriptapiexcelexceltopbottomconditionalformat"></a>[<span data-ttu-id="dcf3c-171">Supérieure/inférieure</span><span class="sxs-lookup"><span data-stu-id="dcf3c-171">Top/bottom</span></span>](/javascript/api/excel/excel.TopBottomconditionalformat)

<span data-ttu-id="dcf3c-172">La mise en forme conditionnelle supérieure/inférieure applique un format aux valeurs les plus élevées ou plus faibles d’une plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-172">Top/bottom conditional formatting applies a format to the highest or lowest values in a range.</span></span> <span data-ttu-id="dcf3c-173">La `rule` propriété, de type [ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule), définit si la condition est basée sur le plus élevé ou le plus bas, ainsi que si l’évaluation est en classement ou pourcentage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-173">The `rule` property, which is of type [ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule), sets whether the condition is based on the highest or lowest, as well as whether the evaluation is ranked or percentage-based.</span></span> 

<span data-ttu-id="dcf3c-174">L’exemple suivant applique un surlignage vert à la cellule de valeur plus élevée dans la plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-174">The following example applies a green highlight to the highest value cell in the range.</span></span>


![Une plage avec le nombre le plus élevé est mise en surbrillance en vert.](../images/excel-conditional-format-topbottom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B21:E23");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.topBottom
);

// for the highest valued cell in the range, make the background green
conditionalFormat.topBottom.format.fill.color = "green"
conditionalFormat.topBottom.rule = { rank: 1, type: "TopItems"}

await context.sync();
```

## <a name="multiple-formats-and-priority"></a><span data-ttu-id="dcf3c-176">Formats multiples et priorité</span><span class="sxs-lookup"><span data-stu-id="dcf3c-176">Multiple formats and priority</span></span>

<span data-ttu-id="dcf3c-177">Vous pouvez appliquer plusieurs mises en forme conditionnelles à une plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-177">You can apply multiple conditional formats to a range.</span></span> <span data-ttu-id="dcf3c-178">Si les formats ont des éléments en conflit, tels que les couleurs de police différentes, la mise en forme s’applique uniquement à un élément particulier.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-178">If the formats have conflicting elements, such as differing font colors, only one format applies that particular element.</span></span> <span data-ttu-id="dcf3c-179">La priorité est définie par la`ConditionalFormat.priority` propriété.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-179">Precedence is defined by the `ConditionalFormat.priority` property.</span></span> <span data-ttu-id="dcf3c-180">La priorité est un nombre (égal à l’index dans le `ConditionalFormatCollection`) et peut être définie lorsque vous créez le format.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-180">Priority is a number (equal to the index in the `ConditionalFormatCollection`) and can be set when creating the format.</span></span> <span data-ttu-id="dcf3c-181">Plus basse la `priority` valeur est, plus élevée la priorité de la mise en forme est.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-181">The lowerer the `priority` value, the higher the priority of the format is.</span></span>

<span data-ttu-id="dcf3c-182">L’exemple suivant montre un choix de couleur de police en conflit entre les deux formats.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-182">The following example shows a conflicting font color choice between the two formats.</span></span> <span data-ttu-id="dcf3c-183">Les nombres négatifs recevront une police en gras, mais pas une police rouge, car la priorité se porte sur le format leur donnant une police bleue.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-183">Negative numbers will get a bold font, but NOT a red font, because priority goes to the format that gives them a blue font.</span></span>

![Une plage avec un petit nombre en gras et rouge, nombres négatifs en bleu et arrière-plan vert.](../images/excel-conditional-format-priority.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const temperatureDataRange = sheet.tables.getItem("TemperatureTable").getDataBodyRange();


// Set low numbers to bold, dark red font and assign priority 1.
const presetFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.presetCriteria);
presetFormat.preset.format.font.color = "red";
presetFormat.preset.format.font.bold = true;
presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };
presetFormat.priority = 1;

// Set negative numbers to blue font with green background and set priority 0.
const cellValueFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.cellValue);
cellValueFormat.cellValue.format.font.color = "blue";
cellValueFormat.cellValue.format.fill.color = "lightgreen";
cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };
cellValueFormat.priority = 0;

await context.sync();

```

### <a name="mutually-exclusive-conditional-formats"></a><span data-ttu-id="dcf3c-185">Formats exclusifs de mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="dcf3c-185">Mutually exclusive conditional formats</span></span>

<span data-ttu-id="dcf3c-186">La `stopIfTrue` propriété de `ConditionalFormat` empêche les mises en forme conditionnelles de priorité inférieure de s’appliquer à la plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-186">The `stopIfTrue` property of `ConditionalFormat` prevents lower priority conditional formats from being applied to the range.</span></span> <span data-ttu-id="dcf3c-187">Lorsqu’une plage correspondant à la mise en forme conditionnelle avec `stopIfTrue === true` est appliquée, aucune mise en forme conditionnelle suivante n’est appliquée, même si ses détails de mise en forme ne sont pas contradictoires.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-187">When a range matching the conditional format with `stopIfTrue === true` is applied, no subsequent conditional formats are applied, even if their formatting details are not contradictory.</span></span>

<span data-ttu-id="dcf3c-188">L’exemple suivant montre deux mises en forme conditionnelles ajoutées à une plage.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-188">The following example shows two conditional formats being added to a range.</span></span> <span data-ttu-id="dcf3c-189">Les nombres négatifs aura une police bleue avec un arrière-plan vert léger, quelle que soit la condition de l’autre format.</span><span class="sxs-lookup"><span data-stu-id="dcf3c-189">Negative numbers will have a blue font with a light green background, regardless of whether the other format condition is true.</span></span>

![Une plage avec les petits nombres en gras et en rouge, sauf s’ils sont négatifs, auquel cas ils ne sont pas en gras, bleu et ont un arrière-plan vert.](../images/excel-conditional-format-stopiftrue.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const temperatureDataRange = sheet.tables.getItem("TemperatureTable").getDataBodyRange();

// Set low numbers to bold, dark red font and assign priority 1.
const presetFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.presetCriteria);
presetFormat.preset.format.font.color = "red";
presetFormat.preset.format.font.bold = true;
presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };
presetFormat.priority = 1;

// Set negative numbers to blue font with green background and 
// set priority 0, but set stopIfTrue to true, so none of the 
// formatting of the conditional format with the higher priority
// value will apply, not even the bolding of the font.
const cellValueFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.cellValue);
cellValueFormat.cellValue.format.font.color = "blue";
cellValueFormat.cellValue.format.fill.color = "lightgreen";
cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };
cellValueFormat.priority = 0;
cellValueFormat.stopIfTrue = true;

await context.sync();
```

## <a name="see-also"></a><span data-ttu-id="dcf3c-191">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dcf3c-191">See also</span></span>

- [<span data-ttu-id="dcf3c-192">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="dcf3c-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](/office/dev/add-ins/excel/excel-add-ins-core-concepts)
- [<span data-ttu-id="dcf3c-193">Utiliser les plages à l’aide de l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="dcf3c-193">Work with ranges using the Excel JavaScript API</span></span>](/office/dev/add-ins/excel/excel-add-ins-ranges)
- [<span data-ttu-id="dcf3c-194">Objet ConditionalFormat (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="dcf3c-194">ConditionalFormat Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.conditionalformat)
- [<span data-ttu-id="dcf3c-195">Ajouter, modifier ou effacer des formats conditionnels</span><span class="sxs-lookup"><span data-stu-id="dcf3c-195">Add, change, or clear conditional formats</span></span>](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
- [<span data-ttu-id="dcf3c-196">Utilisez des formules avec mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="dcf3c-196">Use formulas with conditional formatting</span></span>](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)
