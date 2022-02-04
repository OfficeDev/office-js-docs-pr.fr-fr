---
title: Appliquer la mise en forme conditionnelle aux plages avec l’API Excel JavaScript
description: Découvrez la mise en forme conditionnelle dans le contexte des Excel javaScript.
ms.date: 04/01/2021
ms.localizationpriority: medium
---

# <a name="apply-conditional-formatting-to-excel-ranges"></a>Appliquer une mise en forme conditionnelle à des plages Excel

La bibliothèque JavaScript Excel fournit des API pour appliquer une mise en forme conditionnelle aux plages de données dans vos feuilles de calcul. Cette fonctionnalité simplifie l’analyse visuelle de grands ensembles de données. La mise en forme effectue également des mises à jour dynamiques en fonction des changements dans la plage.

> [!NOTE]
> Cet article décrit la mise en forme conditionnelle dans le contexte de compléments Excel JavaScript. Les articles suivants offrent des informations détaillées sur les fonctionnalités de mise en forme conditionnelles complètes dans Excel.
> -  [Ajouter, modifier ou effacer des formats conditionnels](https://support.microsoft.com/office/fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
> -  [Utilisez des formules avec mise en forme conditionnelle](https://support.microsoft.com/office/fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)

## <a name="programmatic-control-of-conditional-formatting"></a>Contrôle par programme de mise en forme conditionnelle

La `Range.conditionalFormats` propriété est un ensemble d’objets [ConditionalFormat](/javascript/api/excel/excel.conditionalformat)qui s’appliquent à la plage.  L’`ConditionalFormat` objet contient plusieurs propriétés qui définissent le format à appliquer en fonction du [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).

-    `cellValue`
-    `colorScale`
-    `custom`
-    `dataBar`
-    `iconSet`
-    `preset`
-    `textComparison`
-    `topBottom`

> [!NOTE]
> Chacune de ces propriétés de mise en forme a une variante correspondante`*OrNullObject`. En savoir plus sur ce modèle dans la [\*section Méthodes OrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) .

Un seul type de format peut être défini pour l’objet ConditionalFormat. Cela est déterminé par la `type` propriété, c'est-à-dire une [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype) valeur enum. `type` est défini lorsque vous ajoutez une mise en forme conditionnelle à une plage.

## <a name="creating-conditional-formatting-rules"></a>Modification des règles de mise en forme conditionnelles

Les mises en forme conditionnelles sont ajoutées à une plage à l’aide de `conditionalFormats.add`. Une fois ajoutées, vous pouvez définir les propriétés spécifiques à la mise en forme conditionnelle. Les exemples ci-dessous montrent la création de différents types de mise en forme.

### <a name="cell-value"></a>[Valeur de la cellule](/javascript/api/excel/excel.cellvalueconditionalformat)

La mise en forme conditionnelle de valeur de la cellule applique un format défini par l’utilisateur en fonction des résultats d’une ou deux formules dans la [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule). La`operator` propriété est un [ConditionalCellValueOperator](/javascript/api/excel/excel.conditionalcellvalueoperator) définissant comment les expressions qui en résultent sont liées à la mise en forme.

L’exemple suivant montre une coloration de la police en rouge appliquée à une valeur dans la plage inférieure à zéro.

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

### <a name="color-scale"></a>[Échelle de couleur](/javascript/api/excel/excel.colorscaleconditionalformat)

La mise en forme conditionnelle de l’échelle de couleur applique un dégradé de couleur au sein de la plage de données. La`criteria` propriété sur le `ColorScaleConditionalFormat` définit trois[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`et éventuellement, `midpoint`. Les critères des points d’échelle ont trois propriétés :

-    `color` -Le code de couleur HTML pour le point de terminaison.
-    `formula` -Un nombre ou une formule représentant le point de terminaison. Il s’agit de `null` si `type` est `lowestValue` ou `highestValue`.
-    `type` -Comment la formule doit être évaluée. `highestValue` et `lowestValue` font référence à des valeurs dans la plage en cours de mise en forme.

L’exemple suivant montre une plage colorée de bleue à jaune à rouge. Notez que `minimum` et `maximum` sont les valeurs inférieures et supérieures respectivement et utilisent les `null` formules. `midpoint` utilise le `percentage` type avec une formule de `"=50"` donc la cellule jaune est la valeur moyenne.

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

### <a name="custom"></a>[Personnalisé](/javascript/api/excel/excel.customconditionalformat)

La mise en forme conditionnelle personnalisée applique un format défini par l’utilisateur aux cellules en fonction d’une formule de complexité arbitraire. L’objet [ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule) vous permet de définir la formule dans des notations différentes :

-    `formula` -Notation standard.
-    `formulaLocal` - Localisée en fonction de la langue de l’utilisateur.
-    `formulaR1C1` -Notation type L1C1.

L’exemple suivant colore les polices de cellules avec des valeurs supérieures à la cellule située à leur gauche en vert.

![Une plage avec des nombres verts place la valeur de la colonne précédente dans cette ligne comme inférieure.](../images/excel-conditional-format-custom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.custom
);

// if a cell has a higher value than the one to its left, set that cell's font to green
conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT("RC[-1]",0),TRUE)';
conditionalFormat.custom.format.font.color = "green";

await context.sync();

```
### <a name="data-bar"></a>[Barre de données](/javascript/api/excel/excel.databarconditionalformat)

La mise en forme conditionnelle de la barre de données ajoute des barres de données aux cellules. Par défaut, les valeurs minimales et maximales dans la plage forment les limites et les tailles proportionnelles des barres de données. L’objet `DataBarConditionalFormat` possède plusieurs propriétés pour contrôler l’apparence de la barre. 

L’exemple suivant met en forme la plage contenant des barres de données remplissant de gauche à droite.

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

### <a name="icon-set"></a>[Jeu d’icônes](/javascript/api/excel/excel.iconsetconditionalformat)

La mise en forme conditionnelle du jeu d’icônes utilise Excel [icônes](/javascript/api/excel/excel.icon) pour mettre en surbrillance les cellules. La `criteria` propriété est une matrice de [ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion), qui définit le symbole à insérer et la condition sous laquelle celui-ci est inséré. Ce tableau est automatiquement pré-rempli avec éléments critères avec les propriétés par défaut. Les propriétés individuelles ne peut pas être remplacées. Au lieu de cela, l’ensemble de l’objet de critères doit être remplacé. 

L’exemple suivant montre un jeu d’icônes trois triangles utilisé dans la plage.

![Plage avec triangles verts vers le haut pour les valeurs supérieures à 1 000, lignes jaunes pour les valeurs entre 700 et 1 000 et triangles rouges vers le bas pour les valeurs inférieures.](../images/excel-conditional-format-iconset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.iconSet
);

const iconSetCF = conditionalFormat.iconSet;
iconSetCF.style = Excel.IconSet.threeTriangles;

/*
   With a "three*" icon set style, such as "threeTriangles", the third
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

### <a name="preset-criteria"></a>[Critères prédéfinis](/javascript/api/excel/excel.presetcriteriaconditionalformat)

La mise en forme conditionnelle prédéfinie applique un format défini par l’utilisateur pour la plage basée sur une règle standard sélectionnée. Ces règles sont définies par le [ConditionalFormatPresetCriterion](/javascript/api/excel/excel.conditionalformatpresetcriterion) dans le [ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule). 

L’exemple suivant colore la police en blanc lorsque la valeur d’une cellule est au moins un écart-type au-dessus de la moyenne de la plage.

![Une plage de cellules avec police en blanc où les valeurs sont au moins un écart-type standard au-dessus de la moyenne.](../images/excel-conditional-format-preset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.presetCriteria
);

// color every cell's font white that is one standard deviation above average relative to the range
conditionalFormat.preset.format.font.color = "white";
conditionalFormat.preset.rule = {
     criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage
};

await context.sync();
```

### <a name="text-comparison"></a>[Comparaison de texte](/javascript/api/excel/excel.textconditionalformat)

La mise en forme conditionnelle de comparaison de texte utilise des comparaisons de chaînes comme condition. La `rule` propriété est un [ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule) définissant une chaîne à comparer avec la cellule et un opérateur pour spécifier le type de comparaison. 

L’exemple suivant formate la couleur rouge de la police lorsque le texte d’une cellule contient « Delayed ».

![Une plage de cellules contenant « Différé » en rouge.](../images/excel-conditional-format-text.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B16:D18");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.containsText
);

// color the font of every cell containing "Delayed"
conditionalFormat.textComparison.format.font.color = "red";
conditionalFormat.textComparison.rule = {
     operator: Excel.ConditionalTextOperator.contains,
     text: "Delayed"
};

await context.sync();
```

### <a name="topbottom"></a>[Supérieure/inférieure](/javascript/api/excel/excel.topbottomconditionalformat)

La mise en forme conditionnelle supérieure/inférieure applique un format aux valeurs les plus élevées ou plus faibles d’une plage. La `rule` propriété, de type [ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule), définit si la condition est basée sur le plus élevé ou le plus bas, ainsi que si l’évaluation est en classement ou pourcentage. 

L’exemple suivant applique un surlignage vert à la cellule de valeur plus élevée dans la plage.


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

## <a name="multiple-formats-and-priority"></a>Formats multiples et priorité

Vous pouvez appliquer plusieurs mises en forme conditionnelles à une plage. Si les formats ont des éléments en conflit, tels que les couleurs de police différentes, la mise en forme s’applique uniquement à un élément particulier. La priorité est définie par la`ConditionalFormat.priority` propriété. La priorité est un nombre (égal à l’index dans le `ConditionalFormatCollection`) et peut être définie lorsque vous créez le format. Plus basse la `priority` valeur est, plus élevée la priorité de la mise en forme est.

L’exemple suivant montre un choix de couleur de police en conflit entre les deux formats. Les nombres négatifs recevront une police en gras, mais pas une police rouge, car la priorité se porte sur le format leur donnant une police bleue.

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

### <a name="mutually-exclusive-conditional-formats"></a>Formats exclusifs de mise en forme conditionnelle

La `stopIfTrue` propriété de `ConditionalFormat` empêche les mises en forme conditionnelles de priorité inférieure de s’appliquer à la plage. Lorsqu’une plage correspondant à la mise en forme conditionnelle avec `stopIfTrue === true` est appliquée, aucune mise en forme conditionnelle suivante n’est appliquée, même si ses détails de mise en forme ne sont pas contradictoires.

L’exemple suivant montre deux mises en forme conditionnelles ajoutées à une plage. Les nombres négatifs aura une police bleue avec un arrière-plan vert léger, quelle que soit la condition de l’autre format.

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

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](../excel/excel-add-ins-core-concepts.md)
- [Objet ConditionalFormat (API JavaScript pour Excel)](/javascript/api/excel/excel.conditionalformat)
- [Ajouter, modifier ou effacer des formats conditionnels](https://support.microsoft.com/office/fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
- [Utilisez des formules avec mise en forme conditionnelle](https://support.microsoft.com/office/fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
