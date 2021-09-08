---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.6
description: Détails sur l’ensemble de conditions requises ExcelApi 1.6.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9cc7155928fb3f98acf25a38ca5210ca6d57bb19
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938538"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Nouveautés de l’API JavaScript 1.6 pour Excel

## <a name="conditional-formatting"></a>Mise en forme conditionnelle

Introduit la mise en forme conditionnelle d’une plage. Autorise les types suivants de mise en forme conditionnelle.

- Échelle de couleurs
- Barre de données
- Jeu d'icônes
- Personnalisé

De plus :

- Renvoie la plage à laquelle s’applique la mise en forme conditionnelle.
- Supprime la mise en forme conditionnelle.
- Fournit une priorité et des `stopifTrue` fonctionnalités.
- Obtient la collection de toutes les mises en forme conditionnelles sur une plage donnée.
- Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API Excel l’ensemble de conditions requises de l’API JavaScript 1.6. Pour afficher la documentation de référence de l’API pour toutes les API pris en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.6 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.6](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendApiCalculationUntilNextSync__)|Suspend le calcul jusqu’à ce que le `context.sync()` suivant soit appelé.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Spécifie l’objet de règle sur ce format conditionnel.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Critères de l’échelle de couleurs.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threeColorScale)|Si `true` , l’échelle de couleur aura trois points (minimum, milieu, maximum), sinon elle en aura deux (minimum, maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle.|
||[opérateur](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Opérateur de la mise en forme conditionnelle de la valeur de cellule.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Point maximal du critère d’échelle de couleur.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Milieu du critère d’échelle de couleur, si l’échelle de couleurs est une échelle de 3 couleurs.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Point minimal du critère d’échelle de couleur.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Représentation de code couleur HTML de la couleur d’échelle de couleur (par exemple, #FF0000 représente le rouge).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Un nombre, une formule ou `null` (si `type` c’est `lowestValue` le cas).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|Sur quoi la formule conditionnelle critère doit être basée.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#borderColor)|Code couleur HTML représentant la couleur de la bordure, au format #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillColor)|Code couleur HTML représentant la couleur de remplissage, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchPositiveBorderColor)|Spécifie si la barre de données négative a la même couleur de bordure que la barre de données positive.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchPositiveFillColor)|Spécifie si la barre de données négative a la même couleur de remplissage que la barre de données positive.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#borderColor)|Code couleur HTML représentant la couleur de la bordure, au format #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillColor)|Code couleur HTML représentant la couleur de remplissage, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientFill)|Spécifie si la barre de données a un dégradé.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Formule, si nécessaire, sur laquelle évaluer la règle de barre de données.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Type de règle pour la barre de données.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete__)|Supprime cette mise en forme conditionnelle.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getRange__)|Renvoie la plage à laquelle s’applique la mise en forme conditionnelle.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getRangeOrNullObject__)|Renvoie la plage à laquelle le format conditionnel est appliqué.|
||[priority](/javascript/api/excel/excel.conditionalformat#priority)|Priorité (ou index) dans la collection de formats conditionnels dans qui ce format conditionnel existe actuellement.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellValue)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de cellule si la mise en forme conditionnelle actuelle est un `CellValue` type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellValueOrNullObject)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de cellule si la mise en forme conditionnelle actuelle est un `CellValue` type.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorScale)|Renvoie les propriétés de mise en forme conditionnelle d’échelle de couleur si la mise en forme conditionnelle actuelle est un `ColorScale` type.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorScaleOrNullObject)|Renvoie les propriétés de mise en forme conditionnelle d’échelle de couleur si la mise en forme conditionnelle actuelle est un `ColorScale` type.|
||[custom](/javascript/api/excel/excel.conditionalformat#custom)|Renvoie les propriétés de mise en forme conditionnelle personnalisées si la mise en forme conditionnelle actuelle est un type personnalisé.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customOrNullObject)|Renvoie les propriétés de mise en forme conditionnelle personnalisées si la mise en forme conditionnelle actuelle est un type personnalisé.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#dataBar)|Renvoie les propriétés de la barre de données si la mise en forme conditionnelle actuelle est une barre de données.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#dataBarOrNullObject)|Renvoie les propriétés de la barre de données si la mise en forme conditionnelle actuelle est une barre de données.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconSet)|Renvoie les propriétés de mise en forme conditionnelle du jeu d’icônes si la mise en forme conditionnelle actuelle est un `IconSet` type.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconSetOrNullObject)|Renvoie les propriétés de mise en forme conditionnelle du jeu d’icônes si la mise en forme conditionnelle actuelle est un `IconSet` type.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|Priorité de la mise en forme conditionnelle dans la version `ConditionalFormatCollection` actuelle.|
||[preset](/javascript/api/excel/excel.conditionalformat#preset)|Renvoie la mise en forme conditionnelle des critères prédéfinits.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetOrNullObject)|Renvoie la mise en forme conditionnelle des critères prédéfinits.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textComparison)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si la mise en forme conditionnelle actuelle est un type de texte.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textComparisonOrNullObject)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si la mise en forme conditionnelle actuelle est un type de texte.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topBottom)|Renvoie les propriétés de mise en forme conditionnelle supérieure/inférieure si la mise en forme conditionnelle actuelle est un `TopBottom` type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topBottomOrNullObject)|Renvoie les propriétés de mise en forme conditionnelle supérieure/inférieure si la mise en forme conditionnelle actuelle est un `TopBottom` type.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Type de mise en forme conditionnelle.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopIfTrue)|Si les conditions de cette mise en forme conditionnelle sont remplies, aucun format de priorité inférieure ne doit prendre effet sur cette cellule.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add_type_)|Ajoute un nouveau format conditionnel à la collection à la priorité première/supérieure.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearAll__)|Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getCount__)|Renvoie le nombre de formats conditionnels dans le manuel.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItem_id_)|Renvoie une mise en forme conditionnelle à un ID donné.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getItemAt_index_)|Renvoie une mise en forme conditionnelle à l’index donné.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulaLocal)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle dans la langue de l’utilisateur.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formulaR1C1)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle en notation de style R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customIcon)|L’icône personnalisée pour le critère actuel, si elle est différente du jeu d’icônes par défaut, `null` est renvoyée.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Un nombre ou une formule en fonction du type.|
||[opérateur](/javascript/api/excel/excel.conditionaliconcriterion#operator)|`greaterThan` ou `greaterThanOrEqual` pour chacun des types de règles pour la mise en forme conditionnelle d’icône.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|Ce sur quoi la formule conditionnelle de l’icône doit être basée.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[critère](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Critère de la mise en forme conditionnelle.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|Code couleur HTML représentant la couleur de la bordure, au format #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideIndex)|Valeur constante qui indique un côté spécifique de la bordure.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getItem_index_)|Obtient un objet de bordure à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getItemAt_index_)|Obtient un objet de bordure à l’aide de son indice.|
||[bas](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Obtient la bordure inférieure.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Nombre d’objets de bordure de la collection.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Obtient la bordure gauche.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Obtient la bordure droite.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Obtient la bordure supérieure.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear__)|Réinitialise le remplissage.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|Code couleur HTML représentant la couleur du remplissage, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Spécifie si la police est en gras.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear__)|Réinitialise les formats de police.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Représentation de code couleur HTML de la couleur du texte (par exemple, #FF0000 représente le rouge).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Spécifie si la police est en italique.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Spécifie l’état de la police de type strikethrough.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Type de soulignement appliqué à la police.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberFormat)|Représente le Excel de format numérique de la plage donnée.|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Collection d’objets de bordure qui s’appliquent à la plage de mise en forme conditionnelle globale.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Renvoie l’objet de remplissage défini sur la plage de mise en forme conditionnelle globale.|
||[police](/javascript/api/excel/excel.conditionalrangeformat#font)|Renvoie l’objet de police défini sur la plage de mise en forme conditionnelle globale.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[opérateur](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Opérateur de la mise en forme conditionnelle du texte.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Valeur de texte de la mise en forme conditionnelle.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|Rang compris entre 1 et 1000 pour les rangs numériques ou entre 1 et 100 pour les rangs en pourcentage.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Formater les valeurs en fonction du classement supérieur ou inférieur.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.customconditionalformat#rule)|Spécifie `Rule` l’objet sur cette mise en forme conditionnelle.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axisColor)|Code couleur HTML représentant la couleur de la ligne Axe, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisFormat)|Représentation de la façon dont l’axe est déterminé pour une barre Excel données.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#barDirection)|Spécifie la direction sur qui le graphique de barre de données doit être basé.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerBoundRule)|Règle de ce qui constitue la limite inférieure (et comment la calculer, le cas échéant) pour une barre de données.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeFormat)|Représentation de toutes les valeurs à gauche de l’axe dans une barre Excel données.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveFormat)|Représentation de toutes les valeurs à droite de l’axe dans une barre Excel données.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showDataBarOnly)|Si `true` , masque les valeurs des cellules où la barre de données est appliquée.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperBoundRule)|Règle de ce qui constitue la limite supérieure (et comment la calculer, le cas échéant) pour une barre de données.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Tableau de critères et d’ensembles d’icônes pour les règles et icônes personnalisées potentielles pour les icônes conditionnelles.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseIconOrder)|Si `true` , inverse les commandes d’icône pour le jeu d’icônes.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showIconOnly)|Si `true` , masque les valeurs et affiche uniquement les icônes.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Si elle est définie, affiche l’option de jeu d’icônes pour la mise en forme conditionnelle.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Règle de mise en forme conditionnelle.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate__)|Calcule une plage de cellules dans une feuille de calcul.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalFormats)|Collection de cette `ConditionalFormats` plage.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés de la mise en forme conditionnelle.|
||[rule](/javascript/api/excel/excel.textconditionalformat#rule)|Règle de mise en forme conditionnelle.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés de la mise en forme conditionnelle.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Critères de la mise en forme conditionnelle supérieure/inférieure.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate_markAllDirty_)|Calcule toutes les cellules d’une feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
