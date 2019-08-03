---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,6
description: Détails sur l’ensemble de conditions requises ExcelApi 1,6
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c38dd942c3002af05f847846145bc89f1cbbccbe
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064906"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Nouveautés de l’API JavaScript 1.6 pour Excel

## <a name="conditional-formatting"></a>Mise en forme conditionnelle

Présente la mise en forme conditionnelle d’une plage. Autorise les types de mise en forme conditionnelle suivants :

* Échelle de couleurs
* Barre de données
* Jeu d'icônes
* Personnalisé

De plus :

* Renvoie la plage à laquelle s’applique la mise en forme conditionnelle.
* Supprime la mise en forme conditionnelle.
* Fournit la priorité `stopifTrue` et la capacité.
* Obtient la collection de toutes les mises en forme conditionnelles sur une plage donnée.
* Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Excel 1,6. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1,6 ou antérieure, voir [API Excel dans l’ensemble de conditions requises 1,6 ou version antérieure](/javascript/api/excel?view=excel-js-1.6).

| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Interrompt le calcul jusqu'à ce que la prochaine méthode « context.sync() » soit appelée. Une fois cette option définie, il incombe au développeur de recalculer le classeur afin de garantir que toutes les dépendances sont propagées.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[sous](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Représente l’objet Règle sur cette mise en forme conditionnelle.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Les critères de l’étendue de couleur. Le milieu est facultatif lors de l’utilisation d’une graduation de couleurs à deux points.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Si la valeur est true, l’échelle de couleurs aura trois points (minimum, milieu, maximum), sinon elle aura deux (minimum, maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[Formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|
||[Formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|
||[is](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Opérateur du format conditionnel de texte.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Point maximal du critère d’échelle de couleurs.|
||[point](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Point du milieu du critère d’échelle de couleurs, si l’échelle de couleurs est une échelle à 3 couleurs.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Point minimal du critère d’échelle de couleurs.|
|[Troisconditionalcolorscalecriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Représentation de code de couleur HTML de la couleur d’image. Par exemple, #FF0000 représente le rouge.|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Nombre, formule ou null (si le type est LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|La formule conditionnelle de critère qui doit être basée.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Représentation booléenne indiquant si la barre de données négative a une bordure de la même couleur que la barre de données positive.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Représentation booléenne indiquant si la barre de données négative a un remplissage de la même couleur que la barre de données positive.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Représentation booléenne indiquant si la barre de données a un dégradé.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Formule, si nécessaire, servant à évaluer la règle de la barre de données.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Type de règle pour le DataBar.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Supprime cette mise en forme conditionnelle.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Renvoie la plage à laquelle s’applique la mise en forme conditionnelle. Renvoie une erreur si la mise en forme conditionnelle est appliquée à plusieurs plages. En lecture seule.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Renvoie la plage à laquelle le format conditionnelle est appliqué ou un objet null si la mise en forme conditionnelle est appliquée à plusieurs plages. En lecture seule.|
||[prioritaires](/javascript/api/excel/excel.conditionalformat#priority)|Priorité (ou index) dans la collection de mise en forme conditionnelle dans laquelle ce format conditionnel existe actuellement. Modification également|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de la cellule si le format conditionnel actuel est un type CellValue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de la cellule si le format conditionnel actuel est un type CellValue.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Renvoie les propriétés de mise en forme conditionnelle ColorScale si le format conditionnel actuel est un type ColorScale. En lecture seule.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Renvoie les propriétés de mise en forme conditionnelle ColorScale si le format conditionnel actuel est un type ColorScale. En lecture seule.|
||[personnalisé](/javascript/api/excel/excel.conditionalformat#custom)|Renvoie les propriétés de mise en forme conditionnelle personnalisées si le format conditionnel actuel est un type personnalisé. En lecture seule.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Renvoie les propriétés de mise en forme conditionnelle personnalisées si le format conditionnel actuel est un type personnalisé. En lecture seule.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Renvoie les propriétés de la barre de données si le format conditionnel actuel est une barre de données. En lecture seule.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Renvoie les propriétés de la barre de données si le format conditionnel actuel est une barre de données. En lecture seule.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Renvoie les propriétés de format conditionnel IconSet si le format conditionnel actuel est un type IconSet. En lecture seule.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Renvoie les propriétés de format conditionnel IconSet si le format conditionnel actuel est un type IconSet. En lecture seule.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|La priorité de la mise en forme conditionnelle dans la ConditionalFormatCollection actuelle. En lecture seule.|
||[définie](/javascript/api/excel/excel.conditionalformat#preset)|Renvoie le format conditionnel des critères prédéfinis. Pour plus d’informations, voir Excel. PresetCriteriaConditionalFormat.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Renvoie le format conditionnel des critères prédéfinis. Pour plus d’informations, voir Excel. PresetCriteriaConditionalFormat.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si le format conditionnel actuel est un type de texte.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si le format conditionnel actuel est un type de texte.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Renvoie les propriétés de mise en forme conditionnelle de haut en bas si le format conditionnel actuel est un type de niveau inférieur.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Renvoie les propriétés de mise en forme conditionnelle de haut en bas si le format conditionnel actuel est un type de niveau inférieur.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Type de mise en forme conditionnelle. Une seule peut être définie à la fois. En lecture seule.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Si les conditions de cette mise en forme conditionnelle sont remplies, aucun format de priorité inférieure ne doit prendre effet sur cette cellule.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[Add (type: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Ajoute un nouveau format conditionnel à la collection à la priorité la plus haute.|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Renvoie le nombre de mises en forme conditionnelles dans le classeur. En lecture seule.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Renvoie une mise en forme conditionnelle à un ID donné.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Renvoie une mise en forme conditionnelle à l’index donné.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|
||[FormulaLocal,](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle dans la langue de l’utilisateur.|
||[Formular1c1,](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle dans la notation du style R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Icône personnalisée pour le critère en cours si différent de la celui par défaut IconSet. Sinon, null est renvoyé.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Un nombre ou une formule en fonction du type.|
||[is](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan ou GreaterThanOrEqual pour chaque type de règle pour le format conditionnel d’icône.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|Ce sur quoi la formule conditionnelle de l’icône doit être basée.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[critère](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Critère de la mise en forme conditionnelle.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Valeur constante qui indique un côté spécifique de la bordure. Pour plus d’informations, voir Excel. ConditionalRangeBorderIndex. En lecture seule.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure. Pour plus d’informations, voir Excel. BorderLineStyle.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Obtient un objet de bordure à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Obtient un objet de bordure à l’aide de son indice.|
||[bas](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Obtient la bordure inférieure. En lecture seule.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Nombre d’objets de bordure de la collection. En lecture seule.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Obtient la bordure gauche. En lecture seule.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Obtient la bordure droite. En lecture seule.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Obtient la bordure supérieure. En lecture seule.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Réinitialise le remplissage.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Représente le format de police Gras.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Réinitialise les formats de police.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Représente le format de police Italique.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Représente l’état barré de la police.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. ConditionalRangeFontUnderlineStyle.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Représente le code de format de nombre d’Excel pour la plage donnée. Effacé si null est passé dans.|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Collection d’objets Border qui s’appliquent à la plage de mise en forme conditionnelle globale. En lecture seule.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Renvoie l’objet Fill défini sur la plage de mise en forme conditionnelle globale. En lecture seule.|
||[police](/javascript/api/excel/excel.conditionalrangeformat#font)|Renvoie l’objet font défini sur la plage de mise en forme conditionnelle globale. En lecture seule.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[is](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Opérateur du format conditionnel de texte.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Valeur de texte de la mise en forme conditionnelle.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|Rang compris entre 1 et 1000 pour les rangs numériques ou entre 1 et 100 pour les rangs en pourcentage.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Mettre en forme les valeurs en fonction du rang supérieur ou inférieur.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels. En lecture seule.|
||[sous](/javascript/api/excel/excel.customconditionalformat#rule)|Représente l’objet Règle sur cette mise en forme conditionnelle. En lecture seule.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[Axiscolor,](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|Code couleur HTML qui représente la couleur de la ligne Axe, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Représentation de la façon dont l’axe est déterminé pour une barre de données Excel.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Représente le sens de l’image de la barre de données.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Règle de ce qui constitue la limite inférieure (et comment la calculer, le cas échéant) pour une barre de données.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Représentation de toutes les valeurs à gauche de l’axe dans une barre de données Excel. En lecture seule.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Représentation de toutes les valeurs à droite de l’axe dans une barre de données Excel. En lecture seule.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Si la valeur est True, masque les valeurs des cellules où la barre de données est appliquée.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Règle de ce qui constitue la limite supérieure (et comment la calculer, le cas échéant) pour une barre de données.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Un tableau de critères et de IconSets pour les règles et les icônes personnalisées potentielles pour les icônes conditionnelles. Notez que pour le premier critère, seule l’icône personnalisée peut être modifiée, tandis que le type, la formule et l’opérateur seront ignorés lors de la définition.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Si la valeur est true, inverse l’ordre des icônes pour la IconSet. Notez que cette valeur ne peut pas être définie si des icônes personnalisées sont utilisées.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Si la valeur est True, masque les valeurs et affiche uniquement les icônes.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Si ce paramètre est défini, il affiche l’option IconSet pour le format conditionnel.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[sous](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Règle de mise en forme conditionnelle.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Calcule une plage de cellules dans une feuille de calcul.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Collection d’ConditionalFormats qui croisent la plage. En lecture seule.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels. En lecture seule.|
||[sous](/javascript/api/excel/excel.textconditionalformat#rule)|Règle de mise en forme conditionnelle.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels. En lecture seule.|
||[sous](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Critères du format conditionnel le plus haut/bas.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Calculate (markAllDirty: booléen)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Calcule toutes les cellules d’une feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.6)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
