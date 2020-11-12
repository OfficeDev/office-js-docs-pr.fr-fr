---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,6
description: Détails sur l’ensemble de conditions requises ExcelApi 1,6.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 20fe6950db2661d08969bdc4f2b7dc6fa5ad7a97
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996213"
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
* Fournit la priorité et la `stopifTrue` capacité.
* Obtient la collection de toutes les mises en forme conditionnelles sur une plage donnée.
* Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Excel 1,6. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1,6 ou antérieure, voir [API Excel dans l’ensemble de conditions requises 1,6 ou version antérieure](/javascript/api/excel?view=excel-js-1.6&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Interrompt le calcul jusqu'à ce que la prochaine méthode « context.sync() » soit appelée.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[sous](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Spécifie l’objet Rule sur ce format conditionnel.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Les critères de l’étendue de couleur.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Si la valeur est true, l’échelle de couleurs aura trois points (minimum, milieu, maximum), sinon elle aura deux (minimum, maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[Formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|
||[Formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|
||[opérateur](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Opérateur du format conditionnel de la valeur de cellule.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Point maximal du critère d’échelle de couleurs.|
||[point](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Point du milieu du critère d’échelle de couleurs, si l’échelle de couleurs est une échelle à 3 couleurs.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Point minimal du critère d’échelle de couleurs.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Représentation de code de couleur HTML de la couleur d’une image d’une couleur (par exemple, #FF0000 représente le rouge).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Nombre, formule ou null (si le type est LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|La formule conditionnelle de critère qui doit être basée.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|Code couleur HTML représentant la couleur de remplissage, de la #RRGGBB de formulaire (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Indique si le DataBar négatif a la même couleur de bordure que le DataBar positif.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Indique si le DataBar négatif a la même couleur de remplissage que le DataBar positif.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|Code couleur HTML représentant la couleur de remplissage, de la #RRGGBB de formulaire (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Indique si le DataBar a un dégradé.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Formule, si nécessaire, servant à évaluer la règle de la barre de données.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Type de règle pour le DataBar.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Supprime cette mise en forme conditionnelle.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Renvoie la plage à laquelle s’applique la mise en forme conditionnelle.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Renvoie la plage à laquelle le format conditionnelle est appliqué ou un objet null si la mise en forme conditionnelle est appliquée à plusieurs plages.|
||[priorité](/javascript/api/excel/excel.conditionalformat#priority)|Priorité (ou index) dans la collection de mise en forme conditionnelle dans laquelle ce format conditionnel existe actuellement.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de la cellule si le format conditionnel actuel est un type CellValue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de la cellule si le format conditionnel actuel est un type CellValue.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Renvoie les propriétés de mise en forme conditionnelle ColorScale si le format conditionnel actuel est un type ColorScale.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Renvoie les propriétés de mise en forme conditionnelle ColorScale si le format conditionnel actuel est un type ColorScale.|
||[personnalisé](/javascript/api/excel/excel.conditionalformat#custom)|Renvoie les propriétés de mise en forme conditionnelle personnalisées si le format conditionnel actuel est un type personnalisé.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Renvoie les propriétés de mise en forme conditionnelle personnalisées si le format conditionnel actuel est un type personnalisé.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Renvoie les propriétés de la barre de données si le format conditionnel actuel est une barre de données.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Renvoie les propriétés de la barre de données si le format conditionnel actuel est une barre de données.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Renvoie les propriétés de format conditionnel IconSet si le format conditionnel actuel est un type IconSet.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Renvoie les propriétés de format conditionnel IconSet si le format conditionnel actuel est un type IconSet.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|La priorité de la mise en forme conditionnelle dans la ConditionalFormatCollection actuelle.|
||[définie](/javascript/api/excel/excel.conditionalformat#preset)|Renvoie le format conditionnel des critères prédéfinis.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Renvoie le format conditionnel des critères prédéfinis.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si le format conditionnel actuel est un type de texte.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si le format conditionnel actuel est un type de texte.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Renvoie les propriétés de mise en forme conditionnelle de haut en bas si le format conditionnel actuel est un type de niveau inférieur.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Renvoie les propriétés de mise en forme conditionnelle de haut en bas si le format conditionnel actuel est un type de niveau inférieur.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Type de mise en forme conditionnelle.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Si les conditions de cette mise en forme conditionnelle sont remplies, aucun format de priorité inférieure ne doit prendre effet sur cette cellule.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[Add (type : Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Ajoute un nouveau format conditionnel à la collection à la priorité la plus haute.|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Renvoie le nombre de mises en forme conditionnelles dans le classeur.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Renvoie une mise en forme conditionnelle à un ID donné.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Renvoie une mise en forme conditionnelle à l’index donné.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|
||[FormulaLocal,](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle dans la langue de l’utilisateur.|
||[Formular1c1,](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle dans la notation du style R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Icône personnalisée pour le critère en cours si différent de la celui par défaut IconSet. Sinon, null est renvoyé.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Un nombre ou une formule en fonction du type.|
||[opérateur](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan ou GreaterThanOrEqual pour chaque type de règle pour le format conditionnel d’icône.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|Ce sur quoi la formule conditionnelle de l’icône doit être basée.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[critère](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Critère de la mise en forme conditionnelle.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Valeur constante qui indique un côté spécifique de la bordure.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (index : Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Obtient un objet de bordure à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Obtient un objet de bordure à l’aide de son indice.|
||[bas](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Obtient la bordure inférieure.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Nombre d’objets de bordure de la collection.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Obtient la bordure gauche.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Obtient la bordure droite.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Obtient la bordure supérieure.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Réinitialise le remplissage.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|Code couleur HTML qui représente la couleur du remplissage, de la #RRGGBB de formulaire (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Indique si la police est en gras.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Réinitialise les formats de police.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Représentation de code de couleur HTML de la couleur du texte (par exemple, #FF0000 représente le rouge).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Indique si la police est en italique.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Spécifie l’État barré de la police.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Type de soulignement appliqué à la police.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Représente le code de format de nombre d’Excel pour la plage donnée.|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Collection d’objets Border qui s’appliquent à la plage de mise en forme conditionnelle globale.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Renvoie l’objet Fill défini sur la plage de mise en forme conditionnelle globale.|
||[police](/javascript/api/excel/excel.conditionalrangeformat#font)|Renvoie l’objet font défini sur la plage de mise en forme conditionnelle globale.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[opérateur](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Opérateur du format conditionnel de texte.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Valeur de texte de la mise en forme conditionnelle.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|Rang compris entre 1 et 1000 pour les rangs numériques ou entre 1 et 100 pour les rangs en pourcentage.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Mettre en forme les valeurs en fonction du rang supérieur ou inférieur.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[sous](/javascript/api/excel/excel.customconditionalformat#rule)|Spécifie l’objet Rule sur ce format conditionnel.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[Axiscolor,](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|Code couleur HTML qui représente la couleur de la ligne de l’axe, de la #RRGGBB de formulaire (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Représentation de la façon dont l’axe est déterminé pour une barre de données Excel.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Indique le sens de l’image de la barre de données.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Règle de ce qui constitue la limite inférieure (et comment la calculer, le cas échéant) pour une barre de données.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Représentation de toutes les valeurs à gauche de l’axe dans une barre de données Excel.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Représentation de toutes les valeurs à droite de l’axe dans une barre de données Excel.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Si la valeur est True, masque les valeurs des cellules où la barre de données est appliquée.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Règle de ce qui constitue la limite supérieure (et comment la calculer, le cas échéant) pour une barre de données.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Un tableau de critères et de IconSets pour les règles et les icônes personnalisées potentielles pour les icônes conditionnelles.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Si la valeur est true, inverse l’ordre des icônes pour la IconSet.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Si la valeur est True, masque les valeurs et affiche uniquement les icônes.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Si ce paramètre est défini, il affiche l’option IconSet pour le format conditionnel.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[sous](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Règle de mise en forme conditionnelle.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Calcule une plage de cellules dans une feuille de calcul.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Collection d’ConditionalFormats qui croise la plage.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[sous](/javascript/api/excel/excel.textconditionalformat#rule)|Règle de mise en forme conditionnelle.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[sous](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Critères du format conditionnel le plus haut/bas.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Calculate (markAllDirty : booléen)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Calcule toutes les cellules d’une feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
