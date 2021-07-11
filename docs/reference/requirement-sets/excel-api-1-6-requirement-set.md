---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.6
description: Détails sur l’ensemble de conditions requises ExcelApi 1.6.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: bc2eb8f182a329808a46f172868b818027f5e367
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350105"
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

Le tableau suivant répertorie les API Excel l’ensemble de conditions requises de l’API JavaScript 1.6. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.6 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.6](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Interrompt le calcul jusqu'à ce que la prochaine méthode « context.sync() » soit appelée.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Spécifie l’objet Rule sur cette mise en forme conditionnelle.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Critères de l’échelle de couleurs.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Si la valeur est True, l’échelle de couleur aura trois points (minimum, milieu, maximum), sinon elle en aura deux (minimum, maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|
||[opérateur](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Opérateur de la mise en forme conditionnelle de la valeur de cellule.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Point maximal du critère d’échelle de couleurs.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Point du milieu du critère d’échelle de couleurs, si l’échelle de couleurs est une échelle à 3 couleurs.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Point minimal du critère d’échelle de couleurs.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Représentation de code couleur HTML de la couleur d’échelle de couleur (par exemple, #FF0000 représente le rouge).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Nombre, formule ou null (si le type est LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|Sur quoi la formule conditionnelle critère doit être basée.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|Code couleur HTML représentant la couleur de remplissage du formulaire #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Spécifie si la barre de données négative a la même couleur de bordure que la barre de données positive.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Spécifie si la barre de données négative a la même couleur de remplissage que la barre de données positive.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|Code couleur HTML représentant la couleur de remplissage du formulaire #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Spécifie si la barre de données possède un dégradé.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Formule, si nécessaire, servant à évaluer la règle de la barre de données.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Type de règle pour la barre de données.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Supprime cette mise en forme conditionnelle.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Renvoie la plage à laquelle s’applique la mise en forme conditionnelle.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Renvoie la plage à la mise en forme conditionnelle est appliquée ou un objet null si la mise en forme conditionnelle est appliquée à plusieurs plages.|
||[priority](/javascript/api/excel/excel.conditionalformat#priority)|Priorité (ou index) dans la collection de formats conditionnels dans qui ce format conditionnel existe actuellement.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de cellule si la mise en forme conditionnelle actuelle est de type CellValue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de cellule si la mise en forme conditionnelle actuelle est de type CellValue.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Renvoie les propriétés de mise en forme conditionnelle ColorScale si la mise en forme conditionnelle actuelle est un type ColorScale.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Renvoie les propriétés de mise en forme conditionnelle ColorScale si la mise en forme conditionnelle actuelle est un type ColorScale.|
||[custom](/javascript/api/excel/excel.conditionalformat#custom)|Renvoie les propriétés de mise en forme conditionnelle personnalisée si la mise en forme conditionnelle actuelle est un type personnalisé.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Renvoie les propriétés de mise en forme conditionnelle personnalisée si la mise en forme conditionnelle actuelle est un type personnalisé.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Renvoie les propriétés de la barre de données si la mise en forme conditionnelle actuelle est une barre de données.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Renvoie les propriétés de la barre de données si la mise en forme conditionnelle actuelle est une barre de données.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Renvoie les propriétés de mise en forme conditionnelle IconSet si la mise en forme conditionnelle actuelle est un type IconSet.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Renvoie les propriétés de mise en forme conditionnelle IconSet si la mise en forme conditionnelle actuelle est un type IconSet.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|La priorité de la mise en forme conditionnelle dans la ConditionalFormatCollection actuelle.|
||[preset](/javascript/api/excel/excel.conditionalformat#preset)|Renvoie la mise en forme conditionnelle des critères prédéfinits.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Renvoie la mise en forme conditionnelle des critères prédéfinits.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si la mise en forme conditionnelle actuelle est un type de texte.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si la mise en forme conditionnelle actuelle est un type de texte.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Renvoie les propriétés de mise en forme conditionnelle Top/Bottom si la mise en forme conditionnelle actuelle est de type TopBottom.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Renvoie les propriétés de mise en forme conditionnelle Top/Bottom si la mise en forme conditionnelle actuelle est de type TopBottom.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Type de mise en forme conditionnelle.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Si les conditions de cette mise en forme conditionnelle sont remplies, aucun format de priorité inférieure ne doit prendre effet sur cette cellule.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Ajoute un nouveau format conditionnel à la collection à la première/priorité supérieure.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Renvoie le nombre de formats conditionnels dans le manuel.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Renvoie une mise en forme conditionnelle à un ID donné.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Renvoie une mise en forme conditionnelle à l’index donné.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle dans la langue de l’utilisateur.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle dans la notation du style R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Icône personnalisée pour le critère en cours si différent de la celui par défaut IconSet. Sinon, null est renvoyé.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Un nombre ou une formule en fonction du type.|
||[opérateur](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan ou GreaterThanOrEqual pour chaque type de règle pour la mise en forme conditionnelle de l’icône.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|Ce sur quoi la formule conditionnelle de l’icône doit être basée.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[critère](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Critère de la mise en forme conditionnelle.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Valeur constante qui indique un côté spécifique de la bordure.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Obtient un objet de bordure à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Obtient un objet de bordure à l’aide de son indice.|
||[bas](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Obtient la bordure inférieure.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Nombre d’objets de bordure de la collection.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Obtient la bordure gauche.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Obtient la bordure droite.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Obtient la bordure supérieure.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Réinitialise le remplissage.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|Code couleur HTML représentant la couleur du remplissage, du formulaire #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Spécifie si la police est en gras.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Réinitialise les formats de police.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Représentation de code couleur HTML de la couleur du texte (par exemple, #FF0000 représente le rouge).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Spécifie si la police est en italique.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Spécifie l’état de strikethrough de la police.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Type de soulignement appliqué à la police.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Représente le Excel de format numérique de la plage donnée.|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Collection d’objets de bordure qui s’appliquent à la plage de mise en forme conditionnelle globale.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Renvoie l’objet de remplissage défini sur la plage de mise en forme conditionnelle globale.|
||[police](/javascript/api/excel/excel.conditionalrangeformat#font)|Renvoie l’objet de police défini sur la plage de mise en forme conditionnelle globale.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[opérateur](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Opérateur de la mise en forme conditionnelle du texte.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Valeur de texte de la mise en forme conditionnelle.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|Rang compris entre 1 et 1000 pour les rangs numériques ou entre 1 et 100 pour les rangs en pourcentage.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Formater les valeurs en fonction du classement supérieur ou inférieur.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.customconditionalformat#rule)|Spécifie l’objet Rule sur ce format conditionnel.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|Code couleur HTML représentant la couleur de la ligne Axe, au format #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Représentation de la façon dont l’axe est déterminé pour une barre Excel données.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Spécifie la direction sur qui le graphique de barre de données doit être basé.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Règle de ce qui constitue la limite inférieure (et comment la calculer, le cas échéant) pour une barre de données.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Représentation de toutes les valeurs à gauche de l’axe dans une barre Excel données.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Représentation de toutes les valeurs à droite de l’axe dans une barre Excel données.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Si la valeur est True, masque les valeurs des cellules où la barre de données est appliquée.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Règle de ce qui constitue la limite supérieure (et comment la calculer, le cas échéant) pour une barre de données.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Tableau de critères et iconsets pour les règles et icônes personnalisées potentielles pour les icônes conditionnelles.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Si la valeur est True, inverse les commandes d’icône pour IconSet.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Si la valeur est True, masque les valeurs et affiche uniquement les icônes.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Si elle est définie, affiche l’option IconSet pour la mise en forme conditionnelle.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Règle de mise en forme conditionnelle.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Calcule une plage de cellules dans une feuille de calcul.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Collection de conditionalFormats qui coupent la plage.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.textconditionalformat#rule)|Règle de mise en forme conditionnelle.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Critères du format conditionnel Haut/Bas.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Calcule toutes les cellules d’une feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
