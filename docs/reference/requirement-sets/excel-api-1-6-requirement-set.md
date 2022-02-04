---
title: Excel conditions requises de l’API JavaScript 1.6
description: Détails sur l’ensemble de conditions requises ExcelApi 1.6.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
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

Le tableau suivant répertorie les API Excel l’ensemble de conditions requises de l’API JavaScript 1.6. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.6 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.6](/javascript/api/excel?view=excel-js-1.6&preserve-view=true) ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendapicalculationuntilnextsync-member(1))|Suspend le calcul jusqu’à ce que le suivant `context.sync()` soit appelé.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#excel-excel-cellvalueconditionalformat-format-member)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#excel-excel-cellvalueconditionalformat-rule-member)|Spécifie l’objet de règle sur ce format conditionnel.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#excel-excel-colorscaleconditionalformat-criteria-member)|Critères de l’échelle de couleurs.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#excel-excel-colorscaleconditionalformat-threecolorscale-member)|Si `true`, l’échelle de couleurs aura trois points (minimum, milieu, maximum), sinon elle en aura deux (minimum, maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-formula1-member)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-formula2-member)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle.|
||[opérateur](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-operator-member)|Opérateur de la mise en forme conditionnelle de la valeur de cellule.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-maximum-member)|Point maximal du critère d’échelle de couleur.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-midpoint-member)|Milieu du critère d’échelle de couleur, si l’échelle de couleurs est une échelle de 3 couleurs.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-minimum-member)|Point minimal du critère d’échelle de couleur.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-color-member)|Représentation de code couleur HTML de la couleur d’échelle de couleur (par exemple, #FF0000 représente le rouge).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-formula-member)|Un nombre, une formule ou `null` (si c’est `type` le cas `lowestValue`).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-type-member)|Sur quoi la formule conditionnelle critère doit être basée.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-bordercolor-member)|Code couleur HTML représentant la couleur de la bordure, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-fillcolor-member)|Code couleur HTML représentant la couleur de remplissage, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-matchpositivebordercolor-member)|Spécifie si la barre de données négative a la même couleur de bordure que la barre de données positive.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-matchpositivefillcolor-member)|Spécifie si la barre de données négative a la même couleur de remplissage que la barre de données positive.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-bordercolor-member)|Code couleur HTML représentant la couleur de la bordure, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-fillcolor-member)|Code couleur HTML représentant la couleur de remplissage, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-gradientfill-member)|Spécifie si la barre de données a un dégradé.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#excel-excel-conditionaldatabarrule-formula-member)|Formule, si nécessaire, sur laquelle évaluer la règle de barre de données.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#excel-excel-conditionaldatabarrule-type-member)|Type de règle pour la barre de données.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[cellValue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalue-member)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de cellule si la mise en forme conditionnelle actuelle est un `CellValue` type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalueornullobject-member)|Renvoie les propriétés de mise en forme conditionnelle de la valeur de cellule si la mise en forme conditionnelle actuelle est un `CellValue` type.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscale-member)|Renvoie les propriétés de mise en forme conditionnelle d’échelle de couleur si la mise en forme conditionnelle actuelle est un `ColorScale` type.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscaleornullobject-member)|Renvoie les propriétés de mise en forme conditionnelle d’échelle de couleur si la mise en forme conditionnelle actuelle est un `ColorScale` type.|
||[custom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-custom-member)|Renvoie les propriétés de mise en forme conditionnelle personnalisée si la mise en forme conditionnelle actuelle est un type personnalisé.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-customornullobject-member)|Renvoie les propriétés de mise en forme conditionnelle personnalisée si la mise en forme conditionnelle actuelle est un type personnalisé.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databar-member)|Renvoie les propriétés de la barre de données si la mise en forme conditionnelle actuelle est une barre de données.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databarornullobject-member)|Renvoie les propriétés de la barre de données si la mise en forme conditionnelle actuelle est une barre de données.|
||[delete()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-delete-member(1))|Supprime cette mise en forme conditionnelle.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrange-member(1))|Renvoie la plage à laquelle s’applique la mise en forme conditionnelle.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrangeornullobject-member(1))|Renvoie la plage à laquelle le format conditionnel est appliqué.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconset-member)|Renvoie les propriétés de mise en forme conditionnelle du jeu d’icônes si la mise en forme conditionnelle actuelle est un `IconSet` type.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconsetornullobject-member)|Renvoie les propriétés de mise en forme conditionnelle du jeu d’icônes si la mise en forme conditionnelle actuelle est un `IconSet` type.|
||[id](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-id-member)|Priorité de la mise en forme conditionnelle dans la version actuelle `ConditionalFormatCollection`.|
||[preset](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-preset-member)|Renvoie la mise en forme conditionnelle des critères prédéfinits.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-presetornullobject-member)|Renvoie la mise en forme conditionnelle des critères prédéfinits.|
||[priority](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-priority-member)|Priorité (ou index) dans la collection de formats conditionnels dans qui ce format conditionnel existe actuellement.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-stopiftrue-member)|Si les conditions de cette mise en forme conditionnelle sont remplies, aucun format de priorité inférieure ne doit prendre effet sur cette cellule.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparison-member)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si la mise en forme conditionnelle actuelle est un type de texte.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparisonornullobject-member)|Renvoie les propriétés de mise en forme conditionnelle de texte spécifiques si la mise en forme conditionnelle actuelle est un type de texte.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottom-member)|Renvoie les propriétés de mise en forme conditionnelle supérieure/inférieure si la mise en forme conditionnelle actuelle est un `TopBottom` type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottomornullobject-member)|Renvoie les propriétés de mise en forme conditionnelle supérieure/inférieure si la mise en forme conditionnelle actuelle est un `TopBottom` type.|
||[type](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-type-member)|Type de mise en forme conditionnelle.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-add-member(1))|Ajoute un nouveau format conditionnel à la collection à la première/priorité supérieure.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-clearall-member(1))|Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getcount-member(1))|Renvoie le nombre de formats conditionnels dans le manuel.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitem-member(1))|Renvoie une mise en forme conditionnelle à un ID donné.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitemat-member(1))|Renvoie une mise en forme conditionnelle à l’index donné.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formula-member)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formulalocal-member)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle dans la langue de l’utilisateur.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formular1c1-member)|Formule, si nécessaire, sur laquelle évaluer la règle de mise en forme conditionnelle dans la notation de style R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-customicon-member)|L’icône personnalisée pour le critère actuel, si différente du jeu d’icônes par défaut, est `null` renvoyée.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-formula-member)|Un nombre ou une formule en fonction du type.|
||[opérateur](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-operator-member)|`greaterThan` ou `greaterThanOrEqual` pour chacun des types de règles pour la mise en forme conditionnelle d’icône.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-type-member)|Ce sur quoi la formule conditionnelle de l’icône doit être basée.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[critère](/javascript/api/excel/excel.conditionalpresetcriteriarule#excel-excel-conditionalpresetcriteriarule-criterion-member)|Critère de la mise en forme conditionnelle.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-color-member)|Code couleur HTML représentant la couleur de la bordure, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-sideindex-member)|Valeur constante qui indique un côté spécifique de la bordure.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-style-member)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[bas](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-bottom-member)|Obtient la bordure inférieure.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-count-member)|Nombre d’objets de bordure de la collection.|
||[getItem(index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-getitem-member(1))|Obtient un objet de bordure à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-getitemat-member(1))|Obtient un objet de bordure à l’aide de son indice.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-left-member)|Obtient la bordure gauche.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-right-member)|Obtient la bordure droite.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-top-member)|Obtient la bordure supérieure.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#excel-excel-conditionalrangefill-clear-member(1))|Réinitialise le remplissage.|
||[color](/javascript/api/excel/excel.conditionalrangefill#excel-excel-conditionalrangefill-color-member)|Code couleur HTML représentant la couleur du remplissage, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-bold-member)|Spécifie si la police est en gras.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-clear-member(1))|Réinitialise les formats de police.|
||[color](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-color-member)|Représentation de code couleur HTML de la couleur du texte (par exemple, #FF0000 représente le rouge).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-italic-member)|Spécifie si la police est en italique.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-strikethrough-member)|Spécifie l’état de strikethrough de la police.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-underline-member)|Type de soulignement appliqué à la police.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[Borders](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-borders-member)|Collection d’objets de bordure qui s’appliquent à la plage de mise en forme conditionnelle globale.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-fill-member)|Renvoie l’objet de remplissage défini sur la plage de mise en forme conditionnelle globale.|
||[police](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-font-member)|Renvoie l’objet de police défini sur la plage de mise en forme conditionnelle globale.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-numberformat-member)|Représente le Excel de format numérique de la plage donnée.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[opérateur](/javascript/api/excel/excel.conditionaltextcomparisonrule#excel-excel-conditionaltextcomparisonrule-operator-member)|Opérateur de la mise en forme conditionnelle du texte.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#excel-excel-conditionaltextcomparisonrule-text-member)|Valeur de texte de la mise en forme conditionnelle.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#excel-excel-conditionaltopbottomrule-rank-member)|Rang compris entre 1 et 1000 pour les rangs numériques ou entre 1 et 100 pour les rangs en pourcentage.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#excel-excel-conditionaltopbottomrule-type-member)|Formater les valeurs en fonction du classement supérieur ou inférieur.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#excel-excel-customconditionalformat-format-member)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.customconditionalformat#excel-excel-customconditionalformat-rule-member)|Spécifie l’objet `Rule` sur cette mise en forme conditionnelle.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-axiscolor-member)|Code couleur HTML représentant la couleur de la ligne Axe, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-axisformat-member)|Représentation de la façon dont l’axe est déterminé pour une barre Excel données.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-bardirection-member)|Spécifie la direction sur qui le graphique de barre de données doit être basé.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-lowerboundrule-member)|Règle de ce qui constitue la limite inférieure (et comment la calculer, le cas échéant) pour une barre de données.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-negativeformat-member)|Représentation de toutes les valeurs à gauche de l’axe dans une barre Excel données.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-positiveformat-member)|Représentation de toutes les valeurs à droite de l’axe dans une barre Excel données.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-showdatabaronly-member)|Si `true`, masque les valeurs des cellules où la barre de données est appliquée.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-upperboundrule-member)|Règle de ce qui constitue la limite supérieure (et comment la calculer, le cas échéant) pour une barre de données.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-criteria-member)|Tableau de critères et d’ensembles d’icônes pour les règles et icônes personnalisées potentielles pour les icônes conditionnelles.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-reverseiconorder-member)|Si `true`, inverse les commandes d’icône pour le jeu d’icônes.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-showicononly-member)|Si `true`, masque les valeurs et affiche uniquement les icônes.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-style-member)|Si elle est définie, affiche l’option de jeu d’icônes pour la mise en forme conditionnelle.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#excel-excel-presetcriteriaconditionalformat-format-member)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés des formats conditionnels.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#excel-excel-presetcriteriaconditionalformat-rule-member)|Règle de mise en forme conditionnelle.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#excel-excel-range-calculate-member(1))|Calcule une plage de cellules dans une feuille de calcul.|
||[conditionalFormats](/javascript/api/excel/excel.range#excel-excel-range-conditionalformats-member)|Collection de cette plage `ConditionalFormats` .|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#excel-excel-textconditionalformat-format-member)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés de la mise en forme conditionnelle.|
||[rule](/javascript/api/excel/excel.textconditionalformat#excel-excel-textconditionalformat-rule-member)|Règle de mise en forme conditionnelle.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#excel-excel-topbottomconditionalformat-format-member)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures et d’autres propriétés de la mise en forme conditionnelle.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#excel-excel-topbottomconditionalformat-rule-member)|Critères de la mise en forme conditionnelle supérieure/inférieure.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-calculate-member(1))|Calcule toutes les cellules d’une feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
