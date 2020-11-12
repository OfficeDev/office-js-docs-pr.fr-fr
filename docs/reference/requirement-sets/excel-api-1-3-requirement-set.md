---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,3
description: Détails sur l’ensemble de conditions requises ExcelApi 1,3.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 520755fe4b77008da866098d851f47ae3833bf13
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996472"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Excel

ExcelApi 1,3 Ajout de la prise en charge de la liaison de données et de l’accès aux tableaux croisés dynamiques.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Excel 1,3. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1,3 ou antérieure, voir [API Excel dans l’ensemble de conditions requises 1,3 ou version antérieure](/javascript/api/excel?view=excel-js-1.3&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Supprime la liaison.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Add (Range : \| chaîne de plage, bindingType : Excel. bindingType, ID : String)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Ajoute une nouvelle liaison à une plage spécifique.|
||[addFromNamedItem (Name : String, bindingType : Excel. BindingType, ID : String)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Ajoute une nouvelle liaison basée sur un élément nommé dans le classeur.|
||[addFromSelection (bindingType : Excel. BindingType, ID : chaîne)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Ajoute une nouvelle liaison basée sur la sélection en cours.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nom du tableau croisé dynamique.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|Feuille de calcul contenant le tableau croisé dynamique.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Actualise le tableau croisé dynamique.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Obtient un tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[RefreshAll, ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Actualise tous les tableaux croisés dynamiques de la collection.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView ()](/javascript/api/excel/excel.range#getvisibleview--)|Représente les lignes visibles de la plage en cours.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Représente la formule dans le style de notation R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Obtient la plage parent associée à l’affichage de plage actuel.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Représente le code de format de nombre d’Excel pour une cellule donnée.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Représente les adresses de cellule de la RangeView.|
||[NbreColonnes](/javascript/api/excel/excel.rangeview#columncount)|Nombre de colonnes visibles.|
||[index](/javascript/api/excel/excel.rangeview#index)|Renvoie une valeur qui représente l’index de l’affichage de plage.|
||[Stopp](/javascript/api/excel/excel.rangeview#rowcount)|Nombre de lignes visibles.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Représente une collection d’affichages de plage associés à la plage.|
||[text](/javascript/api/excel/excel.rangeview#text)|Valeurs de texte de la plage spécifiée.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Représente le type de données de chaque cellule.|
||[values](/javascript/api/excel/excel.rangeview#values)|Représente les valeurs brutes de l’affichage de plage spécifié.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Obtient une ligne RangeView par le biais de son index.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Indique si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Indique si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Spécifie si les colonnes affichent une mise en forme à bandes dans laquelle les colonnes impaires sont mises en surbrillance différemment des colonnes égales pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Indique si les lignes affichent une mise en forme à bandes dans laquelle les lignes impaires sont mises en surbrillance différemment des lignes paires pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Indique si les boutons de filtre sont visibles en haut de chaque en-tête de colonne.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivottables)|Représente une collection de tableaux croisés dynamiques associés au classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivottables)|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
