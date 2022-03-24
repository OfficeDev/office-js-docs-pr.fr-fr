---
title: Excel l’ensemble de conditions requises de l’API JavaScript 1.3
description: Détails sur l’ensemble de conditions requises ExcelApi 1.3.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1bf8bc604c2c770f517878193994c1ed32640da1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745339"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Excel

ExcelApi 1.3 a ajouté la prise en charge de la liaison de données et de l’accès de base au tableau croisé dynamique.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.3. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.3 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.3](/javascript/api/excel?view=excel-js-1.3&preserve-view=true) ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#excel-excel-binding-delete-member(1))|Supprime la liaison.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-add-member(1))|Ajoute une nouvelle liaison à une plage spécifique.|
||[addFromNamedItem(name: string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromnameditem-member(1))|Ajoute une nouvelle liaison basée sur un élément nommé dans le classeur.|
||[addFromSelection(bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromselection-member(1))|Ajoute une nouvelle liaison basée sur la sélection en cours.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-name-member)|Nom du tableau croisé dynamique.|
||[refresh()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refresh-member(1))|Actualise le tableau croisé dynamique.|
||[worksheet](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-worksheet-member)|Feuille de calcul contenant le tableau croisé dynamique.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitem-member(1))|Obtient un tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-refreshall-member(1))|Actualise tous les tableaux croisés dynamiques de la collection.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#excel-excel-range-getvisibleview-member(1))|Représente les lignes visibles de la plage en cours.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[cellAddresses](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-celladdresses-member)|Représente les adresses de cellule du `RangeView`.|
||[columnCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-columncount-member)|Nombre de colonnes visibles.|
||[formulas](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulas-member)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulaslocal-member)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulasr1c1-member)|Représente la formule dans le style de notation R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-getrange-member(1))|Obtient la plage parent associée à l’actuel `RangeView`.|
||[index](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-index-member)|Renvoie une valeur qui représente l’index du `RangeView`.|
||[numberFormat](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-numberformat-member)|Représente le code de format de nombre d’Excel pour une cellule donnée.|
||[rowCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rowcount-member)|Nombre de lignes visibles.|
||[rows](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rows-member)|Représente une collection d’affichages de plage associés à la plage.|
||[text](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-text-member)|Valeurs de texte de la plage spécifiée.|
||[valueTypes](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuetypes-member)|Représente le type de données de chaque cellule.|
||[values](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-values-member)|Représente les valeurs brutes de l’affichage de plage spécifié.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getitemat-member(1))|Obtient une `RangeView` ligne via son index.|
||[items](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightfirstcolumn-member)|Spécifie si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightlastcolumn-member)|Spécifie si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.table#excel-excel-table-showbandedcolumns-member)|Spécifie si les colonnes indiquent une mise en forme à bandes dans laquelle les colonnes impaires sont mises en surbrillante différemment des colonnes impaires, pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.table#excel-excel-table-showbandedrows-member)|Spécifie si les lignes indiquent une mise en forme à bandes dans laquelle les lignes impaires sont mises en surbrillante différemment des lignes impaires, pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.table#excel-excel-table-showfilterbutton-member)|Spécifie si les boutons de filtre sont visibles en haut de chaque en-tête de colonne.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottables-member)|Représente une collection de tableaux croisés dynamiques associés au classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pivottables-member)|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
