---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.3
description: Détails sur l’ensemble de conditions requises ExcelApi 1.3.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 45a0a3551662997984a5c999b62c651d81e243f2
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149232"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Excel

ExcelApi 1.3 a ajouté la prise en charge de la liaison de données et de l’accès de base au tableau croisé dynamique.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.3. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.3 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.3](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete__)|Supprime la liaison.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#add_range__bindingType__id_)|Ajoute une nouvelle liaison à une plage spécifique.|
||[addFromNamedItem(name: string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addFromNamedItem_name__bindingType__id_)|Ajoute une nouvelle liaison basée sur un élément nommé dans le classeur.|
||[addFromSelection(bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addFromSelection_bindingType__id_)|Ajoute une nouvelle liaison basée sur la sélection en cours.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nom du tableau croisé dynamique.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|Feuille de calcul contenant le tableau croisé dynamique.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh__)|Actualise le tableau croisé dynamique.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getItem_name_)|Obtient un tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#refreshAll__)|Actualise tous les tableaux croisés dynamiques de la collection.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#getVisibleView__)|Représente les lignes visibles de la plage en cours.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulasLocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasR1C1)|Représente la formule dans le style de notation R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getRange__)|Obtient la plage parent associée à l’actuel `RangeView` .|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberFormat)|Représente le code de format de nombre d’Excel pour une cellule donnée.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#cellAddresses)|Représente les adresses de cellule du `RangeView` .|
||[columnCount](/javascript/api/excel/excel.rangeview#columnCount)|Nombre de colonnes visibles.|
||[index](/javascript/api/excel/excel.rangeview#index)|Renvoie une valeur qui représente l’index du `RangeView` .|
||[rowCount](/javascript/api/excel/excel.rangeview#rowCount)|Nombre de lignes visibles.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Représente une collection d’affichages de plage associés à la plage.|
||[text](/javascript/api/excel/excel.rangeview#text)|Valeurs de texte de la plage spécifiée.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valueTypes)|Représente le type de données de chaque cellule.|
||[values](/javascript/api/excel/excel.rangeview#values)|Représente les valeurs brutes de l’affichage de plage spécifié.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getItemAt_index_)|Obtient une `RangeView` ligne via son index.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightFirstColumn)|Spécifie si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightLastColumn)|Spécifie si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.table#showBandedColumns)|Spécifie si les colonnes indiquent une mise en forme à bandes dans laquelle les colonnes impaires sont mises en surbrillante différemment des colonnes impaires, pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.table#showBandedRows)|Spécifie si les lignes indiquent une mise en forme à bandes dans laquelle les lignes impaires sont mises en surbrillante différemment des lignes impaires, pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.table#showFilterButton)|Spécifie si les boutons de filtre sont visibles en haut de chaque en-tête de colonne.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivotTables)|Représente une collection de tableaux croisés dynamiques associés au classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivotTables)|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
