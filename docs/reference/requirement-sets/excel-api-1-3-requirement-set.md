---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,3
description: Détails sur l’ensemble de conditions requises ExcelApi 1,3
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 684b802a32e58591d43d46a37ecc8b53395b652c
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940758"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Excel

ExcelApi 1,3 Ajout de la prise en charge de la liaison de données et de l’accès aux tableaux croisés dynamiques.

## <a name="api-list"></a>Liste des API

| Class | Champs | Description |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Supprime la liaison.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Add (Range: chaîne \| de plage, BindingType: Excel. bindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Ajoute une nouvelle liaison à une plage spécifique.|
||[addFromNamedItem (Name: String, bindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Ajoute une nouvelle liaison basée sur un élément nommé dans le classeur.|
||[addFromSelection (bindingType: Excel. BindingType, ID: chaîne)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Ajoute une nouvelle liaison basée sur la sélection en cours.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nom du tableau croisé dynamique.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|Feuille de calcul contenant le tableau croisé dynamique.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Actualise le tableau croisé dynamique.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Obtient un tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[RefreshAll, ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Actualise tous les tableaux croisés dynamiques de la collection.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView ()](/javascript/api/excel/excel.range#getvisibleview--)|Représente les lignes visibles de la plage en cours.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Représente la formule dans le style de notation R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Obtient la plage parent associée à l’affichage de plage actuel.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Représente le code de format de nombre d’Excel pour une cellule donnée.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Représente les adresses de cellule de la RangeView. En lecture seule.|
||[NbreColonnes](/javascript/api/excel/excel.rangeview#columncount)|Renvoie le nombre de colonnes visibles. En lecture seule.|
||[index](/javascript/api/excel/excel.rangeview#index)|Renvoie une valeur qui représente l’index de l’affichage de plage. En lecture seule.|
||[Stopp](/javascript/api/excel/excel.rangeview#rowcount)|Renvoie le nombre de lignes visibles. En lecture seule.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Représente une collection d’affichages de plage associés à la plage. En lecture seule.|
||[text](/javascript/api/excel/excel.rangeview#text)|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Représente le type de données de chaque cellule. En lecture seule.|
||[values](/javascript/api/excel/excel.rangeview#values)|Représente les valeurs brutes de l’affichage de plage spécifié. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Obtient une ligne RangeView par le biais de son index. Avec index de base zéro.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Indique si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Indique si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Indique si les colonnes affichent une mise en forme à bandes dans laquelle la mise en évidence des colonnes impaires diffère de celle des colonnes paires pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Indique si les lignes affichent une mise en forme à bandes dans laquelle la mise en évidence des lignes impaires diffère de celle des lignes paires pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Indique si les boutons de filtre sont visibles dans la partie supérieure de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivottables)|Représente une collection de tableaux croisés dynamiques associés au classeur. En lecture seule.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivottables)|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul. En lecture seule.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
