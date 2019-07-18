---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,3
description: Détails sur l’ensemble de conditions requises ExcelApi 1,3
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4698b0fad3122c8ecf52117c35d4928305d812fc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771994"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Excel

ExcelApi 1,3 Ajout de la prise en charge de la liaison de données et de l’accès aux tableaux croisés dynamiques.

## <a name="api-list"></a>Liste des API

| Class | Champs | Description |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Supprime la liaison.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Add (Range: plage \| String, bindingType: "Range" \| "table" \| "Text", ID: String)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Ajoute une nouvelle liaison à une plage spécifique.|
||[Add (Range: chaîne \| de plage, BindingType: Excel. bindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Ajoute une nouvelle liaison à une plage spécifique.|
||[addFromNamedItem (Name: String, bindingType: "Range" \| "table" \| "Text", ID: String)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Ajoute une nouvelle liaison basée sur un élément nommé dans le classeur.|
||[addFromNamedItem (Name: String, bindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Ajoute une nouvelle liaison basée sur un élément nommé dans le classeur.|
||[addFromSelection (bindingType: "Range" \| "table" \| "Text", ID: String)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Ajoute une nouvelle liaison basée sur la sélection en cours.|
||[addFromSelection (bindingType: Excel. BindingType, ID: chaîne)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Ajoute une nouvelle liaison basée sur la sélection en cours.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nom du tableau croisé dynamique.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|Feuille de calcul contenant le tableau croisé dynamique.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Actualise le tableau croisé dynamique.|
||[Set (propriétés: Excel. PivotTable)](/javascript/api/excel/excel.pivottable#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. PivotTableUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.pivottable#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Obtient un tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[RefreshAll, ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Actualise tous les tableaux croisés dynamiques de la collection.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablecollectionloadoptions#name)|Pour chaque élément de la collection: nom du tableau croisé dynamique.|
||[worksheet](/javascript/api/excel/excel.pivottablecollectionloadoptions#worksheet)|Pour chaque élément de la collection: feuille de calcul contenant le tableau croisé dynamique actif.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[name](/javascript/api/excel/excel.pivottabledata#name)|Nom du tableau croisé dynamique.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[$all](/javascript/api/excel/excel.pivottableloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottableloadoptions#name)|Nom du tableau croisé dynamique.|
||[worksheet](/javascript/api/excel/excel.pivottableloadoptions#worksheet)|Feuille de calcul contenant le tableau croisé dynamique.|
|[PivotTableUpdateData](/javascript/api/excel/excel.pivottableupdatedata)|[name](/javascript/api/excel/excel.pivottableupdatedata#name)|Nom du tableau croisé dynamique.|
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
||[Set (propriétés: Excel. RangeView)](/javascript/api/excel/excel.rangeview#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. RangeViewUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.rangeview#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[values](/javascript/api/excel/excel.rangeview#values)|Représente les valeurs brutes de l’affichage de plage spécifié. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Obtient une ligne RangeView par le biais de son index. Avec index de base zéro.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[RangeViewCollectionLoadOptions](/javascript/api/excel/excel.rangeviewcollectionloadoptions)|[$all](/javascript/api/excel/excel.rangeviewcollectionloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewcollectionloadoptions#celladdresses)|Pour chaque élément de la collection: représente les adresses des cellules du RangeView. En lecture seule.|
||[NbreColonnes](/javascript/api/excel/excel.rangeviewcollectionloadoptions#columncount)|Pour chaque élément de la collection: renvoie le nombre de colonnes visibles. En lecture seule.|
||[formulas](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulas)|Pour chaque élément de la collection: représente la formule en notation de style a1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulaslocal)|Pour chaque élément de la collection: représente la formule en notation de style a1, dans les paramètres régionaux de la langue de l’utilisateur et de la mise en forme des nombres.  Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulasr1c1)|Pour chaque élément de la collection: représente la formule en notation de style L1C1.|
||[index](/javascript/api/excel/excel.rangeviewcollectionloadoptions#index)|Pour chaque élément de la collection: renvoie une valeur qui représente l’index de la RangeView. En lecture seule.|
||[numberFormat](/javascript/api/excel/excel.rangeviewcollectionloadoptions#numberformat)|Pour chaque élément de la collection: représente le code de format de nombre d’Excel pour la cellule donnée.|
||[Stopp](/javascript/api/excel/excel.rangeviewcollectionloadoptions#rowcount)|Pour chaque élément de la collection: renvoie le nombre de lignes visibles. En lecture seule.|
||[text](/javascript/api/excel/excel.rangeviewcollectionloadoptions#text)|Pour chaque élément de la collection: valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|
||[valueTypes](/javascript/api/excel/excel.rangeviewcollectionloadoptions#valuetypes)|Pour chaque élément de la collection: représente le type de données de chaque cellule. En lecture seule.|
||[values](/javascript/api/excel/excel.rangeviewcollectionloadoptions#values)|Pour chaque élément de la collection: représente les valeurs brutes de l’affichage de plage spécifié. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[RangeViewData](/javascript/api/excel/excel.rangeviewdata)|[cellAddresses](/javascript/api/excel/excel.rangeviewdata#celladdresses)|Représente les adresses de cellule de la RangeView. En lecture seule.|
||[NbreColonnes](/javascript/api/excel/excel.rangeviewdata#columncount)|Renvoie le nombre de colonnes visibles. En lecture seule.|
||[formulas](/javascript/api/excel/excel.rangeviewdata#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewdata#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewdata#formulasr1c1)|Représente la formule dans le style de notation R1C1.|
||[index](/javascript/api/excel/excel.rangeviewdata#index)|Renvoie une valeur qui représente l’index de l’affichage de plage. En lecture seule.|
||[numberFormat](/javascript/api/excel/excel.rangeviewdata#numberformat)|Représente le code de format de nombre d’Excel pour une cellule donnée.|
||[rowCount](/javascript/api/excel/excel.rangeviewdata#rowcount)|Renvoie le nombre de lignes visibles. En lecture seule.|
||[rows](/javascript/api/excel/excel.rangeviewdata#rows)|Représente une collection d’affichages de plage associés à la plage. En lecture seule.|
||[text](/javascript/api/excel/excel.rangeviewdata#text)|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|
||[valueTypes](/javascript/api/excel/excel.rangeviewdata#valuetypes)|Représente le type de données de chaque cellule. En lecture seule.|
||[values](/javascript/api/excel/excel.rangeviewdata#values)|Représente les valeurs brutes de l’affichage de plage spécifié. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[RangeViewLoadOptions](/javascript/api/excel/excel.rangeviewloadoptions)|[$all](/javascript/api/excel/excel.rangeviewloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewloadoptions#celladdresses)|Représente les adresses de cellule de la RangeView. En lecture seule.|
||[NbreColonnes](/javascript/api/excel/excel.rangeviewloadoptions#columncount)|Renvoie le nombre de colonnes visibles. En lecture seule.|
||[formulas](/javascript/api/excel/excel.rangeviewloadoptions#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewloadoptions#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewloadoptions#formulasr1c1)|Représente la formule dans le style de notation R1C1.|
||[index](/javascript/api/excel/excel.rangeviewloadoptions#index)|Renvoie une valeur qui représente l’index de l’affichage de plage. En lecture seule.|
||[numberFormat](/javascript/api/excel/excel.rangeviewloadoptions#numberformat)|Représente le code de format de nombre d’Excel pour une cellule donnée.|
||[rowCount](/javascript/api/excel/excel.rangeviewloadoptions#rowcount)|Renvoie le nombre de lignes visibles. En lecture seule.|
||[text](/javascript/api/excel/excel.rangeviewloadoptions#text)|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|
||[valueTypes](/javascript/api/excel/excel.rangeviewloadoptions#valuetypes)|Représente le type de données de chaque cellule. En lecture seule.|
||[values](/javascript/api/excel/excel.rangeviewloadoptions#values)|Représente les valeurs brutes de l’affichage de plage spécifié. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[RangeViewUpdateData](/javascript/api/excel/excel.rangeviewupdatedata)|[formulas](/javascript/api/excel/excel.rangeviewupdatedata#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewupdatedata#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewupdatedata#formulasr1c1)|Représente la formule dans le style de notation R1C1.|
||[numberFormat](/javascript/api/excel/excel.rangeviewupdatedata#numberformat)|Représente le code de format de nombre d’Excel pour une cellule donnée.|
||[values](/javascript/api/excel/excel.rangeviewupdatedata#values)|Représente les valeurs brutes de l’affichage de plage spécifié. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Indique si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Indique si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Indique si les colonnes affichent une mise en forme à bandes dans laquelle la mise en évidence des colonnes impaires diffère de celle des colonnes paires pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Indique si les lignes affichent une mise en forme à bandes dans laquelle la mise en évidence des lignes impaires diffère de celle des lignes paires pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Indique si les boutons de filtre sont visibles dans la partie supérieure de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightfirstcolumn)|Pour chaque élément de la collection: indique si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightlastcolumn)|Pour chaque élément de la collection: indique si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedcolumns)|Pour chaque élément de la collection: indique si les colonnes affichent une mise en forme à bandes dans laquelle les colonnes impaires sont mises en surbrillance différemment des colonnes égales pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedrows)|Pour chaque élément de la collection: indique si les lignes affichent une mise en forme à bandes dans laquelle les lignes impaires sont mises en surbrillance différemment des lignes paires pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.tablecollectionloadoptions#showfilterbutton)|Pour chaque élément de la collection: indique si les boutons de filtre sont visibles en haut de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|
|[TableData](/javascript/api/excel/excel.tabledata)|[highlightFirstColumn](/javascript/api/excel/excel.tabledata#highlightfirstcolumn)|Indique si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.tabledata#highlightlastcolumn)|Indique si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.tabledata#showbandedcolumns)|Indique si les colonnes affichent une mise en forme à bandes dans laquelle la mise en évidence des colonnes impaires diffère de celle des colonnes paires pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.tabledata#showbandedrows)|Indique si les lignes affichent une mise en forme à bandes dans laquelle la mise en évidence des lignes impaires diffère de celle des lignes paires pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.tabledata#showfilterbutton)|Indique si les boutons de filtre sont visibles dans la partie supérieure de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tableloadoptions#highlightfirstcolumn)|Indique si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.tableloadoptions#highlightlastcolumn)|Indique si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.tableloadoptions#showbandedcolumns)|Indique si les colonnes affichent une mise en forme à bandes dans laquelle la mise en évidence des colonnes impaires diffère de celle des colonnes paires pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.tableloadoptions#showbandedrows)|Indique si les lignes affichent une mise en forme à bandes dans laquelle la mise en évidence des lignes impaires diffère de celle des lignes paires pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.tableloadoptions#showfilterbutton)|Indique si les boutons de filtre sont visibles dans la partie supérieure de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|
|[TableUpdateData](/javascript/api/excel/excel.tableupdatedata)|[highlightFirstColumn](/javascript/api/excel/excel.tableupdatedata#highlightfirstcolumn)|Indique si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.tableupdatedata#highlightlastcolumn)|Indique si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.tableupdatedata#showbandedcolumns)|Indique si les colonnes affichent une mise en forme à bandes dans laquelle la mise en évidence des colonnes impaires diffère de celle des colonnes paires pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.tableupdatedata#showbandedrows)|Indique si les lignes affichent une mise en forme à bandes dans laquelle la mise en évidence des lignes impaires diffère de celle des lignes paires pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.tableupdatedata#showfilterbutton)|Indique si les boutons de filtre sont visibles dans la partie supérieure de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivottables)|Représente une collection de tableaux croisés dynamiques associés au classeur. En lecture seule.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[pivotTables](/javascript/api/excel/excel.workbookdata#pivottables)|Représente une collection de tableaux croisés dynamiques associés au classeur. En lecture seule.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivottables)|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul. En lecture seule.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[pivotTables](/javascript/api/excel/excel.worksheetdata#pivottables)|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul. En lecture seule.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
