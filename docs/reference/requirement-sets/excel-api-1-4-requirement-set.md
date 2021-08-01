---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.4
description: Détails sur l’ensemble de conditions requises ExcelApi 1.4.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: be71d1e0c063bd3902bf57ba8f2024ae5a78ff1d
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671722"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Nouveautés de l’API JavaScript 1.4 pour Excel

Les ajouts apportés aux API JavaScript pour Excel dans l’ensemble de conditions requises 1.4 sont présentés ci-dessous.

## <a name="named-item-add-and-new-properties"></a>Ajout d’élément nommé et nouvelles propriétés

Nouvelles propriétés :

* `comment`
* `scope` - Éléments d’étendue de feuille de calcul ou de feuille de calcul.
* `worksheet` - Renvoie la feuille de calcul dans laquelle l’élément nommé est étendue.

Nouvelles méthodes :

* `add(name: string, reference: Range or string, comment: string)` - Ajoute un nouveau nom à la collection de l’étendue donnée.
* `addFormulaLocal(name: string, formula: string, comment: string)` - Ajoute un nouveau nom à la collection de l’étendue donnée à l’aide des paramètres régionaux de l’utilisateur pour la formule.

## <a name="settings-api-in-the-excel-namespace"></a>API Settings dans l’espace de noms Excel

L’objet [Setting](/javascript/api/excel/excel.setting) représente une paire clé-valeur d’un paramètre conservé dans le document. La fonctionnalité de `Excel.Setting` équivaut à `Office.Settings`, mais utilise la syntaxe d’API par lots plutôt que le modèle de rappel de l’API commune.

Les API incluent l’accès à l’entrée de paramètre via la clé et l’ajout de la paire de `getItem()` `add()` paramètres key:value spécifiée au workbook.

## <a name="others"></a>Autres

* Définissez le nom de colonne de la table.
* Ajoutez une colonne de tableau à la fin du tableau.
* Ajoutez plusieurs lignes à un tableau à la fois.
* `range.getColumnsAfter(count: number)` et `range.getColumnsBefore(count: number)` pour obtenir un certain nombre de colonnes à droite/gauche de l’objet de plage actuel.
* Méthodes [ \* et propriétés OrNullObject](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties): cette fonctionnalité permet d’obtenir un objet à l’aide d’une clé. Si l’objet n’existe pas, la propriété de l’objet `isNullObject` renvoyé est true. Cela permet aux développeurs de vérifier si un objet existe sans avoir à le gérer par le biais de la gestion des exceptions. Une `*OrNullObject` méthode est disponible sur la plupart des objets de collection.

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.4. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.4 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.4](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getCount__)|Obtient le nombre de liaisons de la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getItemOrNullObject_id_)|Obtient un objet de liaison par ID.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getCount__)|Renvoie le nombre de graphiques dans la feuille de calcul.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getItemOrNullObject_name_)|Extrait un graphique à l’aide de son nom.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getCount__)|Renvoie le nombre de points de graphique dans la série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getCount__)|Renvoie le nombre de séries de la collection.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Spécifie le commentaire associé à ce nom.|
||[delete()](/javascript/api/excel/excel.nameditem#delete__)|Supprime le nom donné.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getRangeOrNullObject__)|Renvoie l’objet de plage qui est associé au nom.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Spécifie si le nom est d’étendue au workbook ou à une feuille de calcul spécifique.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Renvoie la feuille de calcul dans laquelle est inclus l’élément nommé.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetOrNullObject)|Renvoie la feuille de calcul dans laquelle l’élément nommé est d’étendue.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#add_name__reference__comment_)|Ajoute un nouveau nom à la collection de l’étendue donnée.|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#addFormulaLocal_name__formula__comment_)|Ajoute un nouveau nom à la collection de l’étendue donnée à l’aide des paramètres régionaux de l’utilisateur pour la formule.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getCount__)|Obtient le nombre d’éléments nommés dans la collection.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getItemOrNullObject_name_)|Obtient un `NamedItem` objet à l’aide de son nom.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getCount__)|Obtient le nombre de tableaux croisés dynamiques de la collection.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getItemOrNullObject_name_)|Obtient un tableau croisé dynamique par nom.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getIntersectionOrNullObject_anotherRange_)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#getUsedRangeOrNullObject_valuesOnly_)|Renvoie la plage utilisée d’un objet de plage donné.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getCount__)|Obtient le nombre `RangeView` d’objets de la collection.|
|[Paramètre](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete__)|Supprime le paramètre.|
||[key](/javascript/api/excel/excel.setting#key)|Clé qui représente l’ID du paramètre.|
||[value](/javascript/api/excel/excel.setting#value)|Représente la valeur stockée pour ce paramètre.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date Array \| <any> \| any)](/javascript/api/excel/excel.settingcollection#add_key__value_)|Définit ou ajoute le paramètre spécifié dans le classeur.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getCount__)|Obtient le nombre de paramètres de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getItem_key_)|Obtient une entrée de paramètre via la clé.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getItemOrNullObject_key_)|Obtient une entrée de paramètre via la clé.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onSettingsChanged)|Se produit lorsque les paramètres du document sont modifiés.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[paramètres](/javascript/api/excel/excel.settingschangedeventargs#settings)|Obtient `Setting` l’objet qui représente la liaison qui a élevé l’événement de changement de paramètres|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getCount__)|Obtient le nombre de tableaux de la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getItemOrNullObject_key_)|Obtient un tableau à l’aide de son nom ou de son ID.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getCount__)|Obtient le nombre de colonnes dans le tableau.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getItemOrNullObject_key_)|Obtient un objet de colonne par son nom ou son ID.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getCount__)|Obtient le nombre de lignes dans le tableau.|
|[Classeur](/javascript/api/excel/excel.workbook)|[paramètres](/javascript/api/excel/excel.workbook#settings)|Représente une collection de paramètres associés au workbook.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#getUsedRangeOrNullObject_valuesOnly_)|La plage utilisée est la plus petite plage qui englobe toutes les cellules auxquelles une valeur ou un format est affecté.|
||[names](/javascript/api/excel/excel.worksheet#names)|Collection de noms inclus dans l’étendue de la feuille de calcul active.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getCount_visibleOnly_)|Obtient le nombre de feuilles de calcul dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getItemOrNullObject_key_)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
