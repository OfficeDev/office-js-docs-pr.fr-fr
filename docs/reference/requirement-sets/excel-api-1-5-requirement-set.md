---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,5
description: Détails sur l’ensemble de conditions requises ExcelApi 1,5
ms.date: 07/15/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b8f767a83b7e373b422b6fc0d9ac65de90c04f5
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771966"
---
#  <a name="whats-new-in-excel-javascript-api-15"></a>Nouveautés de l’API JavaScript 1.5 pour Excel

ExcelApi 1,5 ajoute des parties XML personnalisées. Celles-ci sont accessibles via la [collection de parties XML personnalisée](/javascript/api/excel/excel.workbook#customxmlparts) dans l’objet Workbook.

## <a name="custom-xml-part"></a>Partie XML personnalisée

* Obtenir des parties XML personnalisées à l’aide de leur ID.
* Obtenez une nouvelle collection délimitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.
* Obtient une chaîne XML associée à un composant.
* Fournissez l’ID et l’espace de noms d’un composant.
* Ajoutez une nouvelle partie XML personnalisée au classeur.
* Définir une partie XML entière.
* Supprimez une partie XML personnalisée.
* Supprimez un attribut avec le nom donné dans l’élément identifié par langage XPath.
* Interrogez le contenu XML par langage XPath.
* Attributs d’insertion, de mise à jour et de suppression.

## <a name="api-list"></a>Liste des API

| Class | Champs | Description |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Supprime la partie XML personnalisée.|
||[getXml ()](/javascript/api/excel/excel.customxmlpart#getxml--)|Obtient l’intégralité du contenu XML de la partie XML personnalisée.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|ID de la partie XML personnalisée. En lecture seule.|
||[URI](/javascript/api/excel/excel.customxmlpart#namespaceuri)|URI de l’espace de noms de la partie XML personnalisée. En lecture seule.|
||[setXml (XML: chaîne)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Définit l’intégralité du contenu XML de la partie XML personnalisée.|
|[Uncustomxmlpartcollection](/javascript/api/excel/excel.customxmlpartcollection)|[Add (XML: String)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Ajoute une nouvelle partie XML personnalisée au classeur.|
||[getByNamespace (namespaceUri: String)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtient une nouvelle collection limitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Obtient le nombre de parties CustomXml dans la collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CustomXmlPartCollectionLoadOptions](/javascript/api/excel/excel.customxmlpartcollectionloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#id)|Pour chaque élément de la collection: ID de la partie XML personnalisée. En lecture seule.|
||[URI](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#namespaceuri)|Pour chaque élément de la collection: URI de l’espace de noms de la partie XML personnalisée. En lecture seule.|
|[CustomXmlPartData](/javascript/api/excel/excel.customxmlpartdata)|[id](/javascript/api/excel/excel.customxmlpartdata#id)|ID de la partie XML personnalisée. En lecture seule.|
||[URI](/javascript/api/excel/excel.customxmlpartdata#namespaceuri)|URI de l’espace de noms de la partie XML personnalisée. En lecture seule.|
|[CustomXmlPartLoadOptions](/javascript/api/excel/excel.customxmlpartloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartloadoptions#id)|ID de la partie XML personnalisée. En lecture seule.|
||[URI](/javascript/api/excel/excel.customxmlpartloadoptions#namespaceuri)|URI de l’espace de noms de la partie XML personnalisée. En lecture seule.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Obtient le nombre de parties CustomXML dans cette collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CustomXmlPartScopedCollectionLoadOptions](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#id)|Pour chaque élément de la collection: ID de la partie XML personnalisée. En lecture seule.|
||[URI](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#namespaceuri)|Pour chaque élément de la collection: URI de l’espace de noms de la partie XML personnalisée. En lecture seule.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|ID du tableau croisé dynamique. En lecture seule.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[id](/javascript/api/excel/excel.pivottablecollectionloadoptions#id)|Pour chaque élément de la collection: ID du tableau croisé dynamique. En lecture seule.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[id](/javascript/api/excel/excel.pivottabledata#id)|ID du tableau croisé dynamique. En lecture seule.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[id](/javascript/api/excel/excel.pivottableloadoptions#id)|ID du tableau croisé dynamique. En lecture seule.|
|[Runtime](/javascript/api/excel/excel.runtime)|[Set (propriétés: Excel. Runtime)](/javascript/api/excel/excel.runtime#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. RuntimeUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.runtime#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[RuntimeLoadOptions](/javascript/api/excel/excel.runtimeloadoptions)|[$all](/javascript/api/excel/excel.runtimeloadoptions#$all)||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Représente la collection de parties XML personnalisées contenues dans ce classeur. En lecture seule.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[customXmlParts](/javascript/api/excel/excel.workbookdata#customxmlparts)|Représente la collection de parties XML personnalisées contenues dans ce classeur. En lecture seule.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Obtient la feuille de calcul qui suit celle-ci. S’il n’existe aucune feuille de calcul à la suite de celle-ci, cette méthode génère une erreur.|
||[getNextOrNullObject (visibleOnly?: booléen)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Obtient la feuille de calcul qui suit celle-ci. S’il n’existe aucune feuille de calcul à la suite de celle-ci, cette méthode renvoie un objet null.|
||[getPrevious (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Obtient la feuille de calcul qui précède celle-ci. S’il n’y a pas de feuille de calcul précédente, cette méthode génère une erreur.|
||[getPreviousOrNullObject (visibleOnly?: booléen)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Obtient la feuille de calcul qui précède celle-ci. S’il n’y a pas de feuille de calcul précédente, cette méthode renvoie une valeur null.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Obtient la première feuille de calcul dans la collection.|
||[getLast (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Obtient la dernière feuille de calcul dans la collection.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
