---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,5
description: Détails sur l’ensemble de conditions requises ExcelApi 1,5.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 901ea29253bdfee3aeeefd1595eda32d70e4f1e1
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996255"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Nouveautés de l’API JavaScript 1.5 pour Excel

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

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Excel 1,5. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1,5 ou antérieure, voir [API Excel dans l’ensemble de conditions requises 1,5 ou version antérieure](/javascript/api/excel?view=excel-js-1.5&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Supprime la partie XML personnalisée.|
||[getXml ()](/javascript/api/excel/excel.customxmlpart#getxml--)|Obtient l’intégralité du contenu XML de la partie XML personnalisée.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|ID de la partie XML personnalisée.|
||[URI](/javascript/api/excel/excel.customxmlpart#namespaceuri)|URI de l’espace de noms de la partie XML personnalisée.|
||[setXml (XML : chaîne)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Définit l’intégralité du contenu XML de la partie XML personnalisée.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[Add (XML : String)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Ajoute une nouvelle partie XML personnalisée au classeur.|
||[getByNamespace (namespaceUri : String)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtient une nouvelle collection limitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Obtient le nombre de parties CustomXml dans la collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Obtient le nombre de parties CustomXML dans cette collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|ID du tableau croisé dynamique.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)|[Ensemble d’API : ExcelApi 1,5]|
|[Runtime](/javascript/api/excel/excel.runtime)||[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Représente la collection de parties XML personnalisées contenues dans ce classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly ?: Boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Obtient la feuille de calcul qui suit celle-ci.|
||[getNextOrNullObject (visibleOnly ?: booléen)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Obtient la feuille de calcul qui suit celle-ci.|
||[getPrevious (visibleOnly ?: Boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Obtient la feuille de calcul qui précède celle-ci.|
||[getPreviousOrNullObject (visibleOnly ?: booléen)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Obtient la feuille de calcul qui précède celle-ci.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly ?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Obtient la première feuille de calcul dans la collection.|
||[getLast (visibleOnly ?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Obtient la dernière feuille de calcul dans la collection.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
