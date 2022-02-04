---
title: Excel l’ensemble de conditions requises de l’API JavaScript 1.5
description: Détails sur l’ensemble de conditions requises ExcelApi 1.5.
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-15"></a>Nouveautés de l’API JavaScript 1.5 pour Excel

ExcelApi 1.5 ajoute des parties XML personnalisées. Ceux-ci sont accessibles via [la collection de parties XML personnalisée](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member) dans l’objet debook.

## <a name="custom-xml-part"></a>Partie XML personnalisée

* Obtenir des parties XML personnalisées à l’aide de leur ID.
* Obtenez une nouvelle collection délimitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.
* Obtenez une chaîne XML associée à un élément.
* Fournissez l’ID et l’espace de noms d’un élément.
* Ajoutez une nouvelle partie XML personnalisée au workbook.
* Définissez une partie XML entière.
* Supprimez une partie XML personnalisée.
* Supprimez un attribut avec le nom donné dans l’élément identifié par langage XPath.
* Interrogez le contenu XML par langage XPath.
* Insérer, mettre à jour et supprimer des attributs.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.5. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.5 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.5](/javascript/api/excel?view=excel-js-1.5&preserve-view=true) ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-delete-member(1))|Supprime la partie XML personnalisée.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-getxml-member(1))|Obtient l’intégralité du contenu XML de la partie XML personnalisée.|
||[id](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-id-member)|ID de la partie XML personnalisée.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-namespaceuri-member)|URI d’espace de noms de la partie XML personnalisée.|
||[setXml(xml: string)](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-setxml-member(1))|Définit l’intégralité du contenu XML de la partie XML personnalisée.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-add-member(1))|Ajoute une nouvelle partie XML personnalisée au classeur.|
||[getByNamespace(namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getbynamespace-member(1))|Obtient une nouvelle collection limitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getcount-member(1))|Obtient le nombre de parties XML personnalisées dans la collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getitem-member(1))|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getitemornullobject-member(1))|Obtient une partie XML personnalisée en fonction de son ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getcount-member(1))|Obtient le nombre de parties CustomXML dans cette collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getitem-member(1))|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getitemornullobject-member(1))|Obtient une partie XML personnalisée en fonction de son ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getonlyitem-member(1))|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-id-member)|ID du tableau croisé dynamique.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#excel-excel-requestcontext-runtime-member)||
|[Runtime](/javascript/api/excel/excel.runtime)|||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member)|Représente la collection de parties XML personnalisées contenues dans ce manuel.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext(visibleOnly?: booléen)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getnext-member(1))|Obtient la feuille de calcul qui suit celle-ci.|
||[getNextOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getnextornullobject-member(1))|Obtient la feuille de calcul qui suit celle-ci.|
||[getPrevious(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getprevious-member(1))|Obtient la feuille de calcul qui précède celle-ci.|
||[getPreviousOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getpreviousornullobject-member(1))|Obtient la feuille de calcul qui précède celle-ci.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst(visibleOnly?: booléen)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getfirst-member(1))|Obtient la première feuille de calcul dans la collection.|
||[getLast(visibleOnly?: booléen)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getlast-member(1))|Obtient la dernière feuille de calcul dans la collection.|

## <a name="see-also"></a>Voir aussi

* [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
