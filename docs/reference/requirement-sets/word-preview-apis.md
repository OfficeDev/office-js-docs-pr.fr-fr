---
title: API d’aperçu JavaScript pour Word
description: Détails sur les API JavaScript word à venir
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: c6aa7b8107e0443091f876baa8bd66ccb8db7061
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153071"
---
# <a name="word-javascript-preview-apis"></a>API d’aperçu JavaScript pour Word

Les nouvelles API JavaScript pour Word sont d’abord introduites dans « aperçu », puis font partie d’un ensemble spécifique de conditions requises numérotées une fois que des tests suffisants ont été effectués et que les commentaires des utilisateurs ont été acquis.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Word actuellement en prévisualisation. Pour afficher la liste complète de toutes les API JavaScript pour Word (y compris les API d’aperçu et les API publiées précédemment), consultez toutes les API [JavaScript pour Word.](/javascript/api/word?view=word-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Se produit lorsque les données dans le contrôle de contenu sont modifiées.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Se produit lorsque le contrôle de contenu est supprimé.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Se produit lorsque la sélection dans le contrôle de contenu est modifiée.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|Objet qui a levé l’événement.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|Type d’événement.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Supprime la partie XML personnalisée.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Supprime un attribut avec le nom donné de l’élément identifié par xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Supprime l’élément identifié par xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getxml--)|Obtient le contenu XML complet de la partie XML personnalisée.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Insère un attribut avec le nom et la valeur donnés à l’élément identifié par xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Insère le XML donné sous l’élément parent identifié par xpath à l’index de position enfant.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Interroge le contenu XML de la partie XML personnalisée.|
||[id](/javascript/api/word/word.customxmlpart#id)|Obtient l’ID de la partie XML personnalisée.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceuri)|Obtient l’URI d’espace de noms de la partie XML personnalisée.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Définit le contenu XML complet de la partie XML personnalisée.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Met à jour la valeur d’un attribut avec le nom donné de l’élément identifié par xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Met à jour le XML de l’élément identifié par xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Ajoute une nouvelle partie XML personnalisée au document.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtient une nouvelle collection limitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|Obtient le nombre d'éléments dans la collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|Obtient le nombre d'éléments dans la collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark(name: string)](/javascript/api/word/word.document#deletebookmark-name-)|Supprime un signet, s’il existe, du document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getbookmarkrange-name-)|Obtient la plage d’un signet.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Obtient la plage d’un signet.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Obtient les parties XML personnalisées du document.|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|Se produit lorsqu’un contrôle de contenu est ajouté.|
||[paramètres](/javascript/api/word/word.document#settings)|Obtient les paramètres du add-in dans le document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Supprime un signet, s’il existe, du document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Obtient la plage d’un signet.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Obtient la plage d’un signet.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Obtient les parties XML personnalisées du document.|
||[paramètres](/javascript/api/word/word.documentcreated#settings)|Obtient les paramètres du add-in dans le document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|Obtient le format de l’image fixe.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getlevelfont-level-)|Obtient la police de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getlevelpicture-level-)|Obtient la représentation de chaîne codée en base 64 de l’image au niveau spécifié dans la liste.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Réinitialise la police de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Définit l’image au niveau spécifié dans la liste.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Obtient les noms de tous les signets dans la plage ou qui se chevauchent.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertbookmark-name-)|Insère un signet sur la plage.|
|[Paramètre](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Supprime le paramètre.|
||[key](/javascript/api/word/word.setting#key)|Obtient la clé du paramètre.|
||[value](/javascript/api/word/word.setting#value)|Obtient ou définit la valeur du paramètre.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|Crée un nouveau paramètre ou définit un paramètre existant.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteall--)|Supprime tous les paramètres de ce module.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Obtient le nombre de paramètres.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Obtient un objet de paramètre par sa clé, qui est sensible à la cas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Obtient un objet de paramètre par sa clé, qui est sensible à la cas.|
||[items](/javascript/api/word/word.settingcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Fusionne les cellules délimitées inclusivement par une première et une dernière cellule.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Divise la cellule en nombre de lignes et de colonnes spécifié.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Insère un contrôle de contenu sur la ligne.|
||[merge()](/javascript/api/word/word.tablerow#merge--)|Fusionne la ligne dans une cellule.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
