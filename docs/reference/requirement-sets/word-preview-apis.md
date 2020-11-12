---
title: API de prévisualisation JavaScript pour Word
description: Informations détaillées sur les API JavaScript pour Word à venir
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 6a3b67e65c4ced3f1b89d98afe45d5d6c33f63b6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996402"
---
# <a name="word-javascript-preview-apis"></a>API de prévisualisation JavaScript pour Word

De nouvelles API JavaScript pour Word sont introduites pour la première fois dans « Preview », puis elles deviennent une partie d’un ensemble de conditions requises spécifiques, après un test suffisant, et les commentaires des utilisateurs sont acquis.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Word actuellement en version préliminaire. Pour afficher la liste complète de toutes les API JavaScript Word (y compris les API d’aperçu et les API précédemment publiées), voir [toutes les API JavaScript pour Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Se produit lors de la modification de données dans le contrôle de contenu.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Se produit lorsque le contrôle de contenu est supprimé.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Se produit lors de la modification de la sélection dans le contrôle de contenu.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|Objet qui a déclenché l’événement.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|Type d’événement.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Supprime la partie XML personnalisée.|
||[deleteAttribute (XPath : String, namespaceMappings : any, Name : String)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Supprime un attribut portant le nom donné à partir de l’élément identifié par XPath.|
||[deleteElement (XPath : String, namespaceMappings : any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Supprime l’élément identifié par XPath.|
||[getXml ()](/javascript/api/word/word.customxmlpart#getxml--)|Obtient le contenu XML complet de la partie XML personnalisée.|
||[insertAttribute (XPath : String, namespaceMappings : any, Name : String, value : String)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Insère un attribut avec le nom et la valeur spécifiés pour l’élément identifié par XPath.|
||[insertElement (XPath : String, XML : String, namespaceMappings : any, index ?: Number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Insère le code XML donné sous l’élément parent identifié par XPath à l’index de position enfant.|
||[requête (XPath : String, namespaceMappings : any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Interroge le contenu XML de la partie XML personnalisée.|
||[id](/javascript/api/word/word.customxmlpart#id)|Obtient l’ID de la partie XML personnalisée.|
||[URI](/javascript/api/word/word.customxmlpart#namespaceuri)|Obtient l’URI de l’espace de noms de la partie XML personnalisée.|
||[setXml (XML : chaîne)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Définit le contenu XML complet de la partie XML personnalisée.|
||[updateAttribute (XPath : String, namespaceMappings : any, Name : String, value : String)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Met à jour la valeur d’un attribut avec le nom donné de l’élément identifié par XPath.|
||[updateElement (XPath : String, XML : String, namespaceMappings : any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Met à jour le code XML de l’élément identifié par XPath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[Add (XML : String)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Ajoute une nouvelle partie XML personnalisée dans le document.|
||[getByNamespace (namespaceUri : String)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtient une nouvelle collection limitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.|
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
|[Document](/javascript/api/word/word.document)|[deleteBookmark (Name : chaîne)](/javascript/api/word/word.document#deletebookmark-name-)|Supprime un signet, s’il existe, du document.|
||[getBookmarkRange (Name : chaîne)](/javascript/api/word/word.document#getbookmarkrange-name-)|Obtient la plage d’un signet.|
||[getBookmarkRangeOrNullObject (Name : chaîne)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Obtient la plage d’un signet.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Obtient les parties XML personnalisées dans le document.|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|Se produit lors de l’ajout d’un contrôle de contenu.|
||[paramètres](/javascript/api/word/word.document#settings)|Obtient les paramètres du complément dans le document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark (Name : chaîne)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Supprime un signet, s’il existe, du document.|
||[getBookmarkRange (Name : chaîne)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Obtient la plage d’un signet.|
||[getBookmarkRangeOrNullObject (Name : chaîne)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Obtient la plage d’un signet.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Obtient les parties XML personnalisées dans le document.|
||[paramètres](/javascript/api/word/word.documentcreated#settings)|Obtient les paramètres du complément dans le document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|Obtient le format de l’image incluse.|
|[List](/javascript/api/word/word.list)|[getLevelFont (Level : nombre)](/javascript/api/word/word.list#getlevelfont-level-)|Obtient la police de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[getLevelPicture (Level : nombre)](/javascript/api/word/word.list#getlevelpicture-level-)|Obtient la représentation sous forme de chaîne codée en base64 de l’image au niveau spécifié dans la liste.|
||[resetLevelFont (Level : nombre, resetFontName ?: booléen)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Rétablit la police de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[setLevelPicture (Level : nombre, base64EncodedImage ?: chaîne)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Définit l’image au niveau spécifié dans la liste.|
|[Range](/javascript/api/word/word.range)|[getBookmarks (includeHidden ?: Boolean, includeAdjacent ?: Boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Obtient le nom de tous les signets dans la plage ou qui se chevauchent.|
||[insertBookmark (Name : chaîne)](/javascript/api/word/word.range#insertbookmark-name-)|Insère un signet dans la plage.|
|[Paramètre](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Supprime le paramètre.|
||[key](/javascript/api/word/word.setting#key)|Obtient la clé du paramètre.|
||[value](/javascript/api/word/word.setting#value)|Obtient ou définit la valeur du paramètre.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[Add (Key : chaîne, value : any)](/javascript/api/word/word.settingcollection#add-key--value-)|Crée un nouveau paramètre ou définit un paramètre existant.|
||[deleteAll ()](/javascript/api/word/word.settingcollection#deleteall--)|Supprime tous les paramètres de ce complément.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Obtient le nombre de paramètres.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Obtient un objet Setting par sa clé, qui respecte la casse.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Obtient un objet Setting par sa clé, qui respecte la casse.|
||[items](/javascript/api/word/word.settingcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/word/word.table)|[mergeCells (topRow : nombre, firstCell : nombre, bottomRow : nombre, lastCell : nombre)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Cette méthode fusionne les cellules délimitées de façon inclusive par une première et la dernière cellule.|
|[TableCell](/javascript/api/word/word.tablecell)|[Split (rowCount : nombre, columnCount : nombre)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Divise la cellule en un nombre spécifié de lignes et de colonnes.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Insère un contrôle de contenu sur la ligne.|
||[Merge ()](/javascript/api/word/word.tablerow#merge--)|Cette méthode fusionne la ligne dans une seule cellule.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
