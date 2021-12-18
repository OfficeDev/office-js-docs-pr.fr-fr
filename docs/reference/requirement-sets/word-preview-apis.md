---
title: API d’aperçu JavaScript pour Word
description: Détails sur les API JavaScript word à venir.
ms.date: 12/14/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: c68a63dc57fbcaa8282343c3f3271778c43bc28d
ms.sourcegitcommit: 9b6556563451f9907cb5da50cba757eb9960aa39
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/17/2021
ms.locfileid: "61565363"
---
# <a name="word-javascript-preview-apis"></a>API d’aperçu JavaScript pour Word

Les nouvelles API JavaScript pour Word sont d’abord introduites dans « aperçu », puis font partie d’un ensemble spécifique de conditions requises numérotées une fois que des tests suffisants ont été effectués et que les commentaires des utilisateurs ont été acquis.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Word actuellement en prévisualisation, à l’exception de celles qui sont disponibles uniquement [dans Word sur le web](#web-only-api-list). Pour afficher la liste complète de toutes les API JavaScript pour Word (y compris les API d’aperçu et les API publiées précédemment), consultez toutes les API [JavaScript pour Word.](/javascript/api/word?view=word-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|Se produit lorsque les données dans le contrôle de contenu sont modifiées.|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|Se produit lorsque le contrôle de contenu est supprimé.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|Se produit lorsque la sélection dans le contrôle de contenu est modifiée.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|Objet qui a élevé l’événement.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|Type d’événement.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|Supprime la partie XML personnalisée.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|Supprime un attribut avec le nom donné de l’élément identifié par xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|Supprime l’élément identifié par xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getXml__)|Obtient le contenu XML complet de la partie XML personnalisée.|
||[id](/javascript/api/word/word.customxmlpart#id)|Obtient l’ID de la partie XML personnalisée.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|Insère un attribut avec le nom et la valeur donnés à l’élément identifié par xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|Insère le XML donné sous l’élément parent identifié par xpath à l’index de position enfant.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|Obtient l’URI d’espace de noms de la partie XML personnalisée.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|Interroge le contenu XML de la partie XML personnalisée.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setXml_xml_)|Définit le contenu XML complet de la partie XML personnalisée.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|Met à jour la valeur d’un attribut avec le nom donné de l’élément identifié par xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|Met à jour le XML de l’élément identifié par xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add_xml_)|Ajoute une nouvelle partie XML personnalisée au document.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|Obtient une nouvelle collection limitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|Obtient le nombre d'éléments dans la collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|Obtient une partie XML personnalisée en fonction de son ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|Obtient le nombre d'éléments dans la collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#customXmlParts)|Obtient les parties XML personnalisées du document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#deleteBookmark_name_)|Supprime un signet, s’il existe, du document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getBookmarkRange_name_)|Obtient la plage d’un signet.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|Obtient la plage d’un signet.|
||[ignorePunct](/javascript/api/word/word.document#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.document#ignoreSpace)||
||[matchCase](/javascript/api/word/word.document#matchCase)||
||[matchPrefix](/javascript/api/word/word.document#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.document#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.document#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.document#matchWildcards)||
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|Se produit lorsqu’un contrôle de contenu est ajouté.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.document#search_searchText__searchOptions_)|Effectue une recherche avec les options de recherche spécifiées sur l’étendue du document entier.|
||[paramètres](/javascript/api/word/word.document#settings)|Obtient les paramètres du add-in dans le document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|Obtient les parties XML personnalisées du document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|Supprime un signet, s’il existe, du document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|Obtient la plage d’un signet.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|Obtient la plage d’un signet.|
||[paramètres](/javascript/api/word/word.documentcreated#settings)|Obtient les paramètres du add-in dans le document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|Obtient le format de l’image fixe.|
|[Liste](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getLevelFont_level_)|Obtient la police de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getLevelPicture_level_)|Obtient la représentation de chaîne codée en base 64 de l’image au niveau spécifié dans la liste.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|Réinitialise la police de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|Définit l’image au niveau spécifié dans la liste.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|Obtient les noms de tous les signets dans la plage ou qui se chevauchent.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertBookmark_name_)|Insère un signet dans la plage.|
|[Paramètre](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|Supprime le paramètre.|
||[key](/javascript/api/word/word.setting#key)|Obtient la clé du paramètre.|
||[value](/javascript/api/word/word.setting#value)|Obtient ou définit la valeur du paramètre.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add_key__value_)|Crée un nouveau paramètre ou définit un paramètre existant.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteAll__)|Supprime tous les paramètres de ce module.|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|Obtient le nombre de paramètres.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|Obtient un objet de paramètre par sa clé, qui est sensible à la cas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|Obtient un objet de paramètre par sa clé, qui est sensible à la cas.|
||[items](/javascript/api/word/word.settingcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|Fusionne les cellules délimitées inclusivement par une première et une dernière cellule.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|Divise la cellule en nombre de lignes et de colonnes spécifié.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|Insère un contrôle de contenu sur la ligne.|
||[merge()](/javascript/api/word/word.tablerow#merge__)|Fusionne la ligne en une seule cellule.|

## <a name="web-only-api-list"></a>Liste des API web uniquement

Le tableau suivant répertorie les API JavaScript pour Word actuellement en prévisualisation uniquement dans Word sur le web. Pour afficher la liste complète de toutes les API JavaScript pour Word (y compris les API d’aperçu et les API publiées précédemment), consultez toutes les API [JavaScript pour Word.](/javascript/api/word?view=word-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[Corps](/javascript/api/word/word.body)|[notes de fin](/javascript/api/word/word.body#endnotes)|Obtient la collection de notes de fin dans le corps.|
||[notes de bas de page](/javascript/api/word/word.body#footnotes)|Obtient la collection de notes de bas de page dans le corps.|
||[getComments()](/javascript/api/word/word.body#getComments__)|Obtient les commentaires associés au corps.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#getReviewedText_changeTrackingVersion_)|Obtient le texte révisé en fonction de la sélection changeTrackingVersion.|
||[type](/javascript/api/word/word.body#type)|Obtient le type du corps.|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#authorEmail)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/word/word.comment#authorName)|Obtient le nom de l’auteur du commentaire.|
||[content](/javascript/api/word/word.comment#content)|Obtient ou définit le contenu du commentaire en tant que texte simple.|
||[creationDate](/javascript/api/word/word.comment#creationDate)|Obtient la date de création du commentaire.|
||[delete()](/javascript/api/word/word.comment#delete__)|Supprime le commentaire et ses réponses.|
||[getRange()](/javascript/api/word/word.comment#getRange__)|Obtient la plage du document principal où se trouve le commentaire.|
||[id](/javascript/api/word/word.comment#id)|ID|
||[Réponses](/javascript/api/word/word.comment#replies)|Obtient la collection d’objets de réponse associés au commentaire.|
||[reply(replyText: string)](/javascript/api/word/word.comment#reply_replyText_)|Ajoute une nouvelle réponse à la fin du fil de discussion de commentaires.|
||[résolu](/javascript/api/word/word.comment#resolved)|Obtient ou définit l’état du thread de commentaire.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#getFirst__)|Obtient le premier commentaire de la collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#getFirstOrNullObject__)|Obtient le premier commentaire ou objet null de la collection.|
||[getItem(index : numérique)](/javascript/api/word/word.commentcollection#getItem_index_)|Obtient un objet comment par son index dans la collection.|
||[items](/javascript/api/word/word.commentcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#authorEmail)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/word/word.commentreply#authorName)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[content](/javascript/api/word/word.commentreply#content)|Obtient ou définit le contenu de la réponse au commentaire.|
||[creationDate](/javascript/api/word/word.commentreply#creationDate)|Obtient la date de création de la réponse au commentaire.|
||[delete()](/javascript/api/word/word.commentreply#delete__)|Supprime la réponse de commentaire.|
||[id](/javascript/api/word/word.commentreply#id)|ID|
||[parentComment](/javascript/api/word/word.commentreply#parentComment)|Obtient le commentaire parent de cette réponse.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#getFirst__)|Obtient la première réponse de commentaire dans la collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#getFirstOrNullObject__)|Obtient le premier objet de réponse de commentaire ou null de la collection.|
||[getItem(index : numérique)](/javascript/api/word/word.commentreplycollection#getItem_index_)|Obtient un objet de réponse de commentaire par son index dans la collection.|
||[items](/javascript/api/word/word.commentreplycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[notes de fin](/javascript/api/word/word.contentcontrol#endnotes)|Obtient la collection de notes de fin dans le contentcontrol.|
||[notes de bas de page](/javascript/api/word/word.contentcontrol#footnotes)|Obtient la collection de notes de bas de page dans le contentcontrol.|
||[getComments()](/javascript/api/word/word.contentcontrol#getComments__)|Obtient les commentaires associés au corps.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#getReviewedText_changeTrackingVersion_)|Obtient le texte révisé en fonction de la sélection changeTrackingVersion.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#changeTrackingMode)|Obtient ou définit le mode ChangeTracking.|
||[getEndnoteBody()](/javascript/api/word/word.document#getEndnoteBody__)|Obtient les notes de fin du document dans un corps unique.|
||[getFootnoteBody()](/javascript/api/word/word.document#getFootnoteBody__)|Obtient les notes de bas de page du document dans un corps unique.|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#body)|Représente l’objet body de l’élément de note.|
||[delete()](/javascript/api/word/word.noteitem#delete__)|Supprime l’élément de note.|
||[getNext()](/javascript/api/word/word.noteitem#getNext__)|Obtient l’élément de note suivant du même type.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#getNextOrNullObject__)|Obtient l’élément de note suivant du même type.|
||[reference](/javascript/api/word/word.noteitem#reference)|Représente une référence de note de bas de page ou de note de fin dans le document principal.|
||[type](/javascript/api/word/word.noteitem#type)|Représente le type d’élément de note : note de bas de page ou note de fin.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#getFirst__)|Obtient le premier élément de note de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#getFirstOrNullObject__)|Obtient le premier élément de note de cette collection.|
||[items](/javascript/api/word/word.noteitemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Paragraph](/javascript/api/word/word.paragraph)|[notes de fin](/javascript/api/word/word.paragraph#endnotes)|Obtient la collection de notes de fin du paragraphe.|
||[notes de bas de page](/javascript/api/word/word.paragraph#footnotes)|Obtient la collection de notes de bas de page du paragraphe.|
||[getComments()](/javascript/api/word/word.paragraph#getComments__)|Obtient les commentaires associés au paragraphe.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#getReviewedText_changeTrackingVersion_)|Obtient le texte révisé en fonction de la sélection changeTrackingVersion.|
|[Range](/javascript/api/word/word.range)|[notes de fin](/javascript/api/word/word.range#endnotes)|Obtient la collection de notes de fin de la plage.|
||[notes de bas de page](/javascript/api/word/word.range#footnotes)|Obtient la collection de notes de bas de page de la plage.|
||[getComments()](/javascript/api/word/word.range#getComments__)|Obtient les commentaires associés à la plage.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#getReviewedText_changeTrackingVersion_)|Obtient le texte révisé en fonction de la sélection changeTrackingVersion.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#insertComment_commentText_)|Insérez un commentaire sur la plage.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#insertEndnote_insertText_)|Insère une note de fin.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#insertFootnote_insertText_)|Insère une note de bas de page.|
|[Table](/javascript/api/word/word.table)|[notes de fin](/javascript/api/word/word.table#endnotes)|Obtient la collection de notes de fin du tableau.|
||[notes de bas de page](/javascript/api/word/word.table#footnotes)|Obtient la collection de notes de bas de page du tableau.|
|[TableRow](/javascript/api/word/word.tablerow)|[notes de fin](/javascript/api/word/word.tablerow#endnotes)|Obtient la collection de notes de fin dans la ligne du tableau.|
||[notes de bas de page](/javascript/api/word/word.tablerow#footnotes)|Obtient la collection de notes de bas de page dans la ligne du tableau.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
