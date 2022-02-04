---
title: API d’aperçu JavaScript pour Word
description: Détails sur les API JavaScript word à venir.
ms.date: 02/01/2022
ms.prod: word
ms.localizationpriority: medium
---

# <a name="word-javascript-preview-apis"></a>API d’aperçu JavaScript pour Word

Les nouvelles API JavaScript pour Word sont d’abord introduites dans « aperçu », puis font partie d’un ensemble spécifique de conditions requises numérotées une fois que des tests suffisants ont été effectués et que les commentaires des utilisateurs ont été acquis.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Word actuellement en prévisualisation, à l’exception de celles qui sont [disponibles uniquement dans Word sur le web](#web-only-api-list). Pour afficher la liste complète de toutes les API JavaScript pour Word (y compris les API d’aperçu et les API publiées précédemment), consultez toutes les API [JavaScript pour Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Classe | Champs | Description |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member)|Se produit lorsque les données dans le contrôle de contenu sont modifiées.|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member)|Se produit lorsque le contrôle de contenu est supprimé.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member)|Se produit lorsque la sélection dans le contrôle de contenu est modifiée.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|Objet qui a élevé l’événement.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|Type d’événement.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-delete-member(1))|Supprime la partie XML personnalisée.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteattribute-member(1))|Supprime un attribut avec le nom donné de l’élément identifié par xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteelement-member(1))|Supprime l’élément identifié par xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-getxml-member(1))|Obtient le contenu XML complet de la partie XML personnalisée.|
||[id](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-id-member)|Obtient l’ID de la partie XML personnalisée.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertattribute-member(1))|Insère un attribut avec le nom et la valeur donnés à l’élément identifié par xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertelement-member(1))|Insère le XML donné sous l’élément parent identifié par xpath à l’index de position enfant.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespaceuri-member)|Obtient l’URI d’espace de noms de la partie XML personnalisée.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-query-member(1))|Interroge le contenu XML de la partie XML personnalisée.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-setxml-member(1))|Définit le contenu XML complet de la partie XML personnalisée.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateattribute-member(1))|Met à jour la valeur d’un attribut avec le nom donné de l’élément identifié par xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateelement-member(1))|Met à jour le XML de l’élément identifié par xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-add-member(1))|Ajoute une nouvelle partie XML personnalisée au document.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getbynamespace-member(1))|Obtient une nouvelle collection limitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getcount-member(1))|Obtient le nombre d'éléments dans la collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitem-member(1))|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitemornullobject-member(1))|Obtient une partie XML personnalisée en fonction de son ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getcount-member(1))|Obtient le nombre d'éléments dans la collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitem-member(1))|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitemornullobject-member(1))|Obtient une partie XML personnalisée en fonction de son ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitem-member(1))|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#word-word-document-customxmlparts-member)|Obtient les parties XML personnalisées du document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#word-word-document-deletebookmark-member(1))|Supprime un signet, s’il existe, du document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrange-member(1))|Obtient la plage d’un signet.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrangeornullobject-member(1))|Obtient la plage d’un signet.|
||[ignorePunct](/javascript/api/word/word.document#word-word-document-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.document#word-word-document-ignorespace-member)||
||[matchCase](/javascript/api/word/word.document#word-word-document-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.document#word-word-document-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.document#word-word-document-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.document#word-word-document-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.document#word-word-document-matchwildcards-member)||
||[onContentControlAdded](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member)|Se produit lorsqu’un contrôle de contenu est ajouté.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.document#word-word-document-search-member(1))|Effectue une recherche avec les options de recherche spécifiées sur l’étendue du document entier.|
||[paramètres](/javascript/api/word/word.document#word-word-document-settings-member)|Obtient les paramètres du add-in dans le document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#word-word-documentcreated-customxmlparts-member)|Obtient les parties XML personnalisées du document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-deletebookmark-member(1))|Supprime un signet, s’il existe, du document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrange-member(1))|Obtient la plage d’un signet.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrangeornullobject-member(1))|Obtient la plage d’un signet.|
||[paramètres](/javascript/api/word/word.documentcreated#word-word-documentcreated-settings-member)|Obtient les paramètres du add-in dans le document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Obtient le format de l’image fixe.|
|[Liste](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|Obtient la police de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|Obtient la représentation de chaîne codée en base 64 de l’image au niveau spécifié dans la liste.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|Réinitialise la police de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|Définit l’image au niveau spécifié dans la liste.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#word-word-range-getbookmarks-member(1))|Obtient les noms de tous les signets dans la plage ou qui se chevauchent.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#word-word-range-insertbookmark-member(1))|Insère un signet dans la plage.|
|[Paramètre](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#word-word-setting-delete-member(1))|Supprime le paramètre.|
||[key](/javascript/api/word/word.setting#word-word-setting-key-member)|Obtient la clé du paramètre.|
||[value](/javascript/api/word/word.setting#word-word-setting-value-member)|Obtient ou définit la valeur du paramètre.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#word-word-settingcollection-add-member(1))|Crée un nouveau paramètre ou définit un paramètre existant.|
||[deleteAll()](/javascript/api/word/word.settingcollection#word-word-settingcollection-deleteall-member(1))|Supprime tous les paramètres de ce module.|
||[getCount()](/javascript/api/word/word.settingcollection#word-word-settingcollection-getcount-member(1))|Obtient le nombre de paramètres.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitem-member(1))|Obtient un objet de paramètre par sa clé, qui est sensible à la cas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitemornullobject-member(1))|Obtient un objet de paramètre par sa clé, qui est sensible à la cas.|
||[items](/javascript/api/word/word.settingcollection#word-word-settingcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#word-word-table-mergecells-member(1))|Fusionne les cellules délimitées inclusivement par une première et une dernière cellule.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#word-word-tablecell-split-member(1))|Divise la cellule en nombre de lignes et de colonnes spécifié.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Insère un contrôle de contenu sur la ligne.|
||[merge()](/javascript/api/word/word.tablerow#word-word-tablerow-merge-member(1))|Fusionne la ligne en une seule cellule.|

## <a name="web-only-api-list"></a>Liste des API web uniquement

Le tableau suivant répertorie les API JavaScript pour Word actuellement en prévisualisation uniquement dans Word sur le web. Pour afficher la liste complète de toutes les API JavaScript pour Word (y compris les API d’aperçu et les API publiées précédemment), consultez toutes les API [JavaScript pour Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Classe | Champs | Description |
|:---|:---|:---|
|[Corps](/javascript/api/word/word.body)|[notes de fin](/javascript/api/word/word.body#word-word-body-endnotes-member)|Obtient la collection de notes de fin dans le corps.|
||[notes de bas de page](/javascript/api/word/word.body#word-word-body-footnotes-member)|Obtient la collection de notes de bas de page dans le corps.|
||[getComments()](/javascript/api/word/word.body#word-word-body-getcomments-member(1))|Obtient les commentaires associés au corps.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#word-word-body-getreviewedtext-member(1))|Obtient le texte révisé en fonction de la sélection changeTrackingVersion.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Obtient le type du corps.|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#word-word-comment-authoremail-member)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/word/word.comment#word-word-comment-authorname-member)|Obtient le nom de l’auteur du commentaire.|
||[content](/javascript/api/word/word.comment#word-word-comment-content-member)|Obtient ou définit le contenu du commentaire en tant que texte simple.|
||[contentRange](/javascript/api/word/word.comment#word-word-comment-contentrange-member)|Obtient ou définit l’état du thread de commentaire.|
||[creationDate](/javascript/api/word/word.comment#word-word-comment-creationdate-member)|Obtient la date de création du commentaire.|
||[delete()](/javascript/api/word/word.comment#word-word-comment-delete-member(1))|Supprime le commentaire et ses réponses.|
||[getRange()](/javascript/api/word/word.comment#word-word-comment-getrange-member(1))|Obtient la plage du document principal où se trouve le commentaire.|
||[id](/javascript/api/word/word.comment#word-word-comment-id-member)|ID|
||[Réponses](/javascript/api/word/word.comment#word-word-comment-replies-member)|Obtient la collection d’objets de réponse associés au commentaire.|
||[reply(replyText: string)](/javascript/api/word/word.comment#word-word-comment-reply-member(1))|Ajoute une nouvelle réponse à la fin du fil de discussion de commentaires.|
||[résolu](/javascript/api/word/word.comment#word-word-comment-resolved-member)|Obtient ou définit l’état du thread de commentaires.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirst-member(1))|Obtient le premier commentaire de la collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirstornullobject-member(1))|Obtient le premier commentaire de la collection.|
||[getItem(index : numérique)](/javascript/api/word/word.commentcollection#word-word-commentcollection-getitem-member(1))|Obtient un objet comment par son index dans la collection.|
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-bold-member)|Obtient ou définit une valeur qui indique si le texte du commentaire est en gras.|
||[lien hypertexte](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-hyperlink-member)|Obtient le premier lien hypertexte de la plage ou définit un lien hypertexte sur la plage.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-inserttext-member(1))|Insère du texte à l’emplacement spécifié.|
||[isEmpty](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-isempty-member)|Vérifie si la longueur de la plage est zéro.|
||[italic](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-italic-member)|Obtient ou définit une valeur qui indique si le texte du commentaire est en italique.|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-strikethrough-member)|Obtient ou définit une valeur qui indique si le texte du commentaire a un signet.|
||[text](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-text-member)|Obtient le texte de la plage de commentaires.|
||[underline](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-underline-member)|Obtient ou définit une valeur qui indique le type de soulignement du texte du commentaire.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#word-word-commentreply-authoremail-member)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/word/word.commentreply#word-word-commentreply-authorname-member)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[content](/javascript/api/word/word.commentreply#word-word-commentreply-content-member)|Obtient ou définit le contenu de la réponse au commentaire.|
||[contentRange](/javascript/api/word/word.commentreply#word-word-commentreply-contentrange-member)|Obtient ou définit la plage de contenu de commentReply.|
||[creationDate](/javascript/api/word/word.commentreply#word-word-commentreply-creationdate-member)|Obtient la date de création de la réponse au commentaire.|
||[delete()](/javascript/api/word/word.commentreply#word-word-commentreply-delete-member(1))|Supprime la réponse de commentaire.|
||[id](/javascript/api/word/word.commentreply#word-word-commentreply-id-member)|ID|
||[parentComment](/javascript/api/word/word.commentreply#word-word-commentreply-parentcomment-member)|Obtient le commentaire parent de cette réponse.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirst-member(1))|Obtient la première réponse de commentaire dans la collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirstornullobject-member(1))|Obtient la première réponse de commentaire dans la collection.|
||[getItem(index : numérique)](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getitem-member(1))|Obtient un objet de réponse de commentaire par son index dans la collection.|
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[notes de fin](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|Obtient la collection de notes de fin dans le contentcontrol.|
||[notes de bas de page](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|Obtient la collection de notes de bas de page dans le contentcontrol.|
||[getComments()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcomments-member(1))|Obtient les commentaires associés au corps.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getreviewedtext-member(1))|Obtient le texte révisé en fonction de la sélection changeTrackingVersion.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member)|Obtient ou définit le mode ChangeTracking.|
||[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|Obtient les notes de fin du document dans un corps unique.|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|Obtient les notes de bas de page du document dans un corps unique.|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|Représente l’objet body de l’élément de note.|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|Supprime l’élément de note.|
||[getNext()](/javascript/api/word/word.noteitem#word-word-noteitem-getnext-member(1))|Obtient l’élément de note suivant du même type.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#word-word-noteitem-getnextornullobject-member(1))|Obtient l’élément de note suivant du même type.|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|Représente une référence de note de bas de page ou de note de fin dans le document principal.|
||[type](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|Représente le type d’élément de note : note de bas de page ou note de fin.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirst-member(1))|Obtient le premier élément de note de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirstornullobject-member(1))|Obtient le premier élément de note de cette collection.|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Paragraph](/javascript/api/word/word.paragraph)|[notes de fin](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|Obtient la collection de notes de fin du paragraphe.|
||[notes de bas de page](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|Obtient la collection de notes de bas de page du paragraphe.|
||[getComments()](/javascript/api/word/word.paragraph#word-word-paragraph-getcomments-member(1))|Obtient les commentaires associés au paragraphe.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#word-word-paragraph-getreviewedtext-member(1))|Obtient le texte révisé en fonction de la sélection changeTrackingVersion.|
|[Range](/javascript/api/word/word.range)|[notes de fin](/javascript/api/word/word.range#word-word-range-endnotes-member)|Obtient la collection de notes de fin de la plage.|
||[notes de bas de page](/javascript/api/word/word.range#word-word-range-footnotes-member)|Obtient la collection de notes de bas de page de la plage.|
||[getComments()](/javascript/api/word/word.range#word-word-range-getcomments-member(1))|Obtient les commentaires associés à la plage.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#word-word-range-getreviewedtext-member(1))|Obtient le texte révisé en fonction de la sélection changeTrackingVersion.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#word-word-range-insertcomment-member(1))|Insérez un commentaire sur la plage.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|Insère une note de fin.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|Insère une note de bas de page.|
|[Table](/javascript/api/word/word.table)|[notes de fin](/javascript/api/word/word.table#word-word-table-endnotes-member)|Obtient la collection de notes de fin du tableau.|
||[notes de bas de page](/javascript/api/word/word.table#word-word-table-footnotes-member)|Obtient la collection de notes de bas de page du tableau.|
|[TableRow](/javascript/api/word/word.tablerow)|[notes de fin](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|Obtient la collection de notes de fin dans la ligne du tableau.|
||[notes de bas de page](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|Obtient la collection de notes de bas de page dans la ligne du tableau.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
