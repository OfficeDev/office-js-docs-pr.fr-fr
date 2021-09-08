---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1.3
description: Détails sur l’ensemble de conditions requises WordApi 1.3.
ms.date: 03/09/2021
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: b58bb99e664e982d1d9047f4348755d807ad216d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938920"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Word

WordApi 1.3 a ajouté une prise en charge supplémentaire des contrôles de contenu et des paramètres au niveau du document.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de l’ensemble de conditions requises de l’API JavaScript pour Word 1.3. Pour afficher la documentation de référence de l’API pour toutes les API pris en charge par l’ensemble de conditions requises de l’API JavaScript pour Word 1.3 ou une version antérieure, voir API Word dans l’ensemble de conditions requises [1.3](/javascript/api/word?view=word-js-1.3&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createDocument_base64File_)|Crée un document à l’aide d’un fichier .docx encodé en base 64 facultatif.|
|[Corps](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getRange_rangeLocation_)|Obtient la totalité du corps, ou le point de début ou de fin du corps, sous la forme d’une plage.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#insertTable_rowCount__columnCount__insertLocation__values_)|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[lists](/javascript/api/word/word.body#lists)|Obtient la collection d’objets list dans le corps.|
||[parentBody](/javascript/api/word/word.body#parentBody)|Obtient le corps parent du corps.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentBodyOrNullObject)|Obtient le corps parent du corps.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentContentControlOrNullObject)|Obtient le contrôle de contenu qui contient le corps.|
||[parentSection](/javascript/api/word/word.body#parentSection)|Obtient la section parent du corps.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentSectionOrNullObject)|Obtient la section parent du corps.|
||[tables](/javascript/api/word/word.body#tables)|Obtient la collection d’objets table dans le corps.|
||[type](/javascript/api/word/word.body#type)|Obtient le type du corps.|
||[styleBuiltIn](/javascript/api/word/word.body#styleBuiltIn)|Obtient ou définit le nom de style intégré pour le corps.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getRange_rangeLocation_)|Obtient le contrôle de contenu entier, ou le point de début ou de fin du contrôle de contenu, sous la forme d’une plage.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#getTextRanges_endingMarks__trimSpacing_)|Obtient les plages de texte dans le contrôle de contenu à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#insertTable_rowCount__columnCount__insertLocation__values_)|Insère un tableau avec le nombre spécifié de lignes et de colonnes dans un contrôle de contenu ou à proximité de celui-ci.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Obtient la collection d’objets list du contrôle de contenu.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentBody)|Obtient le corps parent du contrôle de contenu.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentContentControlOrNullObject)|Obtient le contrôle de contenu qui contient le contrôle de contenu spécifié.|
||[parentTable](/javascript/api/word/word.contentcontrol#parentTable)|Obtient le tableau qui contient le contrôle de contenu.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parentTableCell)|Obtient la cellule de tableau qui contient le contrôle de contenu.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parentTableCellOrNullObject)|Obtient la cellule de tableau qui contient le contrôle de contenu.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parentTableOrNullObject)|Obtient le tableau qui contient le contrôle de contenu.|
||[sous-type](/javascript/api/word/word.contentcontrol#subtype)|Obtient le sous-type du contrôle de contenu.|
||[tables](/javascript/api/word/word.contentcontrol#tables)|Obtient la collection d’objets table du contrôle de contenu.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|Fractionne le contrôle de contenu en plages enfants à l’aide de délimiteurs.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#styleBuiltIn)|Obtient ou définit le nom de style intégré pour le contrôle de contenu.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#getByIdOrNullObject_id_)|Obtient un contrôle de contenu par son identificateur.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getByTypes_types_)|Obtient les contrôles de contenu qui ont les types et/ou sous-types spécifiés.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getFirst__)|Obtient le premier contrôle de contenu de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getFirstOrNullObject__)|Obtient le premier contrôle de contenu de cette collection.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete__)|Supprime la propriété personnalisée.|
||[key](/javascript/api/word/word.customproperty#key)|Obtient la clé de la propriété personnalisée.|
||[type](/javascript/api/word/word.customproperty#type)|Obtient le type de valeur de la propriété personnalisée.|
||[value](/javascript/api/word/word.customproperty#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add_key__value_)|Crée une nouvelle propriété personnalisée ou en définit une existante.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteAll__)|Supprime toutes les propriétés personnalisées de cette collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getCount__)|Obtient le nombre des propriétés personnalisées.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getItem_key_)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getItemOrNullObject_key_)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Obtient les propriétés du document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open__)|Ouvre le document.|
||[body](/javascript/api/word/word.documentcreated#body)|Obtient l’objet body du document.|
||[contentControls](/javascript/api/word/word.documentcreated#contentControls)|Obtient la collection d’objets de contrôle de contenu dans le document.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Obtient les propriétés du document.|
||[saved](/javascript/api/word/word.documentcreated#saved)|Indique si les modifications apportées au document ont été enregistrées.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Obtient la collection d’objets de section dans le document.|
||[save()](/javascript/api/word/word.documentcreated#save__)|Enregistre le document.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[author](/javascript/api/word/word.documentproperties#author)|Obtient ou définit l’auteur du document.|
||[category](/javascript/api/word/word.documentproperties#category)|Obtient ou définit la catégorie du document.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Obtient ou définit les commentaires du document.|
||[company](/javascript/api/word/word.documentproperties#company)|Obtient ou définit la société du document.|
||[format](/javascript/api/word/word.documentproperties#format)|Obtient ou définit le format du document.|
||[keywords](/javascript/api/word/word.documentproperties#keywords)|Obtient ou définit les mots clés du document.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Obtient ou définit le responsable du document.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationName)|Obtient le nom d’application du document.|
||[creationDate](/javascript/api/word/word.documentproperties#creationDate)|Obtient la date de création du document.|
||[customProperties](/javascript/api/word/word.documentproperties#customProperties)|Obtient la collection de propriétés personnalisées du document.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastAuthor)|Obtient le dernier auteur du document.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastPrintDate)|Obtient la dernière date d’impression du document.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastSaveTime)|Obtient la dernière heure d’enregistrement du document.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionNumber)|Obtient le numéro de révision du document.|
||[sécurité](/javascript/api/word/word.documentproperties#security)|Obtient les paramètres de sécurité du document.|
||[template](/javascript/api/word/word.documentproperties#template)|Obtient le modèle du document.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Obtient ou définit le sujet du document.|
||[title](/javascript/api/word/word.documentproperties#title)|Obtient ou définit le titre du document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getNext__)|Obtient l’image insérée suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getNextOrNullObject__)|Obtient l’image insérée suivante.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getRange_rangeLocation_)|Obtient l’image, ou le point de départ ou de fin de l’image, sous la forme d’une plage.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentContentControlOrNullObject)|Obtient le contrôle de contenu qui contient l’image incluse.|
||[parentTable](/javascript/api/word/word.inlinepicture#parentTable)|Obtient le tableau qui contient l’image insérée.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parentTableCell)|Obtient la cellule de tableau qui contient l’image insérée.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parentTableCellOrNullObject)|Obtient la cellule de tableau qui contient l’image insérée.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parentTableOrNullObject)|Obtient le tableau qui contient l’image insérée.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getFirst__)|Obtient la première image insérée de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getFirstOrNullObject__)|Obtient la première image insérée de cette collection.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getLevelParagraphs_level_)|Obtient les paragraphes qui figurent au niveau spécifié de la liste.|
||[getLevelString(level: number)](/javascript/api/word/word.list#getLevelString_level_)|Obtient la puce, le numéro ou l’image au niveau spécifié sous la mesure d’une chaîne.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertParagraph_paragraphText__insertLocation_)|Insère un paragraphe à l’emplacement spécifié.|
||[id](/javascript/api/word/word.list#id)|Obtient l’ID de la liste.|
||[levelExistences](/javascript/api/word/word.list#levelExistences)|Vérifie si chacun des 9 niveaux existe dans la liste.|
||[levelTypes](/javascript/api/word/word.list#levelTypes)|Obtient les 9 types de niveau de la liste.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Obtient les paragraphes de la liste.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#setLevelAlignment_level__alignment_)|Définit l’alignement de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setLevelBullet_level__listBullet__charCode__fontName_)|Définit le format de puce au niveau spécifié de la liste.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setLevelIndents_level__textIndent__bulletNumberPictureIndent_)|Définit les deux retraits du niveau spécifié de la liste.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#setLevelNumbering_level__listNumbering__formatString_)|Définit le format de numérotation du niveau spécifié de la liste.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#setLevelStartingNumber_level__startingNumber_)|Définit le numéro de départ du niveau spécifié de la liste.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getById_id_)|Obtient une liste par son identificateur.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#getByIdOrNullObject_id_)|Obtient une liste par son identificateur.|
||[getFirst()](/javascript/api/word/word.listcollection#getFirst__)|Obtient la première liste de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getFirstOrNullObject__)|Obtient la première liste de cette collection.|
||[getItem(index : numérique)](/javascript/api/word/word.listcollection#getItem_index_)|Obtient un objet de liste en fonction de son indice dans la collection.|
||[items](/javascript/api/word/word.listcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getAncestor_parentOnly_)|Obtient le parent de l’élément de liste ou son ancêtre le plus proche si le parent n’existe pas.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getAncestorOrNullObject_parentOnly_)|Obtient le parent de l’élément de liste ou son ancêtre le plus proche si le parent n’existe pas.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getDescendants_directChildrenOnly_)|Obtient tous les éléments de liste descendants de l’élément de liste.|
||[level](/javascript/api/word/word.listitem#level)|Obtient ou définit le niveau de l’élément dans la liste.|
||[listString](/javascript/api/word/word.listitem#listString)|Obtient la puce, le numéro ou l’image de l’élément de liste sous la mesure d’une chaîne.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingIndex)|Obtient le numéro d’ordre de l’élément de liste relativement à ses frères.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachToList_listId__level_)|Permet au paragraphe de rejoindre une liste existante au niveau spécifié.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachFromList__)|Déplace ce paragraphe en dehors de la liste, si le paragraphe est un élément de liste.|
||[getNext()](/javascript/api/word/word.paragraph#getNext__)|Obtient le paragraphe suivant.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getNextOrNullObject__)|Obtient le paragraphe suivant.|
||[getPrevious()](/javascript/api/word/word.paragraph#getPrevious__)|Obtient le paragraphe précédent.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getPreviousOrNullObject__)|Obtient le paragraphe précédent.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getRange_rangeLocation_)|Obtient le paragraphe entier, ou le point de début ou de fin du paragraphe, sous la forme d’une plage.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#getTextRanges_endingMarks__trimSpacing_)|Obtient les plages de texte du paragraphe à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#insertTable_rowCount__columnCount__insertLocation__values_)|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[isLastParagraph](/javascript/api/word/word.paragraph#isLastParagraph)|Indique que le paragraphe est le dernier au sein de son corps parent.|
||[isListItem](/javascript/api/word/word.paragraph#isListItem)|Vérifie si le paragraphe est un élément de liste.|
||[list](/javascript/api/word/word.paragraph#list)|Obtient la liste à laquelle ce paragraphe appartient.|
||[listItem](/javascript/api/word/word.paragraph#listItem)|Obtient l’élément de liste du paragraphe.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listItemOrNullObject)|Obtient l’élément de liste du paragraphe.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listOrNullObject)|Obtient la liste à laquelle ce paragraphe appartient.|
||[parentBody](/javascript/api/word/word.paragraph#parentBody)|Obtient le corps parent du paragraphe.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentContentControlOrNullObject)|Obtient le contrôle de contenu qui contient le paragraphe.|
||[parentTable](/javascript/api/word/word.paragraph#parentTable)|Obtient le tableau qui contient le paragraphe.|
||[parentTableCell](/javascript/api/word/word.paragraph#parentTableCell)|Obtient la cellule de tableau qui contient le paragraphe.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parentTableCellOrNullObject)|Obtient la cellule de tableau qui contient le paragraphe.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parentTableOrNullObject)|Obtient le tableau qui contient le paragraphe.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tableNestingLevel)|Obtient le niveau de tableau du paragraphe.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split_delimiters__trimDelimiters__trimSpacing_)|Divise le paragraphe en plages enfants à l’aide de délimiteurs.|
||[startNewList()](/javascript/api/word/word.paragraph#startNewList__)|Démarre une nouvelle liste avec ce paragraphe.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#styleBuiltIn)|Obtient ou définit le nom du style prédéfini du paragraphe.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getFirst__)|Obtient le premier paragraphe de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getFirstOrNullObject__)|Obtient le premier paragraphe de cette collection.|
||[getLast()](/javascript/api/word/word.paragraphcollection#getLast__)|Obtient le dernier paragraphe dans cette collection.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getLastOrNullObject__)|Obtient le dernier paragraphe dans cette collection.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#compareLocationWith_range_)|Compare l’emplacement de la plage à celui d’une autre plage.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#expandTo_range_)|Renvoie une nouvelle plage qui s’étend dans les deux directions à partir de cette plage pour couvrir une autre plage.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#expandToOrNullObject_range_)|Renvoie une nouvelle plage qui s’étend dans les deux directions à partir de cette plage pour couvrir une autre plage.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#getHyperlinkRanges__)|Obtient les plages enfants d’un lien hypertexte au sein de la plage.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getNextTextRange_endingMarks__trimSpacing_)|Obtient la plage de texte suivante à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getNextTextRangeOrNullObject_endingMarks__trimSpacing_)|Obtient la plage de texte suivante à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getRange_rangeLocation_)|Clone la plage ou obtient le point de début ou de fin de la plage sous la forme d’une nouvelle plage.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getTextRanges_endingMarks__trimSpacing_)|Obtient les plages enfants de texte de la plage à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[lien hypertexte](/javascript/api/word/word.range#hyperlink)|Obtient le premier lien hypertexte de la plage ou définit un lien hypertexte sur la plage.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#insertTable_rowCount__columnCount__insertLocation__values_)|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#intersectWith_range_)|Retourne une nouvelle plage en tant qu’intersection de cette plage avec une autre.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#intersectWithOrNullObject_range_)|Retourne une nouvelle plage en tant qu’intersection de cette plage avec une autre.|
||[isEmpty](/javascript/api/word/word.range#isEmpty)|Vérifie si la longueur de la plage est zéro.|
||[lists](/javascript/api/word/word.range#lists)|Obtient la collection d’objets de liste figurant dans la plage.|
||[parentBody](/javascript/api/word/word.range#parentBody)|Obtient le corps parent de la plage.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentContentControlOrNullObject)|Obtient le contrôle de contenu qui contient la plage.|
||[parentTable](/javascript/api/word/word.range#parentTable)|Obtient le tableau qui contient la plage.|
||[parentTableCell](/javascript/api/word/word.range#parentTableCell)|Obtient la cellule de tableau qui contient la plage.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parentTableCellOrNullObject)|Obtient la cellule de tableau qui contient la plage.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parentTableOrNullObject)|Obtient le tableau qui contient la plage.|
||[tables](/javascript/api/word/word.range#tables)|Obtient la collection d’objets de table dans la plage.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|Divise la plage en plages enfants à l’aide de délimiteurs.|
||[styleBuiltIn](/javascript/api/word/word.range#styleBuiltIn)|Obtient ou définit le nom du style prédéfini de la plage.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getFirst__)|Obtient la première plage de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getFirstOrNullObject__)|Obtient la première plage de cette collection.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Ensemble d’api : WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getNext__)|Obtient la section suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getNextOrNullObject__)|Obtient la section suivante.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getFirst__)|Obtient la première section de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getFirstOrNullObject__)|Obtient la première section de cette collection.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addColumns_insertLocation__columnCount__values_)|Ajoute des colonnes au début ou à la fin du tableau, en utilisant la première ou la dernière colonne existante en tant que modèle.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addRows_insertLocation__rowCount__values_)|Ajoute des lignes au début ou à la fin du tableau, en utilisant la première ou la dernière ligne existante en tant que modèle.|
||[alignement](/javascript/api/word/word.table#alignment)|Obtient ou définit l’alignement du tableau par rapport à la colonne de page.|
||[autoFitWindow()](/javascript/api/word/word.table#autoFitWindow__)|Ajuste automatiquement les colonnes du tableau à la largeur de la fenêtre.|
||[clear()](/javascript/api/word/word.table#clear__)|Efface le contenu du tableau.|
||[delete()](/javascript/api/word/word.table#delete__)|Supprime le tableau entier.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deleteColumns_columnIndex__columnCount_)|Supprime des colonnes spécifiques.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleteRows_rowIndex__rowCount_)|Supprime des lignes spécifiques.|
||[distributeColumns()](/javascript/api/word/word.table#distributeColumns__)|Répartit uniformément les largeurs de colonne.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getBorder_borderLocation_)|Obtient le style de la bordure spécifiée.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getCell_rowIndex__cellIndex_)|Obtient la cellule du tableau à une ligne et une colonne spécifiées.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getCellOrNullObject_rowIndex__cellIndex_)|Obtient la cellule du tableau à une ligne et une colonne spécifiées.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getCellPadding_cellPaddingLocation_)|Obtient la marge intérieure des cellules en points.|
||[getNext()](/javascript/api/word/word.table#getNext__)|Obtient le tableau suivant.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getNextOrNullObject__)|Obtient le tableau suivant.|
||[getParagraphAfter()](/javascript/api/word/word.table#getParagraphAfter__)|Obtient le paragraphe après le tableau.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getParagraphAfterOrNullObject__)|Obtient le paragraphe après le tableau.|
||[getParagraphBefore()](/javascript/api/word/word.table#getParagraphBefore__)|Obtient le paragraphe avant le tableau.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getParagraphBeforeOrNullObject__)|Obtient le paragraphe avant le tableau.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getRange_rangeLocation_)|Obtient la plage qui contient ce tableau, ou la plage située au début ou à la fin du tableau.|
||[headerRowCount](/javascript/api/word/word.table#headerRowCount)|Obtient et définit le nombre de lignes d’en-tête.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalAlignment)|Obtient et définit l’alignement horizontal de chaque cellule du tableau.|
||[ignorePunct](/javascript/api/word/word.table#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignoreSpace)||
||[insertContentControl()](/javascript/api/word/word.table#insertContentControl__)|Insère un contrôle de contenu dans le tableau.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertParagraph_paragraphText__insertLocation_)|Insère un paragraphe à l’emplacement spécifié.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#insertTable_rowCount__columnCount__insertLocation__values_)|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[matchCase](/javascript/api/word/word.table#matchCase)||
||[matchPrefix](/javascript/api/word/word.table#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.table#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.table#matchWildcards)||
||[police](/javascript/api/word/word.table#font)|Obtient la police.|
||[isUniform](/javascript/api/word/word.table#isUniform)|Indique si toutes les lignes du tableau sont uniformes.|
||[nestingLevel](/javascript/api/word/word.table#nestingLevel)|Obtient le niveau d’imbrication du tableau.|
||[parentBody](/javascript/api/word/word.table#parentBody)|Obtient le corps parent du tableau.|
||[parentContentControl](/javascript/api/word/word.table#parentContentControl)|Obtient le contrôle de contenu qui contient le tableau.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentContentControlOrNullObject)|Obtient le contrôle de contenu qui contient le tableau.|
||[parentTable](/javascript/api/word/word.table#parentTable)|Obtient le tableau qui contient ce tableau.|
||[parentTableCell](/javascript/api/word/word.table#parentTableCell)|Obtient la cellule de tableau qui contient ce tableau.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parentTableCellOrNullObject)|Obtient la cellule de tableau qui contient ce tableau.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parentTableOrNullObject)|Obtient le tableau qui contient ce tableau.|
||[rowCount](/javascript/api/word/word.table#rowCount)|Obtient le nombre de lignes dans le tableau.|
||[rows](/javascript/api/word/word.table#rows)|Obtient toutes les lignes du tableau.|
||[tables](/javascript/api/word/word.table#tables)|Obtient les tableaux enfants imbriqués au niveau de profondeur suivant.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de l’objet table.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select_selectionMode_)|Sélectionne le tableau ou la position de début ou de fin du tableau et y accède dans l’interface utilisateur de Word.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setCellPadding_cellPaddingLocation__cellPadding_)|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.table#shadingColor)|Obtient et définit la couleur d’ombrage.|
||[style](/javascript/api/word/word.table#style)|Obtient ou définit le nom de style du tableau.|
||[styleBandedColumns](/javascript/api/word/word.table#styleBandedColumns)|Obtient et définit l’information qui indique que le tableau comporte des colonnes à bandes.|
||[styleBandedRows](/javascript/api/word/word.table#styleBandedRows)|Obtient et définit l’information qui indique que le tableau comporte des lignes à bandes.|
||[styleBuiltIn](/javascript/api/word/word.table#styleBuiltIn)|Obtient ou définit le nom du style prédéfini du tableau.|
||[styleFirstColumn](/javascript/api/word/word.table#styleFirstColumn)|Obtient et définit l’information qui indique si le tableau comporte une première colonne avec un style spécial.|
||[styleLastColumn](/javascript/api/word/word.table#styleLastColumn)|Obtient et définit l’information qui indique si le tableau comporte une dernière colonne avec un style spécial.|
||[styleTotalRow](/javascript/api/word/word.table#styleTotalRow)|Obtient et définit l’information qui indique si le tableau comporte une ligne de total (dernière ligne) avec un style spécial.|
||[values](/javascript/api/word/word.table#values)|Obtient et définit les valeurs de texte du tableau, sous la forme d’un tableau Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.table#verticalAlignment)|Obtient et définit l’alignement vertical de chaque cellule du tableau.|
||[width](/javascript/api/word/word.table#width)|Obtient et définit la largeur du tableau en points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Obtient ou définit la couleur de bordure du tableau.|
||[type](/javascript/api/word/word.tableborder#type)|Obtient ou définit le type de bordure du tableau.|
||[width](/javascript/api/word/word.tableborder#width)|Obtient ou définit la largeur, en points, de la bordure du tableau.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnWidth)|Obtient et définit la largeur de colonne de la cellule en points.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deleteColumn__)|Supprime la colonne qui contient cette cellule.|
||[deleteRow()](/javascript/api/word/word.tablecell#deleteRow__)|Supprime la ligne qui contient cette cellule.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getBorder_borderLocation_)|Obtient le style de la bordure spécifiée.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getCellPadding_cellPaddingLocation_)|Obtient la marge intérieure des cellules en points.|
||[getNext()](/javascript/api/word/word.tablecell#getNext__)|Obtient la cellule suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getNextOrNullObject__)|Obtient la cellule suivante.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalAlignment)|Obtient et définit l’alignement horizontal de la cellule.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertColumns_insertLocation__columnCount__values_)|Ajoute des colonnes à gauche ou à droite de la cellule, en utilisant la colonne de la cellule en tant que modèle.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertRows_insertLocation__rowCount__values_)|Insère les lignes au-dessus ou au-dessous de la cellule, en utilisant la ligne de la cellule en tant que modèle.|
||[body](/javascript/api/word/word.tablecell#body)|Renvoie l’objet corps de la cellule.|
||[cellIndex](/javascript/api/word/word.tablecell#cellIndex)|Obtient l’index de la cellule dans la ligne correspondante.|
||[parentRow](/javascript/api/word/word.tablecell#parentRow)|Obtient la ligne parent de la cellule.|
||[parentTable](/javascript/api/word/word.tablecell#parentTable)|Obtient le tableau parent de la cellule.|
||[rowIndex](/javascript/api/word/word.tablecell#rowIndex)|Obtient l’index de la ligne de la cellule dans le tableau.|
||[width](/javascript/api/word/word.tablecell#width)|Obtient la largeur de la cellule en points.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setCellPadding_cellPaddingLocation__cellPadding_)|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingColor)|Obtient ou définit la couleur d’ombrage de la cellule.|
||[value](/javascript/api/word/word.tablecell#value)|Obtient et définit le texte de la cellule.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalAlignment)|Obtient et définit l’alignement vertical de la cellule.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getFirst__)|Obtient la première cellule de tableau de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getFirstOrNullObject__)|Obtient la première cellule de tableau de cette collection.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getFirst__)|Obtient le premier tableau de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getFirstOrNullObject__)|Obtient le premier tableau de cette collection.|
||[items](/javascript/api/word/word.tablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear__)|Efface le contenu de la ligne.|
||[delete()](/javascript/api/word/word.tablerow#delete__)|Supprime la ligne entière.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getBorder_borderLocation_)|Obtient le style de bordure des cellules de la ligne.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getCellPadding_cellPaddingLocation_)|Obtient la marge intérieure des cellules en points.|
||[getNext()](/javascript/api/word/word.tablerow#getNext__)|Obtient la ligne suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getNextOrNullObject__)|Obtient la ligne suivante.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalAlignment)|Obtient et définit l’alignement horizontal de chaque cellule de la ligne.|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignoreSpace)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#insertRows_insertLocation__rowCount__values_)|Insère des lignes en utilisant cette ligne en tant que modèle.|
||[matchCase](/javascript/api/word/word.tablerow#matchCase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchWildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredHeight)|Obtient et définit la hauteur de ligne préférée en points.|
||[cellCount](/javascript/api/word/word.tablerow#cellCount)|Obtient le nombre de cellules dans la ligne.|
||[cells](/javascript/api/word/word.tablerow#cells)|Obtient les cellules.|
||[police](/javascript/api/word/word.tablerow#font)|Obtient la police.|
||[isHeader](/javascript/api/word/word.tablerow#isHeader)|Vérifie si la ligne est une ligne d’en-tête.|
||[parentTable](/javascript/api/word/word.tablerow#parentTable)|Obtient la table parente.|
||[rowIndex](/javascript/api/word/word.tablerow#rowIndex)|Obtient l’index de la ligne dans le tableau parent correspondant.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de la ligne.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select_selectionMode_)|Sélectionne la ligne et y accède via l’interface utilisateur de Word.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setCellPadding_cellPaddingLocation__cellPadding_)|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingColor)|Obtient et définit la couleur d’ombrage.|
||[values](/javascript/api/word/word.tablerow#values)|Obtient et définit les valeurs de texte de la ligne, sous la forme d’un tableau JavaScript 2D.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalAlignment)|Obtient et définit l’alignement vertical des cellules de la ligne.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getFirst__)|Obtient la première ligne de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getFirstOrNullObject__)|Obtient la première ligne de cette collection.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
