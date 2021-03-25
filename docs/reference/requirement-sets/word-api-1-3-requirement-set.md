---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1.3
description: Détails sur l’ensemble de conditions requises WordApi 1.3.
ms.date: 03/09/2021
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 0291a8a96e0ae38bf9d1061a09dac3d3b9cc3ddb
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178103"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Word

WordApi 1.3 a ajouté une prise en charge supplémentaire des contrôles de contenu et des paramètres au niveau du document.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de l’ensemble de conditions requises de l’API JavaScript pour Word 1.3. Pour afficher la documentation de référence de l’API pour toutes les API pris en charge par l’ensemble de conditions requises de l’API JavaScript pour Word 1.3 ou une version antérieure, voir API Word dans l’ensemble de conditions requises [1.3](/javascript/api/word?view=word-js-1.3&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createdocument-base64file-)|Crée un document à l’aide d’un fichier .docx encodé en base 64 facultatif.|
|[Corps](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Obtient la totalité du corps, ou le point de début ou de fin du corps, sous la forme d’une plage.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[lists](/javascript/api/word/word.body#lists)|Obtient la collection d’objets list dans le corps.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Obtient le corps parent du corps.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Obtient le corps parent du corps.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient le corps.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Obtient la section parent du corps.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Obtient la section parent du corps.|
||[tables](/javascript/api/word/word.body#tables)|Obtient la collection d’objets table dans le corps.|
||[type](/javascript/api/word/word.body#type)|Obtient le type du corps.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Obtient ou définit le nom de style intégré pour le corps.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Obtient le contrôle de contenu entier, ou le point de début ou de fin du contrôle de contenu, sous la forme d’une plage.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Obtient les plages de texte dans le contrôle de contenu à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes dans un contrôle de contenu ou à proximité de celui-ci.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Obtient la collection d’objets list du contrôle de contenu.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Obtient le corps parent du contrôle de contenu.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient le contrôle de contenu spécifié.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Obtient le tableau qui contient le contrôle de contenu.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Obtient la cellule de tableau qui contient le contrôle de contenu.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Obtient la cellule de tableau qui contient le contrôle de contenu.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Obtient le tableau qui contient le contrôle de contenu.|
||[sous-type](/javascript/api/word/word.contentcontrol#subtype)|Obtient le sous-type du contrôle de contenu.|
||[tables](/javascript/api/word/word.contentcontrol#tables)|Obtient la collection d’objets table du contrôle de contenu.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Fractionne le contrôle de contenu en plages enfants à l’aide de délimiteurs.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Obtient ou définit le nom de style intégré pour le contrôle de contenu.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Obtient un contrôle de contenu par son identificateur.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Obtient les contrôles de contenu qui ont les types et/ou sous-types spécifiés.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Obtient le premier contrôle de contenu de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Obtient le premier contrôle de contenu de cette collection.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/word/word.customproperty#key)|Obtient la clé de la propriété personnalisée.|
||[type](/javascript/api/word/word.customproperty#type)|Obtient le type de valeur de la propriété personnalisée.|
||[value](/javascript/api/word/word.customproperty#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Crée une nouvelle propriété personnalisée ou en définit une existante.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteall--)|Supprime toutes les propriétés personnalisées de cette collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Obtient le nombre des propriétés personnalisées.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Obtient les propriétés du document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open--)|Ouvre le document.|
||[body](/javascript/api/word/word.documentcreated#body)|Obtient l’objet body du document.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Obtient la collection d’objets de contrôle de contenu dans le document.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Obtient les propriétés du document.|
||[saved](/javascript/api/word/word.documentcreated#saved)|Indique si les modifications apportées au document ont été enregistrées.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Obtient la collection d’objets de section dans le document.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Enregistre le document.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[author](/javascript/api/word/word.documentproperties#author)|Obtient ou définit l’auteur du document.|
||[category](/javascript/api/word/word.documentproperties#category)|Obtient ou définit la catégorie du document.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Obtient ou définit les commentaires du document.|
||[company](/javascript/api/word/word.documentproperties#company)|Obtient ou définit la société du document.|
||[format](/javascript/api/word/word.documentproperties#format)|Obtient ou définit le format du document.|
||[keywords](/javascript/api/word/word.documentproperties#keywords)|Obtient ou définit les mots clés du document.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Obtient ou définit le responsable du document.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Obtient le nom d’application du document.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Obtient la date de création du document.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Obtient la collection de propriétés personnalisées du document.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Obtient le dernier auteur du document.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Obtient la dernière date d’impression du document.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Obtient la dernière heure d’enregistrement du document.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Obtient le numéro de révision du document.|
||[sécurité](/javascript/api/word/word.documentproperties#security)|Obtient les paramètres de sécurité du document.|
||[template](/javascript/api/word/word.documentproperties#template)|Obtient le modèle du document.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Obtient ou définit le sujet du document.|
||[title](/javascript/api/word/word.documentproperties#title)|Obtient ou définit le titre du document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getnext--)|Obtient l’image insérée suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Obtient l’image insérée suivante.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Obtient l’image, ou le point de départ ou de fin de l’image, sous la forme d’une plage.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient l’image incluse.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Obtient le tableau qui contient l’image insérée.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Obtient la cellule de tableau qui contient l’image insérée.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Obtient la cellule de tableau qui contient l’image insérée.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Obtient le tableau qui contient l’image insérée.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Obtient la première image insérée de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Obtient la première image insérée de cette collection.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Obtient les paragraphes qui figurent au niveau spécifié de la liste.|
||[getLevelString(level: number)](/javascript/api/word/word.list#getlevelstring-level-)|Obtient la puce, le numéro ou l’image au niveau spécifié sous la mesure d’une chaîne.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié.|
||[id](/javascript/api/word/word.list#id)|Obtient l’ID de la liste.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Vérifie si chacun des 9 niveaux existe dans la liste.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Obtient les 9 types de niveau de la liste.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Obtient les paragraphes de la liste.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Définit l’alignement de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Définit le format de puce au niveau spécifié de la liste.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Définit les deux retraits du niveau spécifié de la liste.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Définit le format de numérotation du niveau spécifié de la liste.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Définit le numéro de départ du niveau spécifié de la liste.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Obtient une liste par son identificateur.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Obtient une liste par son identificateur.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Obtient la première liste de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Obtient la première liste de cette collection.|
||[getItem(index : numérique)](/javascript/api/word/word.listcollection#getitem-index-)|Obtient un objet de liste en fonction de son indice dans la collection.|
||[items](/javascript/api/word/word.listcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Obtient le parent de l’élément de liste ou son ancêtre le plus proche si le parent n’existe pas.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Obtient le parent de l’élément de liste ou son ancêtre le plus proche si le parent n’existe pas.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Obtient tous les éléments de liste descendants de l’élément de liste.|
||[level](/javascript/api/word/word.listitem#level)|Obtient ou définit le niveau de l’élément dans la liste.|
||[listString](/javascript/api/word/word.listitem#liststring)|Obtient la puce, le numéro ou l’image de l’élément de liste sous la mesure d’une chaîne.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Obtient le numéro d’ordre de l’élément de liste relativement à ses frères.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Permet au paragraphe de rejoindre une liste existante au niveau spécifié.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Déplace ce paragraphe en dehors de la liste, si le paragraphe est un élément de liste.|
||[getNext()](/javascript/api/word/word.paragraph#getnext--)|Obtient le paragraphe suivant.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|Obtient le paragraphe suivant.|
||[getPrevious()](/javascript/api/word/word.paragraph#getprevious--)|Obtient le paragraphe précédent.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Obtient le paragraphe précédent.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Obtient le paragraphe entier, ou le point de début ou de fin du paragraphe, sous la forme d’une plage.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Obtient les plages de texte du paragraphe à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Indique que le paragraphe est le dernier au sein de son corps parent.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Vérifie si le paragraphe est un élément de liste.|
||[list](/javascript/api/word/word.paragraph#list)|Obtient la liste à laquelle ce paragraphe appartient.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Obtient l’élément de liste du paragraphe.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|Obtient l’élément de liste du paragraphe.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|Obtient la liste à laquelle ce paragraphe appartient.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Obtient le corps parent du paragraphe.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient le paragraphe.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Obtient le tableau qui contient le paragraphe.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Obtient la cellule de tableau qui contient le paragraphe.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Obtient la cellule de tableau qui contient le paragraphe.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|Obtient le tableau qui contient le paragraphe.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Obtient le niveau de tableau du paragraphe.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Divise le paragraphe en plages enfants à l’aide de délimiteurs.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Démarre une nouvelle liste avec ce paragraphe.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Obtient ou définit le nom du style prédéfini du paragraphe.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Obtient le premier paragraphe de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Obtient le premier paragraphe de cette collection.|
||[getLast()](/javascript/api/word/word.paragraphcollection#getlast--)|Obtient le dernier paragraphe dans cette collection.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Obtient le dernier paragraphe dans cette collection.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Compare l’emplacement de la plage à celui d’une autre plage.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#expandto-range-)|Renvoie une nouvelle plage qui s’étend dans les deux directions à partir de cette plage pour couvrir une autre plage.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Renvoie une nouvelle plage qui s’étend dans les deux directions à partir de cette plage pour couvrir une autre plage.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Obtient les plages enfants d’un lien hypertexte au sein de la plage.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Obtient la plage de texte suivante à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Obtient la plage de texte suivante à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Clone la plage ou obtient le point de début ou de fin de la plage sous la forme d’une nouvelle plage.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Obtient les plages enfants de texte de la plage à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[lien hypertexte](/javascript/api/word/word.range#hyperlink)|Obtient le premier lien hypertexte de la plage ou définit un lien hypertexte sur la plage.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#intersectwith-range-)|Retourne une nouvelle plage en tant qu’intersection de cette plage avec une autre.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Retourne une nouvelle plage en tant qu’intersection de cette plage avec une autre.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Vérifie si la longueur de la plage est zéro.|
||[lists](/javascript/api/word/word.range#lists)|Obtient la collection d’objets de liste figurant dans la plage.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Obtient le corps parent de la plage.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient la plage.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Obtient le tableau qui contient la plage.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Obtient la cellule de tableau qui contient la plage.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Obtient la cellule de tableau qui contient la plage.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Obtient le tableau qui contient la plage.|
||[tables](/javascript/api/word/word.range#tables)|Obtient la collection d’objets de table dans la plage.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Divise la plage en plages enfants à l’aide de délimiteurs.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Obtient ou définit le nom du style prédéfini de la plage.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Obtient la première plage de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Obtient la première plage de cette collection.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Ensemble d’api : WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getnext--)|Obtient la section suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|Obtient la section suivante.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Obtient la première section de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Obtient la première section de cette collection.|
|[Tableau](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Ajoute des colonnes au début ou à la fin du tableau, en utilisant la première ou la dernière colonne existante en tant que modèle.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Ajoute des lignes au début ou à la fin du tableau, en utilisant la première ou la dernière ligne existante en tant que modèle.|
||[alignement](/javascript/api/word/word.table#alignment)|Obtient ou définit l’alignement du tableau par rapport à la colonne de page.|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Ajuste automatiquement les colonnes du tableau à la largeur de la fenêtre.|
||[clear()](/javascript/api/word/word.table#clear--)|Efface le contenu du tableau.|
||[delete()](/javascript/api/word/word.table#delete--)|Supprime le tableau entier.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Supprime des colonnes spécifiques.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Supprime des lignes spécifiques.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Répartit uniformément les largeurs de colonne.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Obtient le style de la bordure spécifiée.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Obtient la cellule du tableau à une ligne et une colonne spécifiées.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Obtient la cellule du tableau à une ligne et une colonne spécifiées.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Obtient la marge intérieure des cellules en points.|
||[getNext()](/javascript/api/word/word.table#getnext--)|Obtient le tableau suivant.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|Obtient le tableau suivant.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Obtient le paragraphe après le tableau.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Obtient le paragraphe après le tableau.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Obtient le paragraphe avant le tableau.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Obtient le paragraphe avant le tableau.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Obtient la plage qui contient ce tableau, ou la plage située au début ou à la fin du tableau.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Obtient et définit le nombre de lignes d’en-tête.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Obtient et définit l’alignement horizontal de chaque cellule du tableau.|
||[ignorePunct](/javascript/api/word/word.table#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Insère un contrôle de contenu dans le tableau.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.table#matchwildcards)||
||[police](/javascript/api/word/word.table#font)|Obtient la police.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Indique si toutes les lignes du tableau sont uniformes.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Obtient le niveau d’imbrication du tableau.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Obtient le corps parent du tableau.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Obtient le contrôle de contenu qui contient le tableau.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient le tableau.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Obtient le tableau qui contient ce tableau.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Obtient la cellule de tableau qui contient ce tableau.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|Obtient la cellule de tableau qui contient ce tableau.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|Obtient le tableau qui contient ce tableau.|
||[rowCount](/javascript/api/word/word.table#rowcount)|Obtient le nombre de lignes dans le tableau.|
||[rows](/javascript/api/word/word.table#rows)|Obtient toutes les lignes du tableau.|
||[tables](/javascript/api/word/word.table#tables)|Obtient les tableaux enfants imbriqués au niveau de profondeur suivant.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de l’objet table.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Sélectionne le tableau ou la position de début ou de fin du tableau et y accède dans l’interface utilisateur de Word.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Obtient et définit la couleur d’ombrage.|
||[style](/javascript/api/word/word.table#style)|Obtient ou définit le nom de style du tableau.|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Obtient et définit l’information qui indique que le tableau comporte des colonnes à bandes.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Obtient et définit l’information qui indique que le tableau comporte des lignes à bandes.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Obtient ou définit le nom du style prédéfini du tableau.|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Obtient et définit l’information qui indique si le tableau comporte une première colonne avec un style spécial.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Obtient et définit l’information qui indique si le tableau comporte une dernière colonne avec un style spécial.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Obtient et définit l’information qui indique si le tableau comporte une ligne de total (dernière ligne) avec un style spécial.|
||[values](/javascript/api/word/word.table#values)|Obtient et définit les valeurs de texte du tableau, sous la forme d’un tableau Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Obtient et définit l’alignement vertical de chaque cellule du tableau.|
||[width](/javascript/api/word/word.table#width)|Obtient et définit la largeur du tableau en points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Obtient ou définit la couleur de bordure du tableau.|
||[type](/javascript/api/word/word.tableborder#type)|Obtient ou définit le type de bordure du tableau.|
||[width](/javascript/api/word/word.tableborder#width)|Obtient ou définit la largeur, en points, de la bordure du tableau.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Obtient et définit la largeur de colonne de la cellule en points.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|Supprime la colonne qui contient cette cellule.|
||[deleteRow()](/javascript/api/word/word.tablecell#deleterow--)|Supprime la ligne qui contient cette cellule.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Obtient le style de la bordure spécifiée.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Obtient la marge intérieure des cellules en points.|
||[getNext()](/javascript/api/word/word.tablecell#getnext--)|Obtient la cellule suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)|Obtient la cellule suivante.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Obtient et définit l’alignement horizontal de la cellule.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Ajoute des colonnes à gauche ou à droite de la cellule, en utilisant la colonne de la cellule en tant que modèle.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Insère les lignes au-dessus ou au-dessous de la cellule, en utilisant la ligne de la cellule en tant que modèle.|
||[body](/javascript/api/word/word.tablecell#body)|Renvoie l’objet corps de la cellule.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Obtient l’index de la cellule dans la ligne correspondante.|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|Obtient la ligne parent de la cellule.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Obtient le tableau parent de la cellule.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Obtient l’index de la ligne de la cellule dans le tableau.|
||[width](/javascript/api/word/word.tablecell#width)|Obtient la largeur de la cellule en points.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Obtient ou définit la couleur d’ombrage de la cellule.|
||[value](/javascript/api/word/word.tablecell#value)|Obtient et définit le texte de la cellule.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Obtient et définit l’alignement vertical de la cellule.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Obtient la première cellule de tableau de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Obtient la première cellule de tableau de cette collection.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Obtient le premier tableau de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Obtient le premier tableau de cette collection.|
||[items](/javascript/api/word/word.tablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Efface le contenu de la ligne.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Supprime la ligne entière.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Obtient le style de bordure des cellules de la ligne.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Obtient la marge intérieure des cellules en points.|
||[getNext()](/javascript/api/word/word.tablerow#getnext--)|Obtient la ligne suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)|Obtient la ligne suivante.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Obtient et définit l’alignement horizontal de chaque cellule de la ligne.|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignorespace)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Insère des lignes en utilisant cette ligne en tant que modèle.|
||[matchCase](/javascript/api/word/word.tablerow#matchcase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchprefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchwildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Obtient et définit la hauteur de ligne préférée en points.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Obtient le nombre de cellules dans la ligne.|
||[cells](/javascript/api/word/word.tablerow#cells)|Obtient les cellules.|
||[police](/javascript/api/word/word.tablerow#font)|Obtient la police.|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|Vérifie si la ligne est une ligne d’en-tête.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Obtient la table parente.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Obtient l’index de la ligne dans le tableau parent correspondant.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de la ligne.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Sélectionne la ligne et y accède via l’interface utilisateur de Word.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Obtient et définit la couleur d’ombrage.|
||[values](/javascript/api/word/word.tablerow#values)|Obtient et définit les valeurs de texte de la ligne, sous la forme d’un tableau JavaScript 2D.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Obtient et définit l’alignement vertical des cellules de la ligne.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Obtient la première ligne de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Obtient la première ligne de cette collection.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
