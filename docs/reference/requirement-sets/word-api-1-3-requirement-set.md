---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1.3
description: Détails sur l’ensemble de conditions requises WordApi 1.3.
ms.date: 03/09/2021
ms.prod: word
ms.localizationpriority: medium
---

# <a name="whats-new-in-word-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Word

WordApi 1.3 a ajouté une prise en charge supplémentaire des contrôles de contenu et des paramètres au niveau du document.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de l’ensemble de conditions requises de l’API JavaScript pour Word 1.3. Pour afficher la documentation de référence de l’API pour toutes les API pris en charge par l’ensemble de conditions requises de l’API JavaScript pour Word 1.3 ou une version antérieure, voir API Word dans l’ensemble de conditions requises [1.3](/javascript/api/word?view=word-js-1.3&preserve-view=true) ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#word-word-application-createdocument-member(1))|Crée un document à l’aide d’un fichier .docx encodé en base 64 facultatif.|
|[Corps](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#word-word-body-getrange-member(1))|Obtient la totalité du corps, ou le point de début ou de fin du corps, sous la forme d’une plage.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#word-word-body-inserttable-member(1))|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[lists](/javascript/api/word/word.body#word-word-body-lists-member)|Obtient la collection d’objets list dans le corps.|
||[parentBody](/javascript/api/word/word.body#word-word-body-parentbody-member)|Obtient le corps parent du corps.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#word-word-body-parentbodyornullobject-member)|Obtient le corps parent du corps.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#word-word-body-parentcontentcontrolornullobject-member)|Obtient le contrôle de contenu qui contient le corps.|
||[parentSection](/javascript/api/word/word.body#word-word-body-parentsection-member)|Obtient la section parent du corps.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#word-word-body-parentsectionornullobject-member)|Obtient la section parent du corps.|
||[styleBuiltIn](/javascript/api/word/word.body#word-word-body-stylebuiltin-member)|Obtient ou définit le nom de style intégré pour le corps.|
||[tables](/javascript/api/word/word.body#word-word-body-tables-member)|Obtient la collection d’objets table dans le corps.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Obtient le type du corps.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getrange-member(1))|Obtient le contrôle de contenu entier, ou le point de début ou de fin du contrôle de contenu, sous la forme d’une plage.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gettextranges-member(1))|Obtient les plages de texte dans le contrôle de contenu à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttable-member(1))|Insère un tableau avec le nombre spécifié de lignes et de colonnes dans un contrôle de contenu ou à proximité de celui-ci.|
||[lists](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-lists-member)|Obtient la collection d’objets list du contrôle de contenu.|
||[parentBody](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentbody-member)|Obtient le corps parent du contrôle de contenu.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrolornullobject-member)|Obtient le contrôle de contenu qui contient le contrôle de contenu spécifié.|
||[parentTable](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttable-member)|Obtient le tableau qui contient le contrôle de contenu.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecell-member)|Obtient la cellule de tableau qui contient le contrôle de contenu.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecellornullobject-member)|Obtient la cellule de tableau qui contient le contrôle de contenu.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttableornullobject-member)|Obtient le tableau qui contient le contrôle de contenu.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-split-member(1))|Fractionne le contrôle de contenu en plages enfants à l’aide de délimiteurs.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-stylebuiltin-member)|Obtient ou définit le nom de style intégré pour le contrôle de contenu.|
||[sous-type](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-subtype-member)|Obtient le sous-type du contrôle de contenu.|
||[tables](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tables-member)|Obtient la collection d’objets table du contrôle de contenu.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyidornullobject-member(1))|Obtient un contrôle de contenu par son identificateur.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytypes-member(1))|Obtient les contrôles de contenu qui ont les types et/ou sous-types spécifiés.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirst-member(1))|Obtient le premier contrôle de contenu de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirstornullobject-member(1))|Obtient le premier contrôle de contenu de cette collection.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#word-word-customproperty-delete-member(1))|Supprime la propriété personnalisée.|
||[key](/javascript/api/word/word.customproperty#word-word-customproperty-key-member)|Obtient la clé de la propriété personnalisée.|
||[type](/javascript/api/word/word.customproperty#word-word-customproperty-type-member)|Obtient le type de valeur de la propriété personnalisée.|
||[value](/javascript/api/word/word.customproperty#word-word-customproperty-value-member)|Obtient ou définit la valeur de la propriété personnalisée.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-add-member(1))|Crée une nouvelle propriété personnalisée ou en définit une existante.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-deleteall-member(1))|Supprime toutes les propriétés personnalisées de cette collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getcount-member(1))|Obtient le nombre des propriétés personnalisées.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitem-member(1))|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitemornullobject-member(1))|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[items](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#word-word-document-properties-member)|Obtient les propriétés du document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#word-word-documentcreated-body-member)|Obtient l’objet body du document.|
||[contentControls](/javascript/api/word/word.documentcreated#word-word-documentcreated-contentcontrols-member)|Obtient la collection d’objets de contrôle de contenu dans le document.|
||[open()](/javascript/api/word/word.documentcreated#word-word-documentcreated-open-member(1))|Ouvre le document.|
||[properties](/javascript/api/word/word.documentcreated#word-word-documentcreated-properties-member)|Obtient les propriétés du document.|
||[save()](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|Enregistre le document.|
||[saved](/javascript/api/word/word.documentcreated#word-word-documentcreated-saved-member)|Indique si les modifications apportées au document ont été enregistrées.|
||[sections](/javascript/api/word/word.documentcreated#word-word-documentcreated-sections-member)|Obtient la collection d’objets de section dans le document.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[applicationName](/javascript/api/word/word.documentproperties#word-word-documentproperties-applicationname-member)|Obtient le nom d’application du document.|
||[author](/javascript/api/word/word.documentproperties#word-word-documentproperties-author-member)|Obtient ou définit l’auteur du document.|
||[category](/javascript/api/word/word.documentproperties#word-word-documentproperties-category-member)|Obtient ou définit la catégorie du document.|
||[comments](/javascript/api/word/word.documentproperties#word-word-documentproperties-comments-member)|Obtient ou définit les commentaires du document.|
||[company](/javascript/api/word/word.documentproperties#word-word-documentproperties-company-member)|Obtient ou définit la société du document.|
||[creationDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-creationdate-member)|Obtient la date de création du document.|
||[customProperties](/javascript/api/word/word.documentproperties#word-word-documentproperties-customproperties-member)|Obtient la collection de propriétés personnalisées du document.|
||[format](/javascript/api/word/word.documentproperties#word-word-documentproperties-format-member)|Obtient ou définit le format du document.|
||[keywords](/javascript/api/word/word.documentproperties#word-word-documentproperties-keywords-member)|Obtient ou définit les mots clés du document.|
||[lastAuthor](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastauthor-member)|Obtient le dernier auteur du document.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastprintdate-member)|Obtient la dernière date d’impression du document.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastsavetime-member)|Obtient la dernière heure d’enregistrement du document.|
||[manager](/javascript/api/word/word.documentproperties#word-word-documentproperties-manager-member)|Obtient ou définit le responsable du document.|
||[revisionNumber](/javascript/api/word/word.documentproperties#word-word-documentproperties-revisionnumber-member)|Obtient le numéro de révision du document.|
||[sécurité](/javascript/api/word/word.documentproperties#word-word-documentproperties-security-member)|Obtient les paramètres de sécurité du document.|
||[subject](/javascript/api/word/word.documentproperties#word-word-documentproperties-subject-member)|Obtient ou définit le sujet du document.|
||[template](/javascript/api/word/word.documentproperties#word-word-documentproperties-template-member)|Obtient le modèle du document.|
||[title](/javascript/api/word/word.documentproperties#word-word-documentproperties-title-member)|Obtient ou définit le titre du document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnext-member(1))|Obtient l’image insérée suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnextornullobject-member(1))|Obtient l’image insérée suivante.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getrange-member(1))|Obtient l’image, ou le point de départ ou de fin de l’image, sous la forme d’une plage.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrolornullobject-member)|Obtient le contrôle de contenu qui contient l’image incluse.|
||[parentTable](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttable-member)|Obtient le tableau qui contient l’image insérée.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecell-member)|Obtient la cellule de tableau qui contient l’image insérée.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecellornullobject-member)|Obtient la cellule de tableau qui contient l’image insérée.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttableornullobject-member)|Obtient le tableau qui contient l’image insérée.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirst-member(1))|Obtient la première image insérée de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirstornullobject-member(1))|Obtient la première image insérée de cette collection.|
|[Liste](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#word-word-list-getlevelparagraphs-member(1))|Obtient les paragraphes qui figurent au niveau spécifié de la liste.|
||[getLevelString(level: number)](/javascript/api/word/word.list#word-word-list-getlevelstring-member(1))|Obtient la puce, le numéro ou l’image au niveau spécifié sous la mesure d’une chaîne.|
||[id](/javascript/api/word/word.list#word-word-list-id-member)|Obtient l’ID de la liste.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#word-word-list-insertparagraph-member(1))|Insère un paragraphe à l’emplacement spécifié.|
||[levelExistences](/javascript/api/word/word.list#word-word-list-levelexistences-member)|Vérifie si chacun des 9 niveaux existe dans la liste.|
||[levelTypes](/javascript/api/word/word.list#word-word-list-leveltypes-member)|Obtient les 9 types de niveau de la liste.|
||[paragraphs](/javascript/api/word/word.list#word-word-list-paragraphs-member)|Obtient les paragraphes de la liste.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#word-word-list-setlevelalignment-member(1))|Définit l’alignement de la puce, du numéro ou de l’image au niveau spécifié dans la liste.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#word-word-list-setlevelbullet-member(1))|Définit le format de puce au niveau spécifié de la liste.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#word-word-list-setlevelindents-member(1))|Définit les deux retraits du niveau spécifié de la liste.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#word-word-list-setlevelnumbering-member(1))|Définit le format de numérotation du niveau spécifié de la liste.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#word-word-list-setlevelstartingnumber-member(1))|Définit le numéro de départ du niveau spécifié de la liste.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyid-member(1))|Obtient une liste par son identificateur.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyidornullobject-member(1))|Obtient une liste par son identificateur.|
||[getFirst()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirst-member(1))|Obtient la première liste de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirstornullobject-member(1))|Obtient la première liste de cette collection.|
||[getItem(index : numérique)](/javascript/api/word/word.listcollection#word-word-listcollection-getitem-member(1))|Obtient un objet de liste en fonction de son indice dans la collection.|
||[items](/javascript/api/word/word.listcollection#word-word-listcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestor-member(1))|Obtient le parent de l’élément de liste ou son ancêtre le plus proche si le parent n’existe pas.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestorornullobject-member(1))|Obtient le parent de l’élément de liste ou son ancêtre le plus proche si le parent n’existe pas.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getdescendants-member(1))|Obtient tous les éléments de liste descendants de l’élément de liste.|
||[level](/javascript/api/word/word.listitem#word-word-listitem-level-member)|Obtient ou définit le niveau de l’élément dans la liste.|
||[listString](/javascript/api/word/word.listitem#word-word-listitem-liststring-member)|Obtient la puce, le numéro ou l’image de l’élément de liste sous la mesure d’une chaîne.|
||[siblingIndex](/javascript/api/word/word.listitem#word-word-listitem-siblingindex-member)|Obtient le numéro d’ordre de l’élément de liste relativement à ses frères.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#word-word-paragraph-attachtolist-member(1))|Permet au paragraphe de rejoindre une liste existante au niveau spécifié.|
||[detachFromList()](/javascript/api/word/word.paragraph#word-word-paragraph-detachfromlist-member(1))|Déplace ce paragraphe en dehors de la liste, si le paragraphe est un élément de liste.|
||[getNext()](/javascript/api/word/word.paragraph#word-word-paragraph-getnext-member(1))|Obtient le paragraphe suivant.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getnextornullobject-member(1))|Obtient le paragraphe suivant.|
||[getPrevious()](/javascript/api/word/word.paragraph#word-word-paragraph-getprevious-member(1))|Obtient le paragraphe précédent.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getpreviousornullobject-member(1))|Obtient le paragraphe précédent.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-getrange-member(1))|Obtient le paragraphe entier, ou le point de début ou de fin du paragraphe, sous la forme d’une plage.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-gettextranges-member(1))|Obtient les plages de texte du paragraphe à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#word-word-paragraph-inserttable-member(1))|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[isLastParagraph](/javascript/api/word/word.paragraph#word-word-paragraph-islastparagraph-member)|Indique que le paragraphe est le dernier au sein de son corps parent.|
||[isListItem](/javascript/api/word/word.paragraph#word-word-paragraph-islistitem-member)|Vérifie si le paragraphe est un élément de liste.|
||[list](/javascript/api/word/word.paragraph#word-word-paragraph-list-member)|Obtient la liste à laquelle ce paragraphe appartient.|
||[listItem](/javascript/api/word/word.paragraph#word-word-paragraph-listitem-member)|Obtient l’élément de liste du paragraphe.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listitemornullobject-member)|Obtient l’élément de liste du paragraphe.|
||[listOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listornullobject-member)|Obtient la liste à laquelle ce paragraphe appartient.|
||[parentBody](/javascript/api/word/word.paragraph#word-word-paragraph-parentbody-member)|Obtient le corps parent du paragraphe.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrolornullobject-member)|Obtient le contrôle de contenu qui contient le paragraphe.|
||[parentTable](/javascript/api/word/word.paragraph#word-word-paragraph-parenttable-member)|Obtient le tableau qui contient le paragraphe.|
||[parentTableCell](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecell-member)|Obtient la cellule de tableau qui contient le paragraphe.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecellornullobject-member)|Obtient la cellule de tableau qui contient le paragraphe.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttableornullobject-member)|Obtient le tableau qui contient le paragraphe.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-split-member(1))|Divise le paragraphe en plages enfants à l’aide de délimiteurs.|
||[startNewList()](/javascript/api/word/word.paragraph#word-word-paragraph-startnewlist-member(1))|Démarre une nouvelle liste avec ce paragraphe.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#word-word-paragraph-stylebuiltin-member)|Obtient ou définit le nom du style prédéfini du paragraphe.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#word-word-paragraph-tablenestinglevel-member)|Obtient le niveau de tableau du paragraphe.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirst-member(1))|Obtient le premier paragraphe de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirstornullobject-member(1))|Obtient le premier paragraphe de cette collection.|
||[getLast()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlast-member(1))|Obtient le dernier paragraphe dans cette collection.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlastornullobject-member(1))|Obtient le dernier paragraphe dans cette collection.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-comparelocationwith-member(1))|Compare l’emplacement de la plage à celui d’une autre plage.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandto-member(1))|Renvoie une nouvelle plage qui s’étend dans les deux directions à partir de cette plage pour couvrir une autre plage.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandtoornullobject-member(1))|Renvoie une nouvelle plage qui s’étend dans les deux directions à partir de cette plage pour couvrir une autre plage.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#word-word-range-gethyperlinkranges-member(1))|Obtient les plages enfants d’un lien hypertexte au sein de la plage.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrange-member(1))|Obtient la plage de texte suivante à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrangeornullobject-member(1))|Obtient la plage de texte suivante à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#word-word-range-getrange-member(1))|Clone la plage ou obtient le point de début ou de fin de la plage sous la forme d’une nouvelle plage.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-gettextranges-member(1))|Obtient les plages enfants de texte de la plage à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[lien hypertexte](/javascript/api/word/word.range#word-word-range-hyperlink-member)|Obtient le premier lien hypertexte de la plage ou définit un lien hypertexte sur la plage.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#word-word-range-inserttable-member(1))|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwith-member(1))|Retourne une nouvelle plage en tant qu’intersection de cette plage avec une autre.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwithornullobject-member(1))|Retourne une nouvelle plage en tant qu’intersection de cette plage avec une autre.|
||[isEmpty](/javascript/api/word/word.range#word-word-range-isempty-member)|Vérifie si la longueur de la plage est zéro.|
||[lists](/javascript/api/word/word.range#word-word-range-lists-member)|Obtient la collection d’objets de liste figurant dans la plage.|
||[parentBody](/javascript/api/word/word.range#word-word-range-parentbody-member)|Obtient le corps parent de la plage.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#word-word-range-parentcontentcontrolornullobject-member)|Obtient le contrôle de contenu qui contient la plage.|
||[parentTable](/javascript/api/word/word.range#word-word-range-parenttable-member)|Obtient le tableau qui contient la plage.|
||[parentTableCell](/javascript/api/word/word.range#word-word-range-parenttablecell-member)|Obtient la cellule de tableau qui contient la plage.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#word-word-range-parenttablecellornullobject-member)|Obtient la cellule de tableau qui contient la plage.|
||[parentTableOrNullObject](/javascript/api/word/word.range#word-word-range-parenttableornullobject-member)|Obtient le tableau qui contient la plage.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-split-member(1))|Divise la plage en plages enfants à l’aide de délimiteurs.|
||[styleBuiltIn](/javascript/api/word/word.range#word-word-range-stylebuiltin-member)|Obtient ou définit le nom du style prédéfini de la plage.|
||[tables](/javascript/api/word/word.range#word-word-range-tables-member)|Obtient la collection d’objets de table dans la plage.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirst-member(1))|Obtient la première plage de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirstornullobject-member(1))|Obtient la première plage de cette collection.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#word-word-requestcontext-application-member)|[Ensemble d’api : WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#word-word-section-getnext-member(1))|Obtient la section suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.section#word-word-section-getnextornullobject-member(1))|Obtient la section suivante.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirst-member(1))|Obtient la première section de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirstornullobject-member(1))|Obtient la première section de cette collection.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addcolumns-member(1))|Ajoute des colonnes au début ou à la fin du tableau, en utilisant la première ou la dernière colonne existante en tant que modèle.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addrows-member(1))|Ajoute des lignes au début ou à la fin du tableau, en utilisant la première ou la dernière ligne existante en tant que modèle.|
||[alignement](/javascript/api/word/word.table#word-word-table-alignment-member)|Obtient ou définit l’alignement du tableau par rapport à la colonne de page.|
||[autoFitWindow()](/javascript/api/word/word.table#word-word-table-autofitwindow-member(1))|Ajuste automatiquement les colonnes du tableau à la largeur de la fenêtre.|
||[clear()](/javascript/api/word/word.table#word-word-table-clear-member(1))|Efface le contenu du tableau.|
||[delete()](/javascript/api/word/word.table#word-word-table-delete-member(1))|Supprime le tableau entier.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#word-word-table-deletecolumns-member(1))|Supprime des colonnes spécifiques.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#word-word-table-deleterows-member(1))|Supprime des lignes spécifiques.|
||[distributeColumns()](/javascript/api/word/word.table#word-word-table-distributecolumns-member(1))|Répartit uniformément les largeurs de colonne.|
||[police](/javascript/api/word/word.table#word-word-table-font-member)|Obtient la police.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#word-word-table-getborder-member(1))|Obtient le style de la bordure spécifiée.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcell-member(1))|Obtient la cellule du tableau à une ligne et une colonne spécifiées.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcellornullobject-member(1))|Obtient la cellule du tableau à une ligne et une colonne spécifiées.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#word-word-table-getcellpadding-member(1))|Obtient la marge intérieure des cellules en points.|
||[getNext()](/javascript/api/word/word.table#word-word-table-getnext-member(1))|Obtient le tableau suivant.|
||[getNextOrNullObject()](/javascript/api/word/word.table#word-word-table-getnextornullobject-member(1))|Obtient le tableau suivant.|
||[getParagraphAfter()](/javascript/api/word/word.table#word-word-table-getparagraphafter-member(1))|Obtient le paragraphe après le tableau.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphafterornullobject-member(1))|Obtient le paragraphe après le tableau.|
||[getParagraphBefore()](/javascript/api/word/word.table#word-word-table-getparagraphbefore-member(1))|Obtient le paragraphe avant le tableau.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphbeforeornullobject-member(1))|Obtient le paragraphe avant le tableau.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#word-word-table-getrange-member(1))|Obtient la plage qui contient ce tableau, ou la plage située au début ou à la fin du tableau.|
||[headerRowCount](/javascript/api/word/word.table#word-word-table-headerrowcount-member)|Obtient et définit le nombre de lignes d’en-tête.|
||[horizontalAlignment](/javascript/api/word/word.table#word-word-table-horizontalalignment-member)|Obtient et définit l’alignement horizontal de chaque cellule du tableau.|
||[ignorePunct](/javascript/api/word/word.table#word-word-table-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.table#word-word-table-ignorespace-member)||
||[insertContentControl()](/javascript/api/word/word.table#word-word-table-insertcontentcontrol-member(1))|Insère un contrôle de contenu dans le tableau.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#word-word-table-insertparagraph-member(1))|Insère un paragraphe à l’emplacement spécifié.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#word-word-table-inserttable-member(1))|Insère un tableau avec le nombre spécifié de lignes et de colonnes.|
||[isUniform](/javascript/api/word/word.table#word-word-table-isuniform-member)|Indique si toutes les lignes du tableau sont uniformes.|
||[matchCase](/javascript/api/word/word.table#word-word-table-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.table#word-word-table-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.table#word-word-table-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.table#word-word-table-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.table#word-word-table-matchwildcards-member)||
||[nestingLevel](/javascript/api/word/word.table#word-word-table-nestinglevel-member)|Obtient le niveau d’imbrication du tableau.|
||[parentBody](/javascript/api/word/word.table#word-word-table-parentbody-member)|Obtient le corps parent du tableau.|
||[parentContentControl](/javascript/api/word/word.table#word-word-table-parentcontentcontrol-member)|Obtient le contrôle de contenu qui contient le tableau.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#word-word-table-parentcontentcontrolornullobject-member)|Obtient le contrôle de contenu qui contient le tableau.|
||[parentTable](/javascript/api/word/word.table#word-word-table-parenttable-member)|Obtient le tableau qui contient ce tableau.|
||[parentTableCell](/javascript/api/word/word.table#word-word-table-parenttablecell-member)|Obtient la cellule de tableau qui contient ce tableau.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#word-word-table-parenttablecellornullobject-member)|Obtient la cellule de tableau qui contient ce tableau.|
||[parentTableOrNullObject](/javascript/api/word/word.table#word-word-table-parenttableornullobject-member)|Obtient le tableau qui contient ce tableau.|
||[rowCount](/javascript/api/word/word.table#word-word-table-rowcount-member)|Obtient le nombre de lignes dans le tableau.|
||[rows](/javascript/api/word/word.table#word-word-table-rows-member)|Obtient toutes les lignes du tableau.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#word-word-table-search-member(1))|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de l’objet table.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#word-word-table-select-member(1))|Sélectionne le tableau ou la position de début ou de fin du tableau et y accède dans l’interface utilisateur de Word.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#word-word-table-setcellpadding-member(1))|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.table#word-word-table-shadingcolor-member)|Obtient et définit la couleur d’ombrage.|
||[style](/javascript/api/word/word.table#word-word-table-style-member)|Obtient ou définit le nom de style du tableau.|
||[styleBandedColumns](/javascript/api/word/word.table#word-word-table-stylebandedcolumns-member)|Obtient et définit l’information qui indique que le tableau comporte des colonnes à bandes.|
||[styleBandedRows](/javascript/api/word/word.table#word-word-table-stylebandedrows-member)|Obtient et définit l’information qui indique que le tableau comporte des lignes à bandes.|
||[styleBuiltIn](/javascript/api/word/word.table#word-word-table-stylebuiltin-member)|Obtient ou définit le nom du style prédéfini du tableau.|
||[styleFirstColumn](/javascript/api/word/word.table#word-word-table-stylefirstcolumn-member)|Obtient et définit l’information qui indique si le tableau comporte une première colonne avec un style spécial.|
||[styleLastColumn](/javascript/api/word/word.table#word-word-table-stylelastcolumn-member)|Obtient et définit l’information qui indique si le tableau comporte une dernière colonne avec un style spécial.|
||[styleTotalRow](/javascript/api/word/word.table#word-word-table-styletotalrow-member)|Obtient et définit l’information qui indique si le tableau comporte une ligne de total (dernière ligne) avec un style spécial.|
||[tables](/javascript/api/word/word.table#word-word-table-tables-member)|Obtient les tableaux enfants imbriqués au niveau de profondeur suivant.|
||[values](/javascript/api/word/word.table#word-word-table-values-member)|Obtient et définit les valeurs de texte du tableau, sous la forme d’un tableau Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.table#word-word-table-verticalalignment-member)|Obtient et définit l’alignement vertical de chaque cellule du tableau.|
||[width](/javascript/api/word/word.table#word-word-table-width-member)|Obtient et définit la largeur du tableau en points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#word-word-tableborder-color-member)|Obtient ou définit la couleur de bordure du tableau.|
||[type](/javascript/api/word/word.tableborder#word-word-tableborder-type-member)|Obtient ou définit le type de bordure du tableau.|
||[width](/javascript/api/word/word.tableborder#word-word-tableborder-width-member)|Obtient ou définit la largeur, en points, de la bordure du tableau.|
|[TableCell](/javascript/api/word/word.tablecell)|[body](/javascript/api/word/word.tablecell#word-word-tablecell-body-member)|Renvoie l’objet corps de la cellule.|
||[cellIndex](/javascript/api/word/word.tablecell#word-word-tablecell-cellindex-member)|Obtient l’index de la cellule dans la ligne correspondante.|
||[columnWidth](/javascript/api/word/word.tablecell#word-word-tablecell-columnwidth-member)|Obtient et définit la largeur de colonne de la cellule en points.|
||[deleteColumn()](/javascript/api/word/word.tablecell#word-word-tablecell-deletecolumn-member(1))|Supprime la colonne qui contient cette cellule.|
||[deleteRow()](/javascript/api/word/word.tablecell#word-word-tablecell-deleterow-member(1))|Supprime la ligne qui contient cette cellule.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getborder-member(1))|Obtient le style de la bordure spécifiée.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getcellpadding-member(1))|Obtient la marge intérieure des cellules en points.|
||[getNext()](/javascript/api/word/word.tablecell#word-word-tablecell-getnext-member(1))|Obtient la cellule suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#word-word-tablecell-getnextornullobject-member(1))|Obtient la cellule suivante.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-horizontalalignment-member)|Obtient et définit l’alignement horizontal de la cellule.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertcolumns-member(1))|Ajoute des colonnes à gauche ou à droite de la cellule, en utilisant la colonne de la cellule en tant que modèle.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertrows-member(1))|Insère les lignes au-dessus ou au-dessous de la cellule, en utilisant la ligne de la cellule en tant que modèle.|
||[parentRow](/javascript/api/word/word.tablecell#word-word-tablecell-parentrow-member)|Obtient la ligne parent de la cellule.|
||[parentTable](/javascript/api/word/word.tablecell#word-word-tablecell-parenttable-member)|Obtient le tableau parent de la cellule.|
||[rowIndex](/javascript/api/word/word.tablecell#word-word-tablecell-rowindex-member)|Obtient l’index de la ligne de la cellule dans le tableau.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#word-word-tablecell-setcellpadding-member(1))|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.tablecell#word-word-tablecell-shadingcolor-member)|Obtient ou définit la couleur d’ombrage de la cellule.|
||[value](/javascript/api/word/word.tablecell#word-word-tablecell-value-member)|Obtient et définit le texte de la cellule.|
||[verticalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-verticalalignment-member)|Obtient et définit l’alignement vertical de la cellule.|
||[width](/javascript/api/word/word.tablecell#word-word-tablecell-width-member)|Obtient la largeur de la cellule en points.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirst-member(1))|Obtient la première cellule de tableau de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirstornullobject-member(1))|Obtient la première cellule de tableau de cette collection.|
||[items](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirst-member(1))|Obtient le premier tableau de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirstornullobject-member(1))|Obtient le premier tableau de cette collection.|
||[items](/javascript/api/word/word.tablecollection#word-word-tablecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableRow](/javascript/api/word/word.tablerow)|[cellCount](/javascript/api/word/word.tablerow#word-word-tablerow-cellcount-member)|Obtient le nombre de cellules dans la ligne.|
||[cells](/javascript/api/word/word.tablerow#word-word-tablerow-cells-member)|Obtient les cellules.|
||[clear()](/javascript/api/word/word.tablerow#word-word-tablerow-clear-member(1))|Efface le contenu de la ligne.|
||[delete()](/javascript/api/word/word.tablerow#word-word-tablerow-delete-member(1))|Supprime la ligne entière.|
||[police](/javascript/api/word/word.tablerow#word-word-tablerow-font-member)|Obtient la police.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getborder-member(1))|Obtient le style de bordure des cellules de la ligne.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getcellpadding-member(1))|Obtient la marge intérieure des cellules en points.|
||[getNext()](/javascript/api/word/word.tablerow#word-word-tablerow-getnext-member(1))|Obtient la ligne suivante.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#word-word-tablerow-getnextornullobject-member(1))|Obtient la ligne suivante.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-horizontalalignment-member)|Obtient et définit l’alignement horizontal de chaque cellule de la ligne.|
||[ignorePunct](/javascript/api/word/word.tablerow#word-word-tablerow-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.tablerow#word-word-tablerow-ignorespace-member)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#word-word-tablerow-insertrows-member(1))|Insère des lignes en utilisant cette ligne en tant que modèle.|
||[isHeader](/javascript/api/word/word.tablerow#word-word-tablerow-isheader-member)|Vérifie si la ligne est une ligne d’en-tête.|
||[matchCase](/javascript/api/word/word.tablerow#word-word-tablerow-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.tablerow#word-word-tablerow-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.tablerow#word-word-tablerow-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.tablerow#word-word-tablerow-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.tablerow#word-word-tablerow-matchwildcards-member)||
||[parentTable](/javascript/api/word/word.tablerow#word-word-tablerow-parenttable-member)|Obtient la table parente.|
||[preferredHeight](/javascript/api/word/word.tablerow#word-word-tablerow-preferredheight-member)|Obtient et définit la hauteur de ligne préférée en points.|
||[rowIndex](/javascript/api/word/word.tablerow#word-word-tablerow-rowindex-member)|Obtient l’index de la ligne dans le tableau parent correspondant.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1))|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de la ligne.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#word-word-tablerow-select-member(1))|Sélectionne la ligne et y accède via l’interface utilisateur de Word.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#word-word-tablerow-setcellpadding-member(1))|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.tablerow#word-word-tablerow-shadingcolor-member)|Obtient et définit la couleur d’ombrage.|
||[values](/javascript/api/word/word.tablerow#word-word-tablerow-values-member)|Obtient et définit les valeurs de texte de la ligne, sous la forme d’un tableau JavaScript 2D.|
||[verticalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-verticalalignment-member)|Obtient et définit l’alignement vertical des cellules de la ligne.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirst-member(1))|Obtient la première ligne de cette collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirstornullobject-member(1))|Obtient la première ligne de cette collection.|
||[items](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
