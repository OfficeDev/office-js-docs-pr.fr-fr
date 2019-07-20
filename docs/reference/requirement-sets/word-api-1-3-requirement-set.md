---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1,3
description: Détails sur l’ensemble de conditions requises WordApi 1,3
ms.date: 07/17/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: ca18822a60a384f15149531a59245a7b57ea39c3
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/19/2019
ms.locfileid: "35805297"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Word

WordApi 1,3 Ajout de la prise en charge supplémentaire des contrôles de contenu, du code XML personnalisé et des paramètres au niveau du document.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API ajoutées dans le cadre de l’ensemble de conditions requises WordApi 1,3.

| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument (base64File?: chaîne)](/javascript/api/word/word.application#createdocument-base64file-)|Crée un nouveau document à l’aide d’un fichier. docx codé en base64 facultatif.|
|[Body](/javascript/api/word/word.body)|[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Obtient la totalité du corps, ou le point de début ou de fin du corps, sous la forme d’une plage.|
||[insertTable (rowCount: nombre, columnCount: nombre, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes. La valeur insertLocation peut être « Start » (début) ou « End » (fin).|
||[list](/javascript/api/word/word.body#lists)|Obtient la collection d’objets list dans le corps. En lecture seule.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Obtient le corps parent du corps. Par exemple, le corps parent du corps d’une cellule de tableau peut être un en-tête. Renvoie s’il n’existe pas de corps parent. En lecture seule.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Obtient le corps parent du corps. Par exemple, le corps parent du corps d’une cellule de tableau peut être un en-tête. Renvoie un objet null s’il n’existe pas de corps parent. En lecture seule.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient le corps. Renvoie un objet null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Obtient la section parent du corps. Lève une exception s’il n’existe pas de section parent. En lecture seule.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Obtient la section parent du corps. Renvoie un objet null s’il n’y a pas de section parent. En lecture seule.|
||[emplois](/javascript/api/word/word.body#tables)|Obtient la collection d’objets table dans le corps. En lecture seule.|
||[type](/javascript/api/word/word.body#type)|Obtient le type du corps. Le type peut être « MainDoc », « Section », « Header », « Footer » ou « TableCell ». En lecture seule.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Obtient ou définit le nom de style intégré pour le corps. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Obtient le contrôle de contenu entier, ou le point de début ou de fin du contrôle de contenu, sous la forme d’une plage.|
||[getTextRanges (endingMarks: chaîne [], trimSpacing?: booléen)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Obtient les plages de texte dans le contrôle de contenu à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[insertTable (rowCount: nombre, columnCount: nombre, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes dans un contrôle de contenu ou à proximité de celui-ci. La valeur insertLocation peut être «Start», «end», «Before» ou «after».|
||[list](/javascript/api/word/word.contentcontrol#lists)|Obtient la collection d’objets list du contrôle de contenu. En lecture seule.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Obtient le corps parent du contrôle de contenu. En lecture seule.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient le contrôle de contenu spécifié. Renvoie un objet null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Obtient le tableau qui contient le contrôle de contenu. S’il n’est pas contenu dans un tableau. En lecture seule.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Obtient la cellule de tableau qui contient le contrôle de contenu. S’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Obtient la cellule de tableau qui contient le contrôle de contenu. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Obtient le tableau qui contient le contrôle de contenu. Renvoie un objet null s’il n’est pas contenu dans un tableau. En lecture seule.|
||[sous-type](/javascript/api/word/word.contentcontrol#subtype)|Obtient le sous-type du contrôle de contenu. Le sous-type peut être « RichTextInline », « RichTextParagraphs », « RichTextTableCell », « RichTextTableRow » et « RichTextTable » pour les contrôles de contenu en texte enrichi. En lecture seule.|
||[emplois](/javascript/api/word/word.contentcontrol#tables)|Obtient la collection d’objets table du contrôle de contenu. En lecture seule.|
||[Split (Delimiters: chaîne [], multiparagraphs?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Fractionne le contrôle de contenu en plages enfants à l’aide de délimiteurs.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Obtient ou définit le nom de style intégré pour le contrôle de contenu. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (ID: nombre)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Obtient un contrôle de contenu par son identificateur. Renvoie un objet null s’il n’existe pas de contrôle de contenu portant l’identificateur dans cette collection.|
||[getByTypes (types: Word. ContentControlType [])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Obtient les contrôles de contenu qui ont les types et/ou sous-types spécifiés.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Obtient le premier contrôle de contenu de cette collection. Lève une exception si cette collection est vide.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Obtient le premier contrôle de contenu de cette collection. Renvoie un objet null si cette collection est vide.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/word/word.customproperty#key)|Obtient la clé de la propriété personnalisée. En lecture seule.|
||[type](/javascript/api/word/word.customproperty#type)|Obtient le type de valeur de la propriété personnalisée. Les valeurs possibles sont les suivantes: String, Number, date, Boolean. En lecture seule.|
||[value](/javascript/api/word/word.customproperty#value)|Obtient ou définit la valeur de la propriété personnalisée. Notez que même si Word sur le Web et le format de fichier docx permettent à ces propriétés d’être arbitrairement longues, la version de bureau de Word tronque les valeurs de chaîne en caractères 255 16 bits (en créant éventuellement un Unicode non valide en fractionnant une paire de substitution).|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[Add (Key: chaîne, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Crée une nouvelle propriété personnalisée ou en définit une existante.|
||[deleteAll ()](/javascript/api/word/word.custompropertycollection#deleteall--)|Supprime toutes les propriétés personnalisées de cette collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Obtient le nombre des propriétés personnalisées.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Lève une exception si la propriété personnalisée n’existe pas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Renvoie un objet null si la propriété personnalisée n’existe pas.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Obtient les propriétés du document. En lecture seule.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[Open ()](/javascript/api/word/word.documentcreated#open--)|Ouvre le document.|
||[body](/javascript/api/word/word.documentcreated#body)|Obtient l’objet de corps du document. Le corps du document correspond à l’ensemble du texte, à l’exception des en-têtes, des pieds de page, des notes de bas de page, des zones de texte, etc. En lecture seule.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Obtient la collection d’objets de contrôle de contenu dans le document. Cela inclut les contrôles de contenu dans le corps du document, les en-têtes, les pieds de page, les zones de texte, etc.. En lecture seule.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Obtient les propriétés du document. En lecture seule.|
||[conservé](/javascript/api/word/word.documentcreated#saved)|Indique si les modifications apportées au document ont été enregistrées. La valeur true indique que le document n’a pas été modifié depuis son enregistrement. En lecture seule.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Obtient la collection d’objets section dans le document. En lecture seule.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Enregistre le document. Cette option utilise la convention de dénomination des fichiers par défaut de Word si le document n’a jamais été enregistré précédemment.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[créés](/javascript/api/word/word.documentproperties#author)|Obtient ou définit l’auteur du document.|
||[catégories](/javascript/api/word/word.documentproperties#category)|Obtient ou définit la catégorie du document.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Obtient ou définit les commentaires du document.|
||[company](/javascript/api/word/word.documentproperties#company)|Obtient ou définit la société du document.|
||[format](/javascript/api/word/word.documentproperties#format)|Obtient ou définit le format du document.|
||[Mots clés](/javascript/api/word/word.documentproperties#keywords)|Obtient ou définit les mots clés du document.|
||[dirigeant](/javascript/api/word/word.documentproperties#manager)|Obtient ou définit le responsable du document.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Obtient le nom d’application du document. En lecture seule.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Obtient la date de création du document. En lecture seule.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Obtient la collection de propriétés personnalisées du document. En lecture seule.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Obtient le dernier auteur du document. En lecture seule.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Obtient la dernière date d’impression du document. En lecture seule.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Obtient la dernière heure d’enregistrement du document. En lecture seule.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Obtient le numéro de révision du document. En lecture seule.|
||[caution](/javascript/api/word/word.documentproperties#security)|Obtient la sécurité du document. En lecture seule.|
||[template](/javascript/api/word/word.documentproperties#template)|Obtient le modèle du document. En lecture seule.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Obtient ou définit le sujet du document.|
||[title](/javascript/api/word/word.documentproperties#title)|Obtient ou définit le titre du document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext ()](/javascript/api/word/word.inlinepicture#getnext--)|Obtient l’image insérée suivante. Lève une exception si cette image incluse est la dernière.|
||[getNextOrNullObject ()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Obtient l’image insérée suivante. Renvoie un objet null si cette image incluse est la dernière.|
||[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Obtient l’image, ou le point de départ ou de fin de l’image, sous la forme d’une plage.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient l’image incluse. Renvoie un objet null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Obtient le tableau qui contient l’image insérée. S’il n’est pas contenu dans un tableau. En lecture seule.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Obtient la cellule de tableau qui contient l’image insérée. S’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Obtient la cellule de tableau qui contient l’image insérée. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Obtient le tableau qui contient l’image insérée. Renvoie un objet null s’il n’est pas contenu dans un tableau. En lecture seule.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Obtient la première image insérée de cette collection. Lève une exception si cette collection est vide.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Obtient la première image insérée de cette collection. Renvoie un objet null si cette collection est vide.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs (Level: nombre)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Obtient les paragraphes qui figurent au niveau spécifié de la liste.|
||[getLevelString (Level: nombre)](/javascript/api/word/word.list#getlevelstring-level-)|Obtient la puce, le nombre ou l’image au niveau spécifié, sous la forme d’une chaîne.|
||[insertParagraph (paragraphText: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être «Start», «end», «Before» ou «after».|
||[id](/javascript/api/word/word.list#id)|Obtient l’ID de la liste.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Vérifie si chacun des 9 niveaux existe dans la liste. Une valeur True indique le niveau existe, ce qui signifie qu’il existe au moins un élément de liste à ce niveau. En lecture seule.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Obtient les 9 types de niveau de la liste. Chaque type peut être «puce», «nombre» ou «image». En lecture seule.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Obtient les paragraphes de la liste. En lecture seule.|
||[setLevelAlignment (Level: nombre, Alignment: Word. Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Définit l’alignement de la puce, du numéro ou de l’image au niveau spécifié de la liste.|
||[setLevelBullet (Level: nombre, listBullet: Word. ListBullet, charCode?: Number, fontName?: String)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Définit le format de puce au niveau spécifié de la liste. Si la puce est « Custom », la valeur charCode est requise.|
||[setLevelIndents (Level: nombre, textIndent: nombre, bulletNumberPictureIndent: nombre)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Définit les deux retraits du niveau spécifié de la liste.|
||[setLevelNumbering (Level: nombre, listNumbering: Word. ListNumbering, formatString?: Array<String \| Number>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Définit le format de numérotation du niveau spécifié de la liste.|
||[setLevelStartingNumber (Level: nombre, startingNumber: nombre)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Définit le numéro de départ du niveau spécifié de la liste. La valeur par défaut est 1.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Obtient une liste par son identificateur. Lève une exception s’il n’existe pas de liste avec l’identificateur dans cette collection.|
||[getByIdOrNullObject (ID: nombre)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Obtient une liste par son identificateur. Renvoie un objet null s’il n’existe pas de liste avec l’identificateur dans cette collection.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Obtient la première liste de cette collection. Lève une exception si cette collection est vide.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Obtient la première liste de cette collection. Renvoie un objet null si cette collection est vide.|
||[getItem(index : numérique)](/javascript/api/word/word.listcollection#getitem-index-)|Obtient un objet de liste en fonction de son indice dans la collection.|
||[items](/javascript/api/word/word.listcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor (parentOnly?: booléen)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Obtient le parent de l’élément de liste ou son ancêtre le plus proche si le parent n’existe pas. Lève une exception si l’élément de liste n’a pas d’ancêtre.|
||[getAncestorOrNullObject (parentOnly?: booléen)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Obtient le parent de l’élément de liste ou son ancêtre le plus proche si le parent n’existe pas. Renvoie un objet null si l’élément de liste n’a pas d’ancêtre.|
||[getDescendants (directChildrenOnly?: booléen)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Obtient tous les éléments de liste descendants de l’élément de liste.|
||[level](/javascript/api/word/word.listitem#level)|Obtient ou définit le niveau de l’élément dans la liste.|
||[listString](/javascript/api/word/word.listitem#liststring)|Obtient la puce, le numéro ou l’image de l’élément de liste en tant que chaîne. En lecture seule.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Obtient le numéro d’ordre de l’élément de liste relativement à ses frères. En lecture seule.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (listId: nombre, Level: nombre)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Permet au paragraphe de rejoindre une liste existante au niveau spécifié. Échoue si le paragraphe ne peut pas rejoindre la liste ou s’il est déjà un élément de liste.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Déplace ce paragraphe en dehors de la liste, si le paragraphe est un élément de liste.|
||[getNext ()](/javascript/api/word/word.paragraph#getnext--)|Obtient le paragraphe suivant. Lève une exception si le paragraphe est le dernier.|
||[getNextOrNullObject ()](/javascript/api/word/word.paragraph#getnextornullobject--)|Obtient le paragraphe suivant. Renvoie un objet null si le paragraphe est le dernier.|
||[getPrevious ()](/javascript/api/word/word.paragraph#getprevious--)|Obtient le paragraphe précédent. Lève une exception si le paragraphe est le premier.|
||[getPreviousOrNullObject ()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Obtient le paragraphe précédent. Renvoie un objet null si le paragraphe est le premier.|
||[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Obtient le paragraphe entier, ou le point de début ou de fin du paragraphe, sous la forme d’une plage.|
||[getTextRanges (endingMarks: chaîne [], trimSpacing?: booléen)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Obtient les plages de texte dans le paragraphe à l’aide de signes de ponctuation et/ou d’autres marques de fin.|
||[insertTable (rowCount: nombre, columnCount: nombre, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Indique que le paragraphe est le dernier au sein de son corps parent. En lecture seule.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Vérifie si le paragraphe est un élément de liste. En lecture seule.|
||[liste](/javascript/api/word/word.paragraph#list)|Obtient la liste à laquelle ce paragraphe appartient. Lève une exception si le paragraphe ne se trouve pas dans une liste. En lecture seule.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Obtient l’élément de liste du paragraphe. Lève une exception si le paragraphe ne fait pas partie d’une liste. En lecture seule.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|Obtient l’élément de liste du paragraphe. Renvoie un objet null si le paragraphe ne fait pas partie d’une liste. En lecture seule.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|Obtient la liste à laquelle ce paragraphe appartient. Renvoie un objet null si le paragraphe n’est pas dans une liste. En lecture seule.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Obtient le corps parent du paragraphe. En lecture seule.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient le paragraphe. Renvoie un objet null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Obtient le tableau qui contient le paragraphe. S’il n’est pas contenu dans un tableau. En lecture seule.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Obtient la cellule de tableau qui contient le paragraphe. S’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Obtient la cellule de tableau qui contient le paragraphe. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|Obtient le tableau qui contient le paragraphe. Renvoie un objet null s’il n’est pas contenu dans un tableau. En lecture seule.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Obtient le niveau de tableau du paragraphe. Renvoie 0 si le paragraphe ne figure pas dans un tableau. En lecture seule.|
||[Split (Delimiters: chaîne [], trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Divise le paragraphe en plages enfants à l’aide de délimiteurs.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Démarre une nouvelle liste avec ce paragraphe. Échoue si le paragraphe est déjà un élément de liste.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Obtient ou définit le nom du style prédéfini du paragraphe. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Obtient le premier paragraphe de cette collection. Lève une exception si la collection est vide.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Obtient le premier paragraphe de cette collection. Renvoie un objet null si la collection est vide.|
||[getLast ()](/javascript/api/word/word.paragraphcollection#getlast--)|Obtient le dernier paragraphe dans cette collection. Lève une exception si la collection est vide.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Obtient le dernier paragraphe dans cette collection. Renvoie un objet null si la collection est vide.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (Range: Word. Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Compare l’emplacement de la plage à celui d’une autre plage.|
||[expandTo (Range: Word. Range)](/javascript/api/word/word.range#expandto-range-)|Renvoie une nouvelle plage qui s’étend dans les deux directions à partir de cette plage pour couvrir une autre plage. Cette plage n’est pas modifiée. Lève une exception si les deux plages n’ont pas d’Union.|
||[expandToOrNullObject (Range: Word. Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Renvoie une nouvelle plage qui s’étend dans les deux directions à partir de cette plage pour couvrir une autre plage. Cette plage n’est pas modifiée. Renvoie un objet null si les deux plages n’ont pas d’Union.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Obtient les plages enfants d’un lien hypertexte au sein de la plage.|
||[getNextTextRange (endingMarks: chaîne [], trimSpacing?: booléen)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Obtient la plage de texte suivante à l’aide de signes de ponctuation et/ou d’autres marques de fin. Lève une exception si cette plage de texte est la dernière.|
||[getNextTextRangeOrNullObject (endingMarks: chaîne [], trimSpacing?: booléen)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Obtient la plage de texte suivante à l’aide de signes de ponctuation et/ou d’autres marques de fin. Renvoie un objet null si cette plage de texte est la dernière.|
||[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Clone la plage ou obtient le point de début ou de fin de la plage sous la forme d’une nouvelle plage.|
||[getTextRanges (endingMarks: chaîne [], trimSpacing?: booléen)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Obtient les plages d’enfants de texte de la plage en utilisant des signes de ponctuation et/ou d’autres marques de fin.|
||[lien hypertexte](/javascript/api/word/word.range#hyperlink)|Obtient le premier lien hypertexte de la plage ou définit un lien hypertexte sur la plage. Tous les liens hypertexte de la plage sont supprimés lorsque vous définissez un nouveau lien hypertexte sur celle-ci. Utilisez un «#» pour séparer la partie d’adresse du composant facultatif emplacement.|
||[insertTable (rowCount: nombre, columnCount: nombre, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
||[intersectWith (Range: Word. Range)](/javascript/api/word/word.range#intersectwith-range-)|Retourne une nouvelle plage en tant qu’intersection de cette plage avec une autre. Cette plage n’est pas modifiée. Lève une exception si les deux plages ne sont pas superposées ou adjacentes.|
||[intersectWithOrNullObject (Range: Word. Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Retourne une nouvelle plage en tant qu’intersection de cette plage avec une autre. Cette plage n’est pas modifiée. Renvoie un objet null si les deux plages ne sont pas superposées ou adjacentes.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Vérifie si la longueur de la plage est zéro. En lecture seule.|
||[list](/javascript/api/word/word.range#lists)|Obtient la collection d’objets de liste figurant dans la plage. En lecture seule.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Obtient le corps parent de la plage. En lecture seule.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient la plage. Renvoie un objet null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Obtient le tableau qui contient la plage. S’il n’est pas contenu dans un tableau. En lecture seule.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Obtient la cellule de tableau qui contient la plage. S’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Obtient la cellule de tableau qui contient la plage. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Obtient le tableau qui contient la plage. Renvoie un objet null s’il n’est pas contenu dans un tableau. En lecture seule.|
||[emplois](/javascript/api/word/word.range#tables)|Obtient la collection d’objets de table dans la plage. En lecture seule.|
||[Split (Delimiters: chaîne [], multiparagraphs?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Divise la plage en plages enfants à l’aide de délimiteurs.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Obtient ou définit le nom du style prédéfini de la plage. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Obtient la première plage de cette collection. Lève une exception si cette collection est vide.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Obtient la première plage de cette collection. Renvoie un objet null si cette collection est vide.|
|[Section](/javascript/api/word/word.section)|[getNext ()](/javascript/api/word/word.section#getnext--)|Obtient la section suivante. Lève une exception si cette section est la dernière.|
||[getNextOrNullObject ()](/javascript/api/word/word.section#getnextornullobject--)|Obtient la section suivante. Renvoie un objet null si cette section est la dernière.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Obtient la première section de cette collection. Lève une exception si cette collection est vide.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Obtient la première section de cette collection. Renvoie un objet null si cette collection est vide.|
|[Table](/javascript/api/word/word.table)|[addColumns (insertLocation: Word. InsertLocation, columnCount: nombre, Values?: chaîne [] [])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Ajoute des colonnes au début ou à la fin du tableau, en utilisant la première ou la dernière colonne existante en tant que modèle. Applicable aux tableaux uniformes. Si spécifiées, les valeurs de chaîne sont définies sur les lignes nouvellement insérées.|
||[addRows (insertLocation: Word. InsertLocation, rowCount: nombre, Values?: chaîne [] [])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Ajoute des lignes au début ou à la fin du tableau, en utilisant la première ou la dernière ligne existante en tant que modèle. Si spécifiées, les valeurs de chaîne sont définies sur les lignes nouvellement insérées.|
||[aligne](/javascript/api/word/word.table#alignment)|Obtient ou définit l’alignement du tableau par rapport à la colonne de page. La valeur peut être «Left», «centered» ou «Right».|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Ajuste automatiquement les colonnes du tableau à la largeur de la fenêtre.|
||[clear()](/javascript/api/word/word.table#clear--)|Efface le contenu du tableau.|
||[delete()](/javascript/api/word/word.table#delete--)|Supprime le tableau entier.|
||[deleteColumns (columnIndex: nombre, columnCount?: nombre)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Supprime des colonnes spécifiques. Applicable aux tableaux uniformes.|
||[deleteRows (rowIndex: nombre, rowCount?: nombre)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Supprime des lignes spécifiques.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Répartit uniformément les largeurs de colonne. Applicable aux tableaux uniformes.|
||[getBorder (borderLocation: Word. BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Obtient le style de la bordure spécifiée.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Obtient la cellule du tableau à une ligne et une colonne spécifiées. Renvoie si la cellule de tableau spécifiée n’existe pas.|
||[getCellOrNullObject (rowIndex: nombre, cellIndex: nombre)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Obtient la cellule du tableau à une ligne et une colonne spécifiées. Renvoie un objet null si la cellule de tableau spécifiée n’existe pas.|
||[getCellPadding (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Obtient la marge intérieure des cellules en points.|
||[getNext ()](/javascript/api/word/word.table#getnext--)|Obtient le tableau suivant. Lève une exception si cette table est la dernière.|
||[getNextOrNullObject ()](/javascript/api/word/word.table#getnextornullobject--)|Obtient le tableau suivant. Renvoie un objet null si le tableau est le dernier.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Obtient le paragraphe après le tableau. Lève une exception s’il n’y a pas de paragraphe après le tableau.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Obtient le paragraphe après le tableau. Renvoie un objet null s’il n’y a pas de paragraphe après le tableau.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Obtient le paragraphe avant le tableau. S’il n’y a pas de paragraphe avant le tableau.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Obtient le paragraphe avant le tableau. Renvoie un objet null s’il n’y a pas de paragraphe avant le tableau.|
||[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Obtient la plage qui contient ce tableau, ou la plage située au début ou à la fin du tableau.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Obtient et définit le nombre de lignes d’en-tête.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Obtient et définit l’alignement horizontal de chaque cellule du tableau. La valeur peut être «Left», «centered», «Right» ou «Justified».|
||[Ignorepunct,](/javascript/api/word/word.table#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Insère un contrôle de contenu dans le tableau.|
||[insertParagraph (paragraphText: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
||[insertTable (rowCount: nombre, columnCount: nombre, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Insère un tableau avec le nombre spécifié de lignes et de colonnes. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.table#matchwildcards)||
||[police](/javascript/api/word/word.table#font)|Obtient la police. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Indique si toutes les lignes du tableau sont uniformes. En lecture seule.|
||[NestingLevel,](/javascript/api/word/word.table#nestinglevel)|Obtient le niveau d’imbrication du tableau. Les tableaux de niveau supérieur ont le niveau 1. En lecture seule.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Obtient le corps parent du tableau. En lecture seule.|
||[ParentContentControl,](/javascript/api/word/word.table#parentcontentcontrol)|Obtient le contrôle de contenu qui contient le tableau. S’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Obtient le contrôle de contenu qui contient le tableau. Renvoie un objet null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Obtient le tableau qui contient ce tableau. S’il n’est pas contenu dans un tableau. En lecture seule.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Obtient la cellule de tableau qui contient ce tableau. S’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|Obtient la cellule de tableau qui contient ce tableau. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|Obtient le tableau qui contient ce tableau. Renvoie un objet null s’il n’est pas contenu dans un tableau. En lecture seule.|
||[Stopp](/javascript/api/word/word.table#rowcount)|Obtient le nombre de lignes dans le tableau. En lecture seule.|
||[rows](/javascript/api/word/word.table#rows)|Obtient toutes les lignes du tableau. En lecture seule.|
||[emplois](/javascript/api/word/word.table#tables)|Obtient les tableaux enfants imbriqués au niveau de profondeur suivant. En lecture seule.|
||[recherche (Texted’origine: chaîne, searchOptions?: Word. SearchOptions](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié sur l’étendue de l’objet table. Les résultats de la recherche sont un ensemble d’objets de plage.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Sélectionne le tableau ou la position de début ou de fin du tableau et y accède dans l’interface utilisateur de Word.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: nombre)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Obtient et définit la couleur d’ombrage. La couleur est spécifiée au format « #RRVVBB » ou par son nom de couleur.|
||[style](/javascript/api/word/word.table#style)|Obtient ou définit le nom de style du tableau. Utilisez cette propriété pour les noms des styles personnalisés et localisés. Pour utiliser les styles prédéfinis qui sont portables entre différents paramètres régionaux, voir la propriété « styleBuiltIn ».|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Obtient et définit l’information qui indique que le tableau comporte des colonnes à bandes.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Obtient et définit l’information qui indique que le tableau comporte des lignes à bandes.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Obtient ou définit le nom du style prédéfini du tableau. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Obtient et définit l’information qui indique si le tableau comporte une première colonne avec un style spécial.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Obtient et définit l’information qui indique si le tableau comporte une dernière colonne avec un style spécial.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Obtient et définit l’information qui indique si le tableau comporte une ligne de total (dernière ligne) avec un style spécial.|
||[values](/javascript/api/word/word.table#values)|Obtient et définit les valeurs de texte du tableau, sous la forme d’un tableau Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Obtient et définit l’alignement vertical de chaque cellule du tableau. La valeur peut être «Top», «Center» ou «Bottom».|
||[width](/javascript/api/word/word.table#width)|Obtient et définit la largeur du tableau en points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Obtient ou définit la couleur de la bordure du tableau.|
||[type](/javascript/api/word/word.tableborder#type)|Obtient ou définit le type de bordure du tableau.|
||[width](/javascript/api/word/word.tableborder#width)|Obtient ou définit la largeur, en points, de la bordure du tableau. Non applicable aux types de bordure de tableau qui ont une largeur fixe.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Obtient et définit la largeur de colonne de la cellule en points. Applicable aux tableaux uniformes.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|Supprime la colonne qui contient cette cellule. Applicable aux tableaux uniformes.|
||[deleteRow ()](/javascript/api/word/word.tablecell#deleterow--)|Supprime la ligne qui contient cette cellule.|
||[getBorder (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Obtient le style de la bordure spécifiée.|
||[getCellPadding (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Obtient la marge intérieure des cellules en points.|
||[getNext ()](/javascript/api/word/word.tablecell#getnext--)|Obtient la cellule suivante. Lève une exception si cette cellule est la dernière.|
||[getNextOrNullObject ()](/javascript/api/word/word.tablecell#getnextornullobject--)|Obtient la cellule suivante. Renvoie un objet null si cette cellule est la dernière.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Obtient et définit l’alignement horizontal de la cellule. La valeur peut être «Left», «centered», «Right» ou «Justified».|
||[insertColumns (insertLocation: Word. InsertLocation, columnCount: nombre, Values?: chaîne [] [])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Ajoute des colonnes à gauche ou à droite de la cellule, en utilisant la colonne de la cellule en tant que modèle. Applicable aux tableaux uniformes. Si spécifiées, les valeurs de chaîne sont définies sur les lignes nouvellement insérées.|
||[insertRows (insertLocation: Word. InsertLocation, rowCount: nombre, Values?: chaîne [] [])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Insère les lignes au-dessus ou au-dessous de la cellule, en utilisant la ligne de la cellule en tant que modèle. Si spécifiées, les valeurs de chaîne sont définies sur les lignes nouvellement insérées.|
||[body](/javascript/api/word/word.tablecell#body)|Renvoie l’objet corps de la cellule. En lecture seule.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Obtient l’index de la cellule dans la ligne correspondante. En lecture seule.|
||[parentRow,](/javascript/api/word/word.tablecell#parentrow)|Obtient la ligne parent de la cellule. En lecture seule.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Obtient le tableau parent de la cellule. En lecture seule.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Obtient l’index de la ligne de la cellule dans le tableau. En lecture seule.|
||[width](/javascript/api/word/word.tablecell#width)|Obtient la largeur de la cellule en points. En lecture seule.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: nombre)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Obtient ou définit la couleur d’ombrage de la cellule. La couleur est spécifiée au format « #RRVVBB » ou par son nom de couleur.|
||[value](/javascript/api/word/word.tablecell#value)|Obtient et définit le texte de la cellule.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Obtient et définit l’alignement vertical de la cellule. La valeur peut être «Top», «Center» ou «Bottom».|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Obtient la première cellule de tableau de cette collection. Lève une exception si cette collection est vide.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Obtient la première cellule de tableau de cette collection. Renvoie un objet null si cette collection est vide.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Obtient le premier tableau de cette collection. Lève une exception si cette collection est vide.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Obtient le premier tableau de cette collection. Renvoie un objet null si cette collection est vide.|
||[items](/javascript/api/word/word.tablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Efface le contenu de la ligne.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Supprime la ligne entière.|
||[getBorder (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Obtient le style de bordure des cellules de la ligne.|
||[getCellPadding (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Obtient la marge intérieure des cellules en points.|
||[getNext ()](/javascript/api/word/word.tablerow#getnext--)|Obtient la ligne suivante. Lève une exception si cette ligne est la dernière.|
||[getNextOrNullObject ()](/javascript/api/word/word.tablerow#getnextornullobject--)|Obtient la ligne suivante. Renvoie un objet null si cette ligne est la dernière.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Obtient et définit l’alignement horizontal de chaque cellule de la ligne. La valeur peut être «Left», «centered», «Right» ou «Justified».|
||[Ignorepunct,](/javascript/api/word/word.tablerow#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.tablerow#ignorespace)||
||[insertRows (insertLocation: Word. InsertLocation, rowCount: nombre, Values?: chaîne [] [])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Insère des lignes en utilisant cette ligne en tant que modèle. Si les valeurs sont spécifiées, insère les valeurs sur de nouvelles lignes.|
||[matchCase](/javascript/api/word/word.tablerow#matchcase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchprefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchwildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Obtient et définit la hauteur de ligne préférée en points.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Obtient le nombre de cellules dans la ligne. En lecture seule.|
||[cases](/javascript/api/word/word.tablerow#cells)|Obtient les cellules. En lecture seule.|
||[police](/javascript/api/word/word.tablerow#font)|Obtient la police. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|
||[IsHeader,](/javascript/api/word/word.tablerow#isheader)|Vérifie si la ligne est une ligne d’en-tête. En lecture seule. Pour définir le nombre de lignes d’en-tête, utilisez HeaderRowCount sur l’objet de table.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Obtient la table parente. En lecture seule.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Obtient l’index de la ligne dans le tableau parent correspondant. En lecture seule.|
||[recherche (Texted’origine: chaîne, searchOptions?: Word. SearchOptions)](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié sur l’étendue de la ligne. Les résultats de la recherche sont un ensemble d’objets de plage.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Sélectionne la ligne et y accède via l’interface utilisateur de Word.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: nombre)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Définit la marge intérieure des cellules en points.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Obtient et définit la couleur d’ombrage. La couleur est spécifiée au format « #RRVVBB » ou par son nom de couleur.|
||[values](/javascript/api/word/word.tablerow#values)|Obtient et définit les valeurs de texte de la ligne, sous la forme d’un tableau JavaScript 2D.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Obtient et définit l’alignement vertical des cellules de la ligne. La valeur peut être «Top», «Center» ou «Bottom».|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Obtient la première ligne de cette collection. Lève une exception si cette collection est vide.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Obtient la première ligne de cette collection. Renvoie un objet null si cette collection est vide.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
