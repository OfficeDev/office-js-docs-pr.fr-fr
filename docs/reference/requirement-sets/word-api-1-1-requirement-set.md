---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1,1
description: Détails sur l’ensemble de conditions requises WordApi 1,1
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 371638c18cff882f2b3907f1adedb6748761cc0c
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996437"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Nouveautés de l’API JavaScript pour Word 1,1

WordApi 1,1 est le premier ensemble de conditions requises de l’API JavaScript pour Word. Il s’agit du seul ensemble de conditions requises de l’API Word pris en charge par Word 2016.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Word 1,1. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’API JavaScript pour Word, ensemble de conditions requises 1,1, voir [API Word dans l’ensemble de conditions requises 1,1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[Corps](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|Efface le contenu de l’objet de corps.|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Obtient une représentation HTML de l’objet Body.|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|Obtient la représentation OOXML (Office Open XML) de l’objet de corps.|
||[Ignorepunct,](/javascript/api/word/word.body#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.body#ignorespace)||
||[insertBreak (breakType : Word. BreakType, insertLocation : Word. InsertLocation)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|Encadre l’objet de corps avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64 (base64File : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|Insère un document dans le corps à l’emplacement spécifié.|
||[insertHtml (HTML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|Insère du code HTML à l’emplacement spécifié.|
||[insertOoxml (OOXML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|Insère du code OOXML à l’emplacement spécifié.|
||[insertParagraph (paragraphText : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText (Text : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|Insère du texte dans le corps à l’emplacement spécifié.|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[matchPrefix](/javascript/api/word/word.body#matchprefix)||
||[matchSuffix](/javascript/api/word/word.body#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|Obtient la collection d’objets de contrôle de contenu de texte enrichi dans le corps.|
||[police](/javascript/api/word/word.body#font)|Obtient le format de texte du corps.|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|Obtient la collection d’objets InlinePicture dans le corps.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Obtient la collection d’objets Paragraph dans le corps.|
||[ParentContentControl,](/javascript/api/word/word.body#parentcontentcontrol)|Obtient le contrôle de contenu qui contient le corps.|
||[text](/javascript/api/word/word.body#text)|Obtient le texte du corps.|
||[recherche (Texted’origine : chaîne, searchOptions ?: Word. SearchOptions \| {ignorepunct, ?: Boolean ignorespace, ?: Boolean MatchCase ?: Boolean matchPrefix ?: Boolean matchSuffix ?: Boolean-MatchWholeWord ?: Boolean matchWildcards ?: Boolean})](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié sur l’étendue de l’objet Body.|
||[Select (selectionMode ?: Word. SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|Sélectionne le corps et y accède via l’interface utilisateur de Word.|
||[style](/javascript/api/word/word.body#style)|Obtient ou définit le nom du style du corps.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[apparence](/javascript/api/word/word.contentcontrol#appearance)|Obtient ou définit l’apparence du contrôle de contenu.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotdelete)|Obtient ou définit une valeur qui indique si l’utilisateur peut supprimer le contrôle de contenu.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotedit)|Obtient ou définit une valeur qui indique si l’utilisateur peut modifier le contenu du contrôle.|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|Efface le contenu du contrôle de contenu.|
||[color](/javascript/api/word/word.contentcontrol#color)|Obtient ou définit la couleur du contrôle de contenu.|
||[Delete (keepContent : valeur booléenne)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|Supprime le contrôle de contenu et son contenu.|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|Obtient une représentation HTML de l’objet de contrôle de contenu.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|Obtient la représentation Office Open XML (OOXML) de l’objet de contrôle de contenu.|
||[Ignorepunct,](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak (breakType : Word. BreakType, insertLocation : Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertFileFromBase64 (base64File : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|Insère un document dans le contrôle de contenu à l’emplacement spécifié.|
||[insertHtml (HTML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|Insère du code HTML dans le contrôle de contenu, à l’emplacement spécifié.|
||[insertOoxml (OOXML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|Insère OOXML dans le contrôle de contenu à l’emplacement spécifié.|
||[insertParagraph (paragraphText : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText (Text : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|Insère du texte dans le contrôle de contenu, à l’emplacement spécifié.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchprefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[PlaceholderText,](/javascript/api/word/word.contentcontrol#placeholdertext)|Obtient ou définit le texte de l’espace réservé du contrôle de contenu.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|Obtient la collection d’objets de contrôle de contenu compris dans le contrôle de contenu.|
||[police](/javascript/api/word/word.contentcontrol#font)|Obtient le format de texte du contrôle de contenu.|
||[id](/javascript/api/word/word.contentcontrol#id)|Obtient un entier qui représente l’identificateur du contrôle de contenu.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|Obtient la collection d’objets inlinePicture du contrôle de contenu.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Obtient la collection d’objets de paragraphe du contrôle de contenu.|
||[ParentContentControl,](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|Obtient le contrôle de contenu qui contient le contrôle de contenu spécifié.|
||[text](/javascript/api/word/word.contentcontrol#text)|Obtient le texte du contrôle de contenu.|
||[type](/javascript/api/word/word.contentcontrol#type)|Obtient le type du contrôle de contenu.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removewhenedited)|Obtient ou définit une valeur qui indique si le contrôle de contenu doit être supprimé après modification.|
||[recherche (Texted’origine : chaîne, searchOptions ?: Word. SearchOptions \| {ignorepunct, ?: Boolean ignorespace, ?: Boolean MatchCase ?: Boolean matchPrefix ?: Boolean matchSuffix ?: Boolean-MatchWholeWord ?: Boolean matchWildcards ?: Boolean})](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié dans l’étendue de l’objet de contrôle de contenu.|
||[Select (selectionMode ?: Word. SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|Sélectionne le contrôle de contenu.|
||[style](/javascript/api/word/word.contentcontrol#style)|Obtient ou définit le nom du style pour le contrôle de contenu.|
||[Numéro](/javascript/api/word/word.contentcontrol#tag)|Obtient ou définit un indicateur pour identifier un contrôle de contenu.|
||[title](/javascript/api/word/word.contentcontrol#title)|Obtient ou définit le titre d’un contrôle de contenu.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|Obtient un contrôle de contenu par son identificateur.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|Obtient les contrôles de contenu qui portent l’indicateur spécifié.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|Obtient les contrôles de contenu qui ont le titre spécifié.|
||[getItem(index : numérique)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|Obtient un contrôle de contenu par son index dans la collection.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[getSelection ()](/javascript/api/word/word.document#getselection--)|Obtient la sélection actuelle du document.|
||[body](/javascript/api/word/word.document#body)|Obtient l’objet de corps du document.|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|Obtient la collection d’objets de contrôle de contenu dans le document.|
||[conservé](/javascript/api/word/word.document#saved)|Indique si les modifications apportées au document ont été enregistrées.|
||[sections](/javascript/api/word/word.document#sections)|Obtient la collection d’objets section dans le document.|
||[save()](/javascript/api/word/word.document#save--)|Enregistre le document.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Obtient ou définit une valeur qui indique si la police en gras.|
||[color](/javascript/api/word/word.font#color)|Obtient ou définit la couleur de la police spécifiée.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doublestrikethrough)|Obtient ou définit une valeur qui indique si la police a un double barré.|
||[highlightColor](/javascript/api/word/word.font#highlightcolor)|Obtient ou définit la couleur de surbrillance.|
||[italic](/javascript/api/word/word.font#italic)|Obtient ou définit une valeur qui indique si la police est en italique.|
||[name](/javascript/api/word/word.font#name)|Obtient ou définit une valeur qui représente le nom de la police.|
||[size](/javascript/api/word/word.font#size)|Obtient ou définit une valeur qui représente la taille de police en points.|
||[Doubles](/javascript/api/word/word.font#strikethrough)|Obtient ou définit une valeur qui indique si la police est barrée.|
||[Subscript](/javascript/api/word/word.font#subscript)|Obtient ou définit une valeur qui indique si la police correspond à du texte mis en indice.|
||[superscript](/javascript/api/word/word.font#superscript)|Obtient ou définit une valeur qui indique si la police correspond à du texte en exposant.|
||[underline](/javascript/api/word/word.font#underline)|Obtient ou définit une valeur qui indique le type de trait de soulignement de la police.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|Obtient ou définit une valeur de type String qui représente le texte de remplacement associé à l’image incluse.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|Obtient ou définit une chaîne qui contient le titre de l’image incluse.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|Obtient la représentation de chaîne encodée au format Base64 de l’image incluse.|
||[height](/javascript/api/word/word.inlinepicture#height)|Obtient ou définit un nombre qui décrit la hauteur de l’image incluse.|
||[lien hypertexte](/javascript/api/word/word.inlinepicture#hyperlink)|Obtient ou définit un lien hypertexte sur l’image.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|Encadre l’image incluse avec un contrôle de contenu de texte enrichi.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|Obtient ou définit une valeur qui indique si l’image incluse conserve ses proportions d’origine lorsque vous la redimensionnez.|
||[ParentContentControl,](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|Obtient le contrôle de contenu qui contient l’image incluse.|
||[width](/javascript/api/word/word.inlinepicture#width)|Obtient ou définit un nombre qui décrit la largeur de l’image incluse.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Paragraph](/javascript/api/word/word.paragraph)|[aligne](/javascript/api/word/word.paragraph#alignment)|Obtient ou définit l’alignement d’un paragraphe.|
||[clear()](/javascript/api/word/word.paragraph#clear--)|Efface le contenu de l’objet de paragraphe.|
||[delete()](/javascript/api/word/word.paragraph#delete--)|Supprime le paragraphe et son contenu du document.|
||[FirstLineIndent,](/javascript/api/word/word.paragraph#firstlineindent)|Renvoie ou définit la valeur, en points, du retrait de première ligne ou du retrait négatif.|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Obtient une représentation HTML de l’objet Paragraph.|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Obtient la représentation Office Open XML (OOXML) de l’objet de paragraphe.|
||[Ignorepunct,](/javascript/api/word/word.paragraph#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak (breakType : Word. BreakType, insertLocation : Word. InsertLocation)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|Encadre l’objet de paragraphe avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64 (base64File : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|Insère un document dans le paragraphe à l’emplacement spécifié.|
||[insertHtml (HTML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|Insère du code HTML dans le paragraphe à l’emplacement spécifié.|
||[insertInlinePictureFromBase64 (base64EncodedImage : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insère une image dans le paragraphe à l’emplacement spécifié.|
||[insertOoxml (OOXML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|Insère OOXML dans le paragraphe à l’emplacement spécifié.|
||[insertParagraph (paragraphText : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText (Text : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|Insère du texte dans le paragraphe à l’emplacement spécifié.|
||[leftIndent](/javascript/api/word/word.paragraph#leftindent)|Obtient ou définit la valeur de retrait à gauche, en points, pour le paragraphe.|
||[lineSpacing](/javascript/api/word/word.paragraph#linespacing)|Obtient ou définit l’interligne, en points, pour le paragraphe spécifié.|
||[LineUnitAfter,](/javascript/api/word/word.paragraph#lineunitafter)|Obtient ou définit la quantité d’espace, dans le quadrillage, après le paragraphe.|
||[LineUnitBefore,](/javascript/api/word/word.paragraph#lineunitbefore)|Obtient ou définit la quantité d’espace, en lignes de quadrillage, avant le paragraphe.|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchprefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchwildcards)||
||[OutlineLevel,](/javascript/api/word/word.paragraph#outlinelevel)|Obtient ou définit le niveau hiérarchique pour le paragraphe.|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|Obtient la collection d’objets de contrôle de contenu dans le paragraphe.|
||[police](/javascript/api/word/word.paragraph#font)|Obtient le format de texte du paragraphe.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|Obtient la collection d’objets InlinePicture dans le paragraphe.|
||[ParentContentControl,](/javascript/api/word/word.paragraph#parentcontentcontrol)|Obtient le contrôle de contenu qui contient le paragraphe.|
||[text](/javascript/api/word/word.paragraph#text)|Obtient le texte du paragraphe.|
||[rightIndent](/javascript/api/word/word.paragraph#rightindent)|Obtient ou définit la valeur de retrait à droite, en points, pour le paragraphe.|
||[recherche (Texted’origine : chaîne, searchOptions ?: Word. SearchOptions \| {ignorepunct, ?: Boolean ignorespace, ?: Boolean MatchCase ?: Boolean matchPrefix ?: Boolean matchSuffix ?: Boolean-MatchWholeWord ?: Boolean matchWildcards ?: Boolean})](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié sur l’étendue de l’objet Paragraph.|
||[Select (selectionMode ?: Word. SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|Sélectionne le paragraphe et y accède via l’interface utilisateur de Word.|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceafter)|Obtient ou définit l’espacement, en points, après le paragraphe.|
||[spaceBefore](/javascript/api/word/word.paragraph#spacebefore)|Obtient ou définit l’espacement, en points, avant le paragraphe.|
||[style](/javascript/api/word/word.paragraph#style)|Obtient ou définit le nom du style du paragraphe.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|Efface le contenu de l’objet de plage.|
||[delete()](/javascript/api/word/word.range#delete--)|Supprime la plage et son contenu du document.|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Obtient une représentation HTML de l’objet de plage.|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Obtient la représentation OOXML de l’objet de plage.|
||[Ignorepunct,](/javascript/api/word/word.range#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.range#ignorespace)||
||[insertBreak (breakType : Word. BreakType, insertLocation : Word. InsertLocation)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|Encadre l’objet de plage avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64 (base64File : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|Insère un document à l’emplacement spécifié.|
||[insertHtml (HTML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|Insère du code HTML à l’emplacement spécifié.|
||[insertOoxml (OOXML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|Insère du code OOXML à l’emplacement spécifié.|
||[insertParagraph (paragraphText : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText (Text : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|Insère du texte à l’emplacement spécifié.|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[matchPrefix](/javascript/api/word/word.range#matchprefix)||
||[matchSuffix](/javascript/api/word/word.range#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|Obtient la collection d’objets de contrôle de contenu dans la plage.|
||[police](/javascript/api/word/word.range#font)|Obtient le format de texte de la plage.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Obtient la collection d’objets Paragraph dans la plage.|
||[ParentContentControl,](/javascript/api/word/word.range#parentcontentcontrol)|Obtient le contrôle de contenu qui contient la plage.|
||[text](/javascript/api/word/word.range#text)|Obtient le texte de la plage.|
||[recherche (Texted’origine : chaîne, searchOptions ?: Word. SearchOptions \| {ignorepunct, ?: Boolean ignorespace, ?: Boolean MatchCase ?: Boolean matchPrefix ?: Boolean matchSuffix ?: Boolean-MatchWholeWord ?: Boolean matchWildcards ?: Boolean})](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié sur l’étendue de l’objet Range.|
||[Select (selectionMode ?: Word. SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|Sélectionne la plage et y accède via l’interface utilisateur de Word.|
||[style](/javascript/api/word/word.range#style)|Obtient ou définit le nom du style de la plage.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[Ignorepunct,](/javascript/api/word/word.searchoptions#ignorepunct)|Obtient ou définit une valeur indiquant si toutes les marques de ponctuation entre les mots doivent être ignorées.|
||[ignorespace,](/javascript/api/word/word.searchoptions#ignorespace)|Obtient ou définit une valeur qui indique s’il faut ignorer tous les espaces entre les mots.|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|Obtient ou définit une valeur indiquant si la recherche respecte la casse.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchprefix)|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui commencent par la chaîne entrée.|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchsuffix)|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui se terminent par la chaîne entrée.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchwholeword)|Obtient ou définit une valeur indiquant si la recherche doit uniquement porter sur des mots entiers et exclure le texte s’il est inclus dans un mot plus long.|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchwildcards)|Obtient ou définit une valeur indiquant si la recherche est effectuée à l’aide d’opérateurs de recherche spéciaux.|
|[Section](/javascript/api/word/word.section)|[getFooter (type : Word. HeaderFooterType)](/javascript/api/word/word.section#getfooter-type-)|Obtient l’un des pieds de page de la section.|
||[getHeader (type : Word. HeaderFooterType)](/javascript/api/word/word.section#getheader-type-)|Obtient l’un des en-têtes de la section.|
||[body](/javascript/api/word/word.section#body)|Obtient l’objet Body de la section.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
