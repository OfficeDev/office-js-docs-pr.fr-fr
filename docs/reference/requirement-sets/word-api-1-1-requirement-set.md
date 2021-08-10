---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1.1
description: Détails sur l’ensemble de conditions requises WordApi 1.1
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: c7b1adfa1af76f9994ced793dfddcf457cf733858fd27ba0ef763a67c35611c2
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092452"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Nouveautés de l’API JavaScript 1.1 pour Word

WordApi 1.1 est le premier ensemble de conditions requises de l’API JavaScript pour Word. Il s’agit du seul ensemble de conditions requises de l’API Word pris en charge par Word 2016.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de l’ensemble de conditions requises de l’API JavaScript pour Word 1.1. Pour afficher la documentation de référence de l’API pour toutes les API pris en charge par l’ensemble de conditions requises de l’API JavaScript pour Word 1.1, voir API Word dans l’ensemble de conditions requises [1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Classe | Champs | Description |
|:---|:---|:---|
|[Corps](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear__)|Efface le contenu de l’objet de corps.|
||[getHtml()](/javascript/api/word/word.body#getHtml__)|Obtient une représentation HTML de l’objet body.|
||[getOoxml()](/javascript/api/word/word.body#getOoxml__)|Obtient la représentation OOXML (Office Open XML) de l’objet de corps.|
||[ignorePunct](/javascript/api/word/word.body#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertBreak_breakType__insertLocation_)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertContentControl()](/javascript/api/word/word.body#insertContentControl__)|Encadre l’objet de corps avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertFileFromBase64_base64File__insertLocation_)|Insère un document dans le corps à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertHtml_html__insertLocation_)|Insère du code HTML à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertOoxml_ooxml__insertLocation_)|Insère du code OOXML à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertParagraph_paragraphText__insertLocation_)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertText_text__insertLocation_)|Insère du texte dans le corps à l’emplacement spécifié.|
||[matchCase](/javascript/api/word/word.body#matchCase)||
||[matchPrefix](/javascript/api/word/word.body#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.body#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.body#matchWildcards)||
||[contentControls](/javascript/api/word/word.body#contentControls)|Obtient la collection d’objets de contrôle de contenu de texte enrichi dans le corps.|
||[police](/javascript/api/word/word.body#font)|Obtient le format de texte du corps.|
||[inlinePictures](/javascript/api/word/word.body#inlinePictures)|Obtient la collection d’objets InlinePicture dans le corps.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Obtient la collection d’objets de paragraphe dans le corps.|
||[parentContentControl](/javascript/api/word/word.body#parentContentControl)|Obtient le contrôle de contenu qui contient le corps.|
||[text](/javascript/api/word/word.body#text)|Obtient le texte du corps.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.body#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de l’objet body.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#select_selectionMode_)|Sélectionne le corps et y accède via l’interface utilisateur de Word.|
||[style](/javascript/api/word/word.body#style)|Obtient ou définit le nom du style du corps.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[apparence](/javascript/api/word/word.contentcontrol#appearance)|Obtient ou définit l’apparence du contrôle de contenu.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotDelete)|Obtient ou définit une valeur qui indique si l’utilisateur peut supprimer le contrôle de contenu.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotEdit)|Obtient ou définit une valeur qui indique si l’utilisateur peut modifier le contenu du contrôle.|
||[clear()](/javascript/api/word/word.contentcontrol#clear__)|Efface le contenu du contrôle de contenu.|
||[color](/javascript/api/word/word.contentcontrol#color)|Obtient ou définit la couleur du contrôle de contenu.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#delete_keepContent_)|Supprime le contrôle de contenu et son contenu.|
||[getHtml()](/javascript/api/word/word.contentcontrol#getHtml__)|Obtient une représentation HTML de l’objet de contrôle de contenu.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getOoxml__)|Obtient la représentation Office Open XML (OOXML) de l’objet de contrôle de contenu.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertBreak_breakType__insertLocation_)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertFileFromBase64_base64File__insertLocation_)|Insère un document dans le contrôle de contenu à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertHtml_html__insertLocation_)|Insère du code HTML dans le contrôle de contenu, à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertOoxml_ooxml__insertLocation_)|Insère du contenu OOXML dans le contrôle de contenu à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertParagraph_paragraphText__insertLocation_)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertText_text__insertLocation_)|Insère du texte dans le contrôle de contenu, à l’emplacement spécifié.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchCase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchWildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholderText)|Obtient ou définit le texte de l’espace réservé du contrôle de contenu.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentControls)|Obtient la collection d’objets de contrôle de contenu compris dans le contrôle de contenu.|
||[police](/javascript/api/word/word.contentcontrol#font)|Obtient le format de texte du contrôle de contenu.|
||[id](/javascript/api/word/word.contentcontrol#id)|Obtient un entier qui représente l’identificateur du contrôle de contenu.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinePictures)|Obtient la collection d’objets inlinePicture du contrôle de contenu.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Obtient la collection d’objets de paragraphe du contrôle de contenu.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentContentControl)|Obtient le contrôle de contenu qui contient le contrôle de contenu spécifié.|
||[text](/javascript/api/word/word.contentcontrol#text)|Obtient le texte du contrôle de contenu.|
||[type](/javascript/api/word/word.contentcontrol#type)|Obtient le type du contrôle de contenu.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removeWhenEdited)|Obtient ou définit une valeur qui indique si le contrôle de contenu doit être supprimé après modification.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.contentcontrol#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de l’objet de contrôle de contenu.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#select_selectionMode_)|Sélectionne le contrôle de contenu.|
||[style](/javascript/api/word/word.contentcontrol#style)|Obtient ou définit le nom du style du contrôle de contenu.|
||[tag](/javascript/api/word/word.contentcontrol#tag)|Obtient ou définit un indicateur pour identifier un contrôle de contenu.|
||[title](/javascript/api/word/word.contentcontrol#title)|Obtient ou définit le titre d’un contrôle de contenu.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getById_id_)|Obtient un contrôle de contenu par son identificateur.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getByTag_tag_)|Obtient les contrôles de contenu qui portent l’indicateur spécifié.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getByTitle_title_)|Obtient les contrôles de contenu qui ont le titre spécifié.|
||[getItem(index : numérique)](/javascript/api/word/word.contentcontrolcollection#getItem_index_)|Obtient un contrôle de contenu par son index dans la collection.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[getSelection()](/javascript/api/word/word.document#getSelection__)|Obtient la sélection actuelle du document.|
||[body](/javascript/api/word/word.document#body)|Obtient l’objet body du document.|
||[contentControls](/javascript/api/word/word.document#contentControls)|Obtient la collection d’objets de contrôle de contenu dans le document.|
||[saved](/javascript/api/word/word.document#saved)|Indique si les modifications apportées au document ont été enregistrées.|
||[sections](/javascript/api/word/word.document#sections)|Obtient la collection d’objets de section dans le document.|
||[save()](/javascript/api/word/word.document#save__)|Enregistre le document.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Obtient ou définit une valeur qui indique si la police en gras.|
||[color](/javascript/api/word/word.font#color)|Obtient ou définit la couleur de la police spécifiée.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doubleStrikeThrough)|Obtient ou définit une valeur qui indique si la police a un double strikethrough.|
||[highlightColor](/javascript/api/word/word.font#highlightColor)|Obtient ou définit la couleur de surbrillage.|
||[italic](/javascript/api/word/word.font#italic)|Obtient ou définit une valeur qui indique si la police est en italique.|
||[name](/javascript/api/word/word.font#name)|Obtient ou définit une valeur qui représente le nom de la police.|
||[size](/javascript/api/word/word.font#size)|Obtient ou définit une valeur qui représente la taille de police en points.|
||[strikeThrough](/javascript/api/word/word.font#strikeThrough)|Obtient ou définit une valeur qui indique si la police a un strikethrough.|
||[Subscript](/javascript/api/word/word.font#subscript)|Obtient ou définit une valeur qui indique si la police correspond à du texte mis en indice.|
||[superscript](/javascript/api/word/word.font#superscript)|Obtient ou définit une valeur qui indique si la police correspond à du texte en exposant.|
||[underline](/javascript/api/word/word.font#underline)|Obtient ou définit une valeur qui indique le type de trait de soulignement de la police.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#altTextDescription)|Obtient ou définit une chaîne qui représente le texte de remplacement associé à l’image fixe.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#altTextTitle)|Obtient ou définit une chaîne qui contient le titre de l’image incluse.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getBase64ImageSrc__)|Obtient la représentation de chaîne encodée au format Base64 de l’image incluse.|
||[height](/javascript/api/word/word.inlinepicture#height)|Obtient ou définit un nombre qui décrit la hauteur de l’image incluse.|
||[lien hypertexte](/javascript/api/word/word.inlinepicture#hyperlink)|Obtient ou définit un lien hypertexte sur l’image.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertContentControl__)|Encadre l’image incluse avec un contrôle de contenu de texte enrichi.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockAspectRatio)|Obtient ou définit une valeur qui indique si l’image incluse conserve ses proportions d’origine lorsque vous la redimensionnez.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentContentControl)|Obtient le contrôle de contenu qui contient l’image incluse.|
||[width](/javascript/api/word/word.inlinepicture#width)|Obtient ou définit un nombre qui décrit la largeur de l’image incluse.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignement](/javascript/api/word/word.paragraph#alignment)|Obtient ou définit l’alignement d’un paragraphe.|
||[clear()](/javascript/api/word/word.paragraph#clear__)|Efface le contenu de l’objet de paragraphe.|
||[delete()](/javascript/api/word/word.paragraph#delete__)|Supprime le paragraphe et son contenu du document.|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstLineIndent)|Renvoie ou définit la valeur, en points, du retrait de première ligne ou du retrait négatif.|
||[getHtml()](/javascript/api/word/word.paragraph#getHtml__)|Obtient une représentation HTML de l’objet de paragraphe.|
||[getOoxml()](/javascript/api/word/word.paragraph#getOoxml__)|Obtient la représentation Office Open XML (OOXML) de l’objet de paragraphe.|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertBreak_breakType__insertLocation_)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertContentControl__)|Encadre l’objet de paragraphe avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertFileFromBase64_base64File__insertLocation_)|Insère un document dans le paragraphe à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertHtml_html__insertLocation_)|Insère du code HTML dans le paragraphe à l’emplacement spécifié.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insère une image dans le paragraphe à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertOoxml_ooxml__insertLocation_)|Insère du texte OOXML dans le paragraphe à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertParagraph_paragraphText__insertLocation_)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertText_text__insertLocation_)|Insère du texte dans le paragraphe à l’emplacement spécifié.|
||[leftIndent](/javascript/api/word/word.paragraph#leftIndent)|Obtient ou définit la valeur de retrait à gauche, en points, pour le paragraphe.|
||[lineSpacing](/javascript/api/word/word.paragraph#lineSpacing)|Obtient ou définit l’interligne, en points, pour le paragraphe spécifié.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineUnitAfter)|Obtient ou définit l’espacement, en lignes de grille, après le paragraphe.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineUnitBefore)|Obtient ou définit la quantité d’espace, en lignes de quadrillage, avant le paragraphe.|
||[matchCase](/javascript/api/word/word.paragraph#matchCase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchWildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlineLevel)|Obtient ou définit le niveau hiérarchique pour le paragraphe.|
||[contentControls](/javascript/api/word/word.paragraph#contentControls)|Obtient la collection d’objets de contrôle de contenu dans le paragraphe.|
||[police](/javascript/api/word/word.paragraph#font)|Obtient le format de texte du paragraphe.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinePictures)|Obtient la collection d’objets InlinePicture dans le paragraphe.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentContentControl)|Obtient le contrôle de contenu qui contient le paragraphe.|
||[text](/javascript/api/word/word.paragraph#text)|Obtient le texte du paragraphe.|
||[rightIndent](/javascript/api/word/word.paragraph#rightIndent)|Obtient ou définit la valeur de retrait à droite, en points, pour le paragraphe.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.paragraph#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Effectue une recherche avec les objets SearchOption spécifiés dans l’étendue de l’objet de paragraphe.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#select_selectionMode_)|Sélectionne le paragraphe et y accède via l’interface utilisateur de Word.|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceAfter)|Obtient ou définit l’espacement, en points, après le paragraphe.|
||[spaceBefore](/javascript/api/word/word.paragraph#spaceBefore)|Obtient ou définit l’espacement, en points, avant le paragraphe.|
||[style](/javascript/api/word/word.paragraph#style)|Obtient ou définit le nom du style du paragraphe.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear__)|Efface le contenu de l’objet de plage.|
||[delete()](/javascript/api/word/word.range#delete__)|Supprime la plage et son contenu du document.|
||[getHtml()](/javascript/api/word/word.range#getHtml__)|Obtient une représentation HTML de l’objet de plage.|
||[getOoxml()](/javascript/api/word/word.range#getOoxml__)|Obtient la représentation OOXML de l’objet de plage.|
||[ignorePunct](/javascript/api/word/word.range#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertBreak_breakType__insertLocation_)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertContentControl()](/javascript/api/word/word.range#insertContentControl__)|Encadre l’objet de plage avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertFileFromBase64_base64File__insertLocation_)|Insère un document à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertHtml_html__insertLocation_)|Insère du code HTML à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertOoxml_ooxml__insertLocation_)|Insère du code OOXML à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertParagraph_paragraphText__insertLocation_)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertText_text__insertLocation_)|Insère du texte à l’emplacement spécifié.|
||[matchCase](/javascript/api/word/word.range#matchCase)||
||[matchPrefix](/javascript/api/word/word.range#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.range#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.range#matchWildcards)||
||[contentControls](/javascript/api/word/word.range#contentControls)|Obtient la collection d’objets de contrôle de contenu dans la plage.|
||[police](/javascript/api/word/word.range#font)|Obtient le format de texte de la plage.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Obtient la collection d’objets de paragraphe de la plage.|
||[parentContentControl](/javascript/api/word/word.range#parentContentControl)|Obtient le contrôle de contenu qui contient la plage.|
||[text](/javascript/api/word/word.range#text)|Obtient le texte de la plage.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.range#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de l’objet de plage.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#select_selectionMode_)|Sélectionne la plage et y accède via l’interface utilisateur de Word.|
||[style](/javascript/api/word/word.range#style)|Obtient ou définit le nom du style de la plage.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorePunct)|Obtient ou définit une valeur indiquant si toutes les marques de ponctuation entre les mots doivent être ignorées.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignoreSpace)|Obtient ou définit une valeur qui indique s’il faut ignorer tous les espaces entre les mots.|
||[matchCase](/javascript/api/word/word.searchoptions#matchCase)|Obtient ou définit une valeur indiquant si la recherche respecte la casse.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchPrefix)|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui commencent par la chaîne entrée.|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchSuffix)|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui se terminent par la chaîne entrée.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchWholeWord)|Obtient ou définit une valeur indiquant si la recherche doit uniquement porter sur des mots entiers et exclure le texte s’il est inclus dans un mot plus long.|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchWildcards)|Obtient ou définit une valeur indiquant si la recherche est effectuée à l’aide d’opérateurs de recherche spéciaux.|
|[Section](/javascript/api/word/word.section)|[getFooter(type : Word.HeaderFooterType)](/javascript/api/word/word.section#getFooter_type_)|Obtient l’un des pieds de page de la section.|
||[getHeader(type : Word.HeaderFooterType)](/javascript/api/word/word.section#getHeader_type_)|Obtient l’un des en-têtes de la section.|
||[body](/javascript/api/word/word.section#body)|Obtient l’objet body de la section.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
