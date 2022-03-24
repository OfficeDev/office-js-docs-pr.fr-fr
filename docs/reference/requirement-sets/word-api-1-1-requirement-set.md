---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1.1
description: Détails sur l’ensemble de conditions requises WordApi 1.1.
ms.date: 11/01/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: dfcb1954cd9522de6165130cc115fddbb5f3ec45
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744212"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Nouveautés de l’API JavaScript 1.1 pour Word

WordApi 1.1 est le premier ensemble de conditions requises de l’API JavaScript pour Word. Il s’agit du seul ensemble de conditions requises de l’API Word pris en charge par Word 2016.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de l’ensemble de conditions requises de l’API JavaScript pour Word 1.1. Pour afficher la documentation de référence de l’API pour toutes les API pris en charge par l’ensemble de conditions requises de l’API JavaScript pour Word 1.1, voir API Word dans l’ensemble de conditions requises [1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Classe | Champs | Description |
|:---|:---|:---|
|[Corps](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#word-word-body-clear-member(1))|Efface le contenu de l’objet de corps.|
||[contentControls](/javascript/api/word/word.body#word-word-body-contentcontrols-member)|Obtient la collection d’objets de contrôle de contenu de texte enrichi dans le corps.|
||[police](/javascript/api/word/word.body#word-word-body-font-member)|Obtient le format de texte du corps.|
||[getHtml()](/javascript/api/word/word.body#word-word-body-gethtml-member(1))|Obtient une représentation HTML de l’objet body.|
||[getOoxml()](/javascript/api/word/word.body#word-word-body-getooxml-member(1))|Obtient la représentation OOXML (Office Open XML) de l’objet de corps.|
||[ignorePunct](/javascript/api/word/word.body#word-word-body-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.body#word-word-body-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.body#word-word-body-inlinepictures-member)|Obtient la collection d’objets InlinePicture dans le corps.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertbreak-member(1))|Insère un saut à l’emplacement spécifié du document principal.|
||[insertContentControl()](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|Encadre l’objet de corps avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1))|Insère un document dans le corps à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserthtml-member(1))|Insère du code HTML à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertooxml-member(1))|Insère du code OOXML à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertparagraph-member(1))|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserttext-member(1))|Insère du texte dans le corps à l’emplacement spécifié.|
||[matchCase](/javascript/api/word/word.body#word-word-body-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.body#word-word-body-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.body#word-word-body-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.body#word-word-body-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.body#word-word-body-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.body#word-word-body-paragraphs-member)|Obtient la collection d’objets de paragraphe dans le corps.|
||[parentContentControl](/javascript/api/word/word.body#word-word-body-parentcontentcontrol-member)|Obtient le contrôle de contenu qui contient le corps.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.body#word-word-body-search-member(1))|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de l’objet body.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#word-word-body-select-member(1))|Sélectionne le corps et y accède via l’interface utilisateur de Word.|
||[style](/javascript/api/word/word.body#word-word-body-style-member)|Obtient ou définit le nom du style du corps.|
||[text](/javascript/api/word/word.body#word-word-body-text-member)|Obtient le texte du corps.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[apparence](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-appearance-member)|Obtient ou définit l’apparence du contrôle de contenu.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotdelete-member)|Obtient ou définit une valeur qui indique si l’utilisateur peut supprimer le contrôle de contenu.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotedit-member)|Obtient ou définit une valeur qui indique si l’utilisateur peut modifier le contenu du contrôle.|
||[clear()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-clear-member(1))|Efface le contenu du contrôle de contenu.|
||[color](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-color-member)|Obtient ou définit la couleur du contrôle de contenu.|
||[contentControls](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-contentcontrols-member)|Obtient la collection d’objets de contrôle de contenu compris dans le contrôle de contenu.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-delete-member(1))|Supprime le contrôle de contenu et son contenu.|
||[police](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-font-member)|Obtient le format de texte du contrôle de contenu.|
||[getHtml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gethtml-member(1))|Obtient une représentation HTML de l’objet de contrôle de contenu.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getooxml-member(1))|Obtient la représentation Office Open XML (OOXML) de l’objet de contrôle de contenu.|
||[id](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-id-member)|Obtient un entier qui représente l’identificateur du contrôle de contenu.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inlinepictures-member)|Obtient la collection d’objets inlinePicture du contrôle de contenu.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertbreak-member(1))|Insère un saut à l’emplacement spécifié du document principal.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertfilefrombase64-member(1))|Insère un document dans le contrôle de contenu à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserthtml-member(1))|Insère du code HTML dans le contrôle de contenu, à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertooxml-member(1))|Insère du contenu OOXML dans le contrôle de contenu à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertparagraph-member(1))|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttext-member(1))|Insère du texte dans le contrôle de contenu, à l’emplacement spécifié.|
||[matchCase](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-paragraphs-member)|Obtient la collection d’objets de paragraphe dans le contrôle de contenu.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrol-member)|Obtient le contrôle de contenu qui contient le contrôle de contenu spécifié.|
||[placeholderText](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-placeholdertext-member)|Obtient ou définit le texte de l’espace réservé du contrôle de contenu.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-removewhenedited-member)|Obtient ou définit une valeur qui indique si le contrôle de contenu doit être supprimé après modification.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-search-member(1))|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de l’objet de contrôle de contenu.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-select-member(1))|Sélectionne le contrôle de contenu.|
||[style](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-style-member)|Obtient ou définit le nom du style du contrôle de contenu.|
||[tag](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tag-member)|Obtient ou définit un indicateur pour identifier un contrôle de contenu.|
||[text](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-text-member)|Obtient le texte du contrôle de contenu.|
||[title](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-title-member)|Obtient ou définit le titre d’un contrôle de contenu.|
||[type](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-type-member)|Obtient le type du contrôle de contenu.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyid-member(1))|Obtient un contrôle de contenu par son identificateur.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytag-member(1))|Obtient les contrôles de contenu qui portent l’indicateur spécifié.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytitle-member(1))|Obtient les contrôles de contenu qui ont le titre spécifié.|
||[getItem(index : numérique)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getitem-member(1))|Obtient un contrôle de contenu par son index dans la collection.|
||[items](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[body](/javascript/api/word/word.document#word-word-document-body-member)|Obtient l’objet body du document principal.|
||[contentControls](/javascript/api/word/word.document#word-word-document-contentcontrols-member)|Obtient la collection d’objets de contrôle de contenu dans le document.|
||[getSelection()](/javascript/api/word/word.document#word-word-document-getselection-member(1))|Obtient la sélection actuelle du document.|
||[save()](/javascript/api/word/word.document#word-word-document-save-member(1))|Enregistre le document.|
||[saved](/javascript/api/word/word.document#word-word-document-saved-member)|Indique si les modifications apportées au document ont été enregistrées.|
||[sections](/javascript/api/word/word.document#word-word-document-sections-member)|Obtient la collection d’objets de section dans le document.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#word-word-font-bold-member)|Obtient ou définit une valeur qui indique si la police en gras.|
||[color](/javascript/api/word/word.font#word-word-font-color-member)|Obtient ou définit la couleur de la police spécifiée.|
||[doubleStrikeThrough](/javascript/api/word/word.font#word-word-font-doublestrikethrough-member)|Obtient ou définit une valeur qui indique si la police a un double strikethrough.|
||[highlightColor](/javascript/api/word/word.font#word-word-font-highlightcolor-member)|Obtient ou définit la couleur de surbrillage.|
||[italic](/javascript/api/word/word.font#word-word-font-italic-member)|Obtient ou définit une valeur qui indique si la police est en italique.|
||[name](/javascript/api/word/word.font#word-word-font-name-member)|Obtient ou définit une valeur qui représente le nom de la police.|
||[taille](/javascript/api/word/word.font#word-word-font-size-member)|Obtient ou définit une valeur qui représente la taille de police en points.|
||[strikeThrough](/javascript/api/word/word.font#word-word-font-strikethrough-member)|Obtient ou définit une valeur qui indique si la police a un strikethrough.|
||[Subscript](/javascript/api/word/word.font#word-word-font-subscript-member)|Obtient ou définit une valeur qui indique si la police correspond à du texte mis en indice.|
||[superscript](/javascript/api/word/word.font#word-word-font-superscript-member)|Obtient ou définit une valeur qui indique si la police correspond à du texte en exposant.|
||[underline](/javascript/api/word/word.font#word-word-font-underline-member)|Obtient ou définit une valeur qui indique le type de trait de soulignement de la police.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttextdescription-member)|Obtient ou définit une chaîne qui représente le texte de remplacement associé à l’image fixe.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttexttitle-member)|Obtient ou définit une chaîne qui contient le titre de l’image incluse.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getbase64imagesrc-member(1))|Obtient la représentation de chaîne encodée au format Base64 de l’image incluse.|
||[height](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-height-member)|Obtient ou définit un nombre qui décrit la hauteur de l’image incluse.|
||[lien hypertexte](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-hyperlink-member)|Obtient ou définit un lien hypertexte sur l’image.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertcontentcontrol-member(1))|Encadre l’image incluse avec un contrôle de contenu de texte enrichi.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-lockaspectratio-member)|Obtient ou définit une valeur qui indique si l’image incluse conserve ses proportions d’origine lorsque vous la redimensionnez.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrol-member)|Obtient le contrôle de contenu qui contient l’image incluse.|
||[width](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-width-member)|Obtient ou définit un nombre qui décrit la largeur de l’image incluse.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignement](/javascript/api/word/word.paragraph#word-word-paragraph-alignment-member)|Obtient ou définit l’alignement d’un paragraphe.|
||[clear()](/javascript/api/word/word.paragraph#word-word-paragraph-clear-member(1))|Efface le contenu de l’objet de paragraphe.|
||[contentControls](/javascript/api/word/word.paragraph#word-word-paragraph-contentcontrols-member)|Obtient la collection d’objets de contrôle de contenu dans le paragraphe.|
||[delete()](/javascript/api/word/word.paragraph#word-word-paragraph-delete-member(1))|Supprime le paragraphe et son contenu du document.|
||[firstLineIndent](/javascript/api/word/word.paragraph#word-word-paragraph-firstlineindent-member)|Renvoie ou définit la valeur, en points, du retrait de première ligne ou du retrait négatif.|
||[police](/javascript/api/word/word.paragraph#word-word-paragraph-font-member)|Obtient le format de texte du paragraphe.|
||[getHtml()](/javascript/api/word/word.paragraph#word-word-paragraph-gethtml-member(1))|Obtient une représentation HTML de l’objet de paragraphe.|
||[getOoxml()](/javascript/api/word/word.paragraph#word-word-paragraph-getooxml-member(1))|Obtient la représentation Office Open XML (OOXML) de l’objet de paragraphe.|
||[ignorePunct](/javascript/api/word/word.paragraph#word-word-paragraph-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.paragraph#word-word-paragraph-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.paragraph#word-word-paragraph-inlinepictures-member)|Obtient la collection d’objets InlinePicture dans le paragraphe.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertbreak-member(1))|Insère un saut à l’emplacement spécifié du document principal.|
||[insertContentControl()](/javascript/api/word/word.paragraph#word-word-paragraph-insertcontentcontrol-member(1))|Encadre l’objet de paragraphe avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertfilefrombase64-member(1))|Insère un document dans le paragraphe à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserthtml-member(1))|Insère du code HTML dans le paragraphe à l’emplacement spécifié.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertinlinepicturefrombase64-member(1))|Insère une image dans le paragraphe à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertooxml-member(1))|Insère du texte OOXML dans le paragraphe à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertparagraph-member(1))|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserttext-member(1))|Insère du texte dans le paragraphe à l’emplacement spécifié.|
||[leftIndent](/javascript/api/word/word.paragraph#word-word-paragraph-leftindent-member)|Obtient ou définit la valeur de retrait à gauche, en points, pour le paragraphe.|
||[lineSpacing](/javascript/api/word/word.paragraph#word-word-paragraph-linespacing-member)|Obtient ou définit l’interligne, en points, pour le paragraphe spécifié.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitafter-member)|Obtient ou définit l’espacement, en lignes de grille, après le paragraphe.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitbefore-member)|Obtient ou définit la quantité d’espace, en lignes de quadrillage, avant le paragraphe.|
||[matchCase](/javascript/api/word/word.paragraph#word-word-paragraph-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.paragraph#word-word-paragraph-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.paragraph#word-word-paragraph-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.paragraph#word-word-paragraph-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.paragraph#word-word-paragraph-matchwildcards-member)||
||[outlineLevel](/javascript/api/word/word.paragraph#word-word-paragraph-outlinelevel-member)|Obtient ou définit le niveau hiérarchique pour le paragraphe.|
||[parentContentControl](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrol-member)|Obtient le contrôle de contenu qui contient le paragraphe.|
||[rightIndent](/javascript/api/word/word.paragraph#word-word-paragraph-rightindent-member)|Obtient ou définit la valeur de retrait à droite, en points, pour le paragraphe.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.paragraph#word-word-paragraph-search-member(1))|Effectue une recherche avec les objets SearchOption spécifiés dans l’étendue de l’objet de paragraphe.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#word-word-paragraph-select-member(1))|Sélectionne le paragraphe et y accède via l’interface utilisateur de Word.|
||[spaceAfter](/javascript/api/word/word.paragraph#word-word-paragraph-spaceafter-member)|Obtient ou définit l’espacement, en points, après le paragraphe.|
||[spaceBefore](/javascript/api/word/word.paragraph#word-word-paragraph-spacebefore-member)|Obtient ou définit l’espacement, en points, avant le paragraphe.|
||[style](/javascript/api/word/word.paragraph#word-word-paragraph-style-member)|Obtient ou définit le nom du style du paragraphe.|
||[text](/javascript/api/word/word.paragraph#word-word-paragraph-text-member)|Obtient le texte du paragraphe.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#word-word-range-clear-member(1))|Efface le contenu de l’objet de plage.|
||[contentControls](/javascript/api/word/word.range#word-word-range-contentcontrols-member)|Obtient la collection d’objets de contrôle de contenu dans la plage.|
||[delete()](/javascript/api/word/word.range#word-word-range-delete-member(1))|Supprime la plage et son contenu du document.|
||[police](/javascript/api/word/word.range#word-word-range-font-member)|Obtient le format de texte de la plage.|
||[getHtml()](/javascript/api/word/word.range#word-word-range-gethtml-member(1))|Obtient une représentation HTML de l’objet de plage.|
||[getOoxml()](/javascript/api/word/word.range#word-word-range-getooxml-member(1))|Obtient la représentation OOXML de l’objet de plage.|
||[ignorePunct](/javascript/api/word/word.range#word-word-range-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.range#word-word-range-ignorespace-member)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertbreak-member(1))|Insère un saut à l’emplacement spécifié du document principal.|
||[insertContentControl()](/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1))|Encadre l’objet de plage avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertfilefrombase64-member(1))|Insère un document à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserthtml-member(1))|Insère du code HTML à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertooxml-member(1))|Insère du code OOXML à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertparagraph-member(1))|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserttext-member(1))|Insère du texte à l’emplacement spécifié.|
||[matchCase](/javascript/api/word/word.range#word-word-range-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.range#word-word-range-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.range#word-word-range-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.range#word-word-range-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.range#word-word-range-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.range#word-word-range-paragraphs-member)|Obtient la collection d’objets de paragraphe de la plage.|
||[parentContentControl](/javascript/api/word/word.range#word-word-range-parentcontentcontrol-member)|Obtient le contrôle de contenu qui contient la plage.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.range#word-word-range-search-member(1))|Effectue une recherche avec les searchOptions spécifiées sur l’étendue de l’objet de plage.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#word-word-range-select-member(1))|Sélectionne la plage et y accède via l’interface utilisateur de Word.|
||[style](/javascript/api/word/word.range#word-word-range-style-member)|Obtient ou définit le nom du style de la plage.|
||[text](/javascript/api/word/word.range#word-word-range-text-member)|Obtient le texte de la plage.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#word-word-rangecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorepunct-member)|Obtient ou définit une valeur indiquant si toutes les marques de ponctuation entre les mots doivent être ignorées.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorespace-member)|Obtient ou définit une valeur qui indique s’il faut ignorer tous les espaces entre les mots.|
||[matchCase](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchcase-member)|Obtient ou définit une valeur indiquant si la recherche respecte la casse.|
||[matchPrefix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchprefix-member)|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui commencent par la chaîne entrée.|
||[matchSuffix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchsuffix-member)|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui se terminent par la chaîne entrée.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwholeword-member)|Obtient ou définit une valeur indiquant si la recherche doit uniquement porter sur des mots entiers et exclure le texte s’il est inclus dans un mot plus long.|
||[matchWildcards](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwildcards-member)|Obtient ou définit une valeur indiquant si la recherche est effectuée à l’aide d’opérateurs de recherche spéciaux.|
|[Section](/javascript/api/word/word.section)|[body](/javascript/api/word/word.section#word-word-section-body-member)|Obtient l’objet body de la section.|
||[getFooter(type: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getfooter-member(1))|Obtient l’un des pieds de page de la section.|
||[getHeader(type : Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getheader-member(1))|Obtient l’un des en-têtes de la section.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
