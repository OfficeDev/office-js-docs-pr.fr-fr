---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1,1
description: Détails sur l’ensemble de conditions requises WordApi 1,1
ms.date: 07/25/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: a2839a2553d42701956fd2e75a86564c133d9a93
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064913"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Nouveautés de l’API JavaScript pour Word 1,1

WordApi 1,1 est le premier ensemble de conditions requises de l’API JavaScript pour Word. Il s’agit du seul ensemble de conditions requises de l’API Word pris en charge par Word 2016.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Word 1,1. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’API JavaScript pour Word, ensemble de conditions requises 1,1, voir [API Word dans l’ensemble de conditions requises 1,1](/javascript/api/word?view=word-js-1.1).

| Class | Champs | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|Efface le contenu de l’objet de corps. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Obtient une représentation HTML de l’objet Body. Lorsqu’elle est affichée dans une page Web ou dans la visionneuse HTML, la mise en forme est une correspondance ferme, mais pas exacte, avec la mise en forme du document. Cette méthode ne renvoie pas exactement le même code HTML pour le même document sur différentes plateformes (Windows, Mac, etc.). Si vous avez besoin d’une fidélité exacte ou d’une cohérence `Body.getOoxml()` entre les plateformes, utilisez et convertissez le code XML renvoyé en html.|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|Obtient la représentation OOXML (Office Open XML) de l’objet de corps.|
||[Ignorepunct,](/javascript/api/word/word.body#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.body#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|Insère un saut à l’emplacement spécifié du document principal. La valeur insertLocation peut être « Start » (début) ou « End » (fin).|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|Encadre l’objet de corps avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64 (base64File: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|Insère un document dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertHtml (HTML: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|Insère du code HTML à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertOoxml (OOXML: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|Insère du code OOXML à l’emplacement spécifié.  La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertParagraph (paragraphText: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Start » (début) ou « End » (fin).|
||[insertText (Text: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|Insère du texte dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[matchPrefix](/javascript/api/word/word.body#matchprefix)||
||[matchSuffix](/javascript/api/word/word.body#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|Obtient la collection d’objets de contrôle de contenu de texte enrichi dans le corps. En lecture seule.|
||[police](/javascript/api/word/word.body#font)|Obtient le format de texte du corps. Utilisez cette valeur pour obtenir et définir le nom, la taille, la couleur et d’autres propriétés de la police. En lecture seule.|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|Obtient la collection d’objets InlinePicture dans le corps. La collection n’inclut pas d’images flottantes. En lecture seule.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Obtient la collection d’objets Paragraph dans le corps. En lecture seule.|
||[ParentContentControl,](/javascript/api/word/word.body#parentcontentcontrol)|Obtient le contrôle de contenu qui contient le corps. S’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[text](/javascript/api/word/word.body#text)|Obtient le texte du corps. Utilisez la méthode insertText pour insérer du texte. En lecture seule.|
||[recherche (Texted’origine: chaîne, searchOptions?: Word. SearchOptions)](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié sur l’étendue de l’objet Body. Les résultats de la recherche sont un ensemble d’objets de plage.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|Sélectionne le corps et y accède via l’interface utilisateur de Word.|
||[style](/javascript/api/word/word.body#style)|Obtient ou définit le nom de style du corps. Utilisez cette propriété pour les noms des styles personnalisés et localisés. Pour utiliser les styles prédéfinis qui sont portables entre différents paramètres régionaux, voir la propriété « styleBuiltIn ».|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[apparence](/javascript/api/word/word.contentcontrol#appearance)|Obtient ou définit l’apparence du contrôle de contenu. La valeur peut être’BoundingBox', 'Tags’ou’Hidden'.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotdelete)|Obtient ou définit une valeur qui indique si l’utilisateur peut supprimer le contrôle de contenu. Non compatible avec removeWhenEdited.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotedit)|Obtient ou définit une valeur qui indique si l’utilisateur peut modifier le contenu du contrôle.|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|Efface le contenu du contrôle de contenu. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.|
||[color](/javascript/api/word/word.contentcontrol#color)|Obtient ou définit la couleur du contrôle de contenu. La couleur est spécifiée au format «#RRGGBB» ou en utilisant le nom de la couleur.|
||[Delete (keepContent: valeur booléenne)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|Supprime le contrôle de contenu et son contenu. Si keepContent est défini sur true, le contenu n’est pas supprimé.|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|Obtient une représentation HTML de l’objet de contrôle de contenu. Lorsqu’elle est affichée dans une page Web ou dans la visionneuse HTML, la mise en forme est une correspondance ferme, mais pas exacte, avec la mise en forme du document. Cette méthode ne renvoie pas exactement le même code HTML pour le même document sur différentes plateformes (Windows, Mac, etc.). Si vous avez besoin d’une fidélité exacte ou d’une cohérence `ContentControl.getOoxml()` entre les plateformes, utilisez et convertissez le code XML renvoyé en html.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|Obtient la représentation Office Open XML (OOXML) de l’objet de contrôle de contenu.|
||[Ignorepunct,](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|Insère un saut à l’emplacement spécifié du document principal. La valeur insertLocation peut être «Start», «end», «Before» ou «after». Cette méthode ne peut pas être utilisée avec les contrôles de contenu «RichTextTable», «RichTextTableRow» et «RichTextTableCell».|
||[insertFileFromBase64 (base64File: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|Insère un document dans le contrôle de contenu à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertHtml (HTML: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|Insère du code HTML dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertOoxml (OOXML: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|Insère OOXML dans le contrôle de contenu à l’emplacement spécifié.  La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertParagraph (paragraphText: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être «Start», «end», «Before» ou «after».|
||[insertText (Text: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|Insère du texte dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchprefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[PlaceholderText,](/javascript/api/word/word.contentcontrol#placeholdertext)|Obtient ou définit le texte de l’espace réservé du contrôle de contenu. Ce texte apparaît de façon estompée lorsque le contrôle de contenu est vide.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|Obtient la collection d’objets de contrôle de contenu compris dans le contrôle de contenu. En lecture seule.|
||[police](/javascript/api/word/word.contentcontrol#font)|Obtient le format de texte du contrôle de contenu. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|
||[id](/javascript/api/word/word.contentcontrol#id)|Obtient un entier qui représente l’identificateur du contrôle de contenu. En lecture seule.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|Obtient la collection d’objets inlinePicture du contrôle de contenu. La collection n’inclut pas d’images flottantes. En lecture seule.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Obtient la collection d’objets de paragraphe du contrôle de contenu. En lecture seule.|
||[ParentContentControl,](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|Obtient le contrôle de contenu qui contient le contrôle de contenu spécifié. S’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[text](/javascript/api/word/word.contentcontrol#text)|Obtient le texte du contrôle de contenu. En lecture seule.|
||[type](/javascript/api/word/word.contentcontrol#type)|Obtient le type du contrôle de contenu. Actuellement, seuls les contrôles de contenu à texte enrichi sont pris en charge. En lecture seule.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removewhenedited)|Obtient ou définit une valeur qui indique si le contrôle de contenu doit être supprimé après modification. Non compatible avec cannotDelete.|
||[recherche (Texted’origine: chaîne, searchOptions?: Word. SearchOptions)](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié dans l’étendue de l’objet de contrôle de contenu. Les résultats de la recherche sont un ensemble d’objets de plage.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|Sélectionne le contrôle de contenu. Word fait défiler le document jusqu’à accéder à la sélection.|
||[style](/javascript/api/word/word.contentcontrol#style)|Obtient ou définit le nom du style pour le contrôle de contenu. Utilisez cette propriété pour les noms des styles personnalisés et localisés. Pour utiliser les styles prédéfinis qui sont portables entre différents paramètres régionaux, voir la propriété « styleBuiltIn ».|
||[Numéro](/javascript/api/word/word.contentcontrol#tag)|Obtient ou définit un indicateur pour identifier un contrôle de contenu.|
||[title](/javascript/api/word/word.contentcontrol#title)|Obtient ou définit le titre d’un contrôle de contenu.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|Obtient un contrôle de contenu par son identificateur. Lève une exception s’il n’existe pas de contrôle de contenu avec l’identificateur dans cette collection.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|Obtient les contrôles de contenu qui portent l’indicateur spécifié.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|Obtient les contrôles de contenu qui ont le titre spécifié.|
||[getItem(index : numérique)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|Obtient un contrôle de contenu par son index dans la collection.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Document](/javascript/api/word/word.document)|[getSelection ()](/javascript/api/word/word.document#getselection--)|Obtient la sélection actuelle du document. Les sélections multiples ne sont pas prises en charge.|
||[body](/javascript/api/word/word.document#body)|Obtient l’objet de corps du document. Le corps du document correspond à l’ensemble du texte, à l’exception des en-têtes, des pieds de page, des notes de bas de page, des zones de texte, etc. En lecture seule.|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|Obtient la collection d’objets de contrôle de contenu dans le document. Cela inclut les contrôles de contenu dans le corps du document, les en-têtes, les pieds de page, les zones de texte, etc.. En lecture seule.|
||[conservé](/javascript/api/word/word.document#saved)|Indique si les modifications apportées au document ont été enregistrées. La valeur true indique que le document n’a pas été modifié depuis son enregistrement. En lecture seule.|
||[sections](/javascript/api/word/word.document#sections)|Obtient la collection d’objets section dans le document. En lecture seule.|
||[save()](/javascript/api/word/word.document#save--)|Enregistre le document. Cette option utilise la convention de dénomination des fichiers par défaut de Word si le document n’a jamais été enregistré précédemment.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Obtient ou définit une valeur qui indique si la police en gras. Renvoie true si la police est mise en forme en gras, sinon, false.|
||[color](/javascript/api/word/word.font#color)|Obtient ou définit la couleur de la police spécifiée. Vous pouvez fournir la valeur au format «#RRGGBB» ou au nom de la couleur.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doublestrikethrough)|Obtient ou définit une valeur qui indique si la police a un double barré. Renvoie true si la police est mise en forme en tant que texte barré double, sinon, false.|
||[highlightColor](/javascript/api/word/word.font#highlightcolor)|Obtient ou définit la couleur de surbrillance. Pour la définir, utilisez une valeur au format «#RRGGBB» ou au nom de la couleur. Pour supprimer la couleur de surbrillance, affectez-lui la valeur null. La couleur de surbrillance renvoyée peut être au format «#RRGGBB», une chaîne vide pour les couleurs de mise en surbrillance mixtes ou null pour aucune couleur de surbrillance.|
||[italic](/javascript/api/word/word.font#italic)|Obtient ou définit une valeur qui indique si la police est en italique. Renvoie true si la police est en italique, sinon, false.|
||[name](/javascript/api/word/word.font#name)|Obtient ou définit une valeur qui représente le nom de la police.|
||[size](/javascript/api/word/word.font#size)|Obtient ou définit une valeur qui représente la taille de police en points.|
||[Doubles](/javascript/api/word/word.font#strikethrough)|Obtient ou définit une valeur qui indique si la police est barrée. Renvoie true si la police est mise en forme en tant que texte barré, sinon, false.|
||[Subscript](/javascript/api/word/word.font#subscript)|Obtient ou définit une valeur qui indique si la police correspond à du texte mis en indice. Renvoie true si la police correspond à du texte mis en indice, sinon, false.|
||[superscript](/javascript/api/word/word.font#superscript)|Obtient ou définit une valeur qui indique si la police correspond à du texte en exposant. Renvoie true si la police correspond à du texte mis en exposant, sinon, false.|
||[underline](/javascript/api/word/word.font#underline)|Obtient ou définit une valeur qui indique le type de trait de soulignement de la police. «Aucun» si la police n’est pas soulignée.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|Obtient ou définit une valeur de type String qui représente le texte de remplacement associé à l’image incluse.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|Obtient ou définit une chaîne qui contient le titre de l’image incluse.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|Obtient la représentation de chaîne encodée au format Base64 de l’image incluse.|
||[height](/javascript/api/word/word.inlinepicture#height)|Obtient ou définit un nombre qui décrit la hauteur de l’image incluse.|
||[lien hypertexte](/javascript/api/word/word.inlinepicture#hyperlink)|Obtient ou définit un lien hypertexte sur l’image. Utilisez un «#» pour séparer la partie d’adresse du composant facultatif emplacement.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|Encadre l’image incluse avec un contrôle de contenu de texte enrichi.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|Obtient ou définit une valeur qui indique si l’image incluse conserve ses proportions d’origine lorsque vous la redimensionnez.|
||[ParentContentControl,](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|Obtient le contrôle de contenu qui contient l’image incluse. S’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[width](/javascript/api/word/word.inlinepicture#width)|Obtient ou définit un nombre qui décrit la largeur de l’image incluse.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Paragraph](/javascript/api/word/word.paragraph)|[aligne](/javascript/api/word/word.paragraph#alignment)|Obtient ou définit l’alignement d’un paragraphe. La valeur peut être « left » (gauche), « centered » (centré), « right » (droite) ou « justified » (justifié).|
||[clear()](/javascript/api/word/word.paragraph#clear--)|Efface le contenu de l’objet de paragraphe. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.|
||[delete()](/javascript/api/word/word.paragraph#delete--)|Supprime le paragraphe et son contenu du document.|
||[FirstLineIndent,](/javascript/api/word/word.paragraph#firstlineindent)|Renvoie ou définit la valeur, en points, du retrait de première ligne ou du retrait négatif. Utilisez une valeur positive pour définir un retrait de première ligne et une valeur négative pour définir un retrait négatif.|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Obtient une représentation HTML de l’objet Paragraph. Lorsqu’elle est affichée dans une page Web ou dans la visionneuse HTML, la mise en forme est une correspondance ferme, mais pas exacte, avec la mise en forme du document. Cette méthode ne renvoie pas exactement le même code HTML pour le même document sur différentes plateformes (Windows, Mac, etc.). Si vous avez besoin d’une fidélité exacte ou d’une cohérence `Paragraph.getOoxml()` entre les plateformes, utilisez et convertissez le code XML renvoyé en html.|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Obtient la représentation Office Open XML (OOXML) de l’objet de paragraphe.|
||[Ignorepunct,](/javascript/api/word/word.paragraph#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|Insère un saut à l’emplacement spécifié du document principal. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|Encadre l’objet de paragraphe avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64 (base64File: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|Insère un document dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertHtml (HTML: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|Insère du code HTML dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertInlinePictureFromBase64 (base64EncodedImage: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insère une image dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertOoxml (OOXML: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|Insère OOXML dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[insertParagraph (paragraphText: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
||[insertText (Text: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|Insère du texte dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
||[leftIndent](/javascript/api/word/word.paragraph#leftindent)|Obtient ou définit la valeur de retrait à gauche, en points, pour le paragraphe.|
||[lineSpacing](/javascript/api/word/word.paragraph#linespacing)|Obtient ou définit l’interligne, en points, pour le paragraphe spécifié. Dans l’interface utilisateur de Word, cette valeur est divisée par 12.|
||[LineUnitAfter,](/javascript/api/word/word.paragraph#lineunitafter)|Obtient ou définit la quantité d’espace, dans le quadrillage, après le paragraphe.|
||[LineUnitBefore,](/javascript/api/word/word.paragraph#lineunitbefore)|Obtient ou définit la quantité d’espace, en lignes de quadrillage, avant le paragraphe.|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchprefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchwildcards)||
||[OutlineLevel,](/javascript/api/word/word.paragraph#outlinelevel)|Obtient ou définit le niveau hiérarchique pour le paragraphe.|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|Obtient la collection d’objets de contrôle de contenu dans le paragraphe. En lecture seule.|
||[police](/javascript/api/word/word.paragraph#font)|Obtient le format de texte du paragraphe. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|Obtient la collection d’objets InlinePicture dans le paragraphe. La collection n’inclut pas d’images flottantes. En lecture seule.|
||[ParentContentControl,](/javascript/api/word/word.paragraph#parentcontentcontrol)|Obtient le contrôle de contenu qui contient le paragraphe. S’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[text](/javascript/api/word/word.paragraph#text)|Obtient le texte du paragraphe. En lecture seule.|
||[rightIndent](/javascript/api/word/word.paragraph#rightindent)|Obtient ou définit la valeur de retrait à droite, en points, pour le paragraphe.|
||[recherche (Texted’origine: chaîne, searchOptions?: Word. SearchOptions})](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié sur l’étendue de l’objet Paragraph. Les résultats de la recherche sont un ensemble d’objets de plage.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|Sélectionne le paragraphe et y accède via l’interface utilisateur de Word.|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceafter)|Obtient ou définit l’espacement, en points, après le paragraphe.|
||[spaceBefore](/javascript/api/word/word.paragraph#spacebefore)|Obtient ou définit l’espacement, en points, avant le paragraphe.|
||[style](/javascript/api/word/word.paragraph#style)|Obtient ou définit le nom du style du paragraphe. Utilisez cette propriété pour les noms des styles personnalisés et localisés. Pour utiliser les styles prédéfinis qui sont portables entre différents paramètres régionaux, voir la propriété « styleBuiltIn ».|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|Efface le contenu de l’objet de plage. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.|
||[delete()](/javascript/api/word/word.range#delete--)|Supprime la plage et son contenu du document.|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Obtient une représentation HTML de l’objet de plage. Lorsqu’elle est affichée dans une page Web ou dans la visionneuse HTML, la mise en forme est une correspondance ferme, mais pas exacte, avec la mise en forme du document. Cette méthode ne renvoie pas exactement le même code HTML pour le même document sur différentes plateformes (Windows, Mac, etc.). Si vous avez besoin d’une fidélité exacte ou d’une cohérence `Range.getOoxml()` entre les plateformes, utilisez et convertissez le code XML renvoyé en html.|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Obtient la représentation OOXML de l’objet de plage.|
||[Ignorepunct,](/javascript/api/word/word.range#ignorepunct)||
||[ignorespace,](/javascript/api/word/word.range#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|Insère un saut à l’emplacement spécifié du document principal. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|Encadre l’objet de plage avec un contrôle de contenu de texte enrichi.|
||[insertFileFromBase64 (base64File: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|Insère un document à l’emplacement spécifié. La valeur insertLocation peut être «Replace», «Start», «end», «Before» ou «after».|
||[insertHtml (HTML: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|Insère du code HTML à l’emplacement spécifié. La valeur insertLocation peut être «Replace», «Start», «end», «Before» ou «after».|
||[insertOoxml (OOXML: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|Insère du code OOXML à l’emplacement spécifié.  La valeur insertLocation peut être «Replace», «Start», «end», «Before» ou «after».|
||[insertParagraph (paragraphText: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
||[insertText (Text: chaîne, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|Insère du texte à l’emplacement spécifié. La valeur insertLocation peut être «Replace», «Start», «end», «Before» ou «after».|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[matchPrefix](/javascript/api/word/word.range#matchprefix)||
||[matchSuffix](/javascript/api/word/word.range#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|Obtient la collection d’objets de contrôle de contenu dans la plage. En lecture seule.|
||[police](/javascript/api/word/word.range#font)|Obtient le format de texte de la plage. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Obtient la collection d’objets Paragraph dans la plage. En lecture seule.|
||[ParentContentControl,](/javascript/api/word/word.range#parentcontentcontrol)|Obtient le contrôle de contenu qui contient la plage. S’il n’existe pas de contrôle de contenu parent. En lecture seule.|
||[text](/javascript/api/word/word.range#text)|Obtient le texte de la plage. En lecture seule.|
||[recherche (Texted’origine: chaîne, searchOptions?: Word. SearchOptions)](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Effectue une recherche avec le SearchOptions spécifié sur l’étendue de l’objet Range. Les résultats de la recherche sont un ensemble d’objets de plage.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|Sélectionne la plage et y accède via l’interface utilisateur de Word.|
||[style](/javascript/api/word/word.range#style)|Obtient ou définit le nom du style de la plage. Utilisez cette propriété pour les noms des styles personnalisés et localisés. Pour utiliser les styles prédéfinis qui sont portables entre différents paramètres régionaux, voir la propriété « styleBuiltIn ».|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[Ignorepunct,](/javascript/api/word/word.searchoptions#ignorepunct)|Obtient ou définit une valeur indiquant si toutes les marques de ponctuation entre les mots doivent être ignorées. Correspond à la case à cocher Ignorer les marques de ponctuation de la boîte de dialogue Rechercher et remplacer.|
||[ignorespace,](/javascript/api/word/word.searchoptions#ignorespace)|Obtient ou définit une valeur qui indique s’il faut ignorer tous les espaces entre les mots. Correspond à la case à cocher Ignorer les caractères d’espacement dans la boîte de dialogue Rechercher et remplacer.|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|Obtient ou définit une valeur indiquant si la recherche respecte la casse. Correspond à la case à cocher respecter la casse dans la boîte de dialogue Rechercher et remplacer.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchprefix)|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui commencent par la chaîne entrée. Correspond à la case à cocher Préfixe de la boîte de dialogue Rechercher et remplacer.|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchsuffix)|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui se terminent par la chaîne entrée. Correspond à la case à cocher Suffixe de la boîte de dialogue Rechercher et remplacer.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchwholeword)|Obtient ou définit une valeur indiquant si la recherche doit uniquement porter sur des mots entiers et exclure le texte s’il est inclus dans un mot plus long. Correspond à la case à cocher Mot entier de la boîte de dialogue Rechercher et remplacer.|
||[matchWildCards](/javascript/api/word/word.searchoptions#matchwildcards)||
||[matchWildcards](/javascript/api/word/word.searchoptions#matchwildcards)|Obtient ou définit une valeur indiquant si la recherche est effectuée à l’aide d’opérateurs de recherche spéciaux. Correspond à la case Caractères génériques de la boîte de dialogue Rechercher et remplacer.|
|[Section](/javascript/api/word/word.section)|[getFooter (type: Word. HeaderFooterType)](/javascript/api/word/word.section#getfooter-type-)|Obtient l’un des pieds de page de la section.|
||[getHeader (type: Word. HeaderFooterType)](/javascript/api/word/word.section#getheader-type-)|Obtient l’un des en-têtes de la section.|
||[body](/javascript/api/word/word.section#body)|Obtient l’objet Body de la section. Cela n’inclut pas l’en-tête/pied de page et d’autres métadonnées de section. En lecture seule.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
