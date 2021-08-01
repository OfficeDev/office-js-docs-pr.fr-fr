---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1.2
description: Détails sur l’ensemble de conditions requises WordApi 1.2
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: fd33b043a9205e793a248c35118ed86efcdf0036
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671855"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Nouveautés de l’API JavaScript 1.2 pour Word

WordApi 1.2 a ajouté la prise en charge des images en ligne.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de l’ensemble de conditions requises de l’API JavaScript pour Word 1.2. Pour afficher la documentation de référence de l’API pour toutes les API pris en charge par l’ensemble de conditions requises de l’API JavaScript pour Word 1.2 ou une version antérieure, voir API Word dans l’ensemble de conditions requises [1.2](/javascript/api/word?view=word-js-1.2&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Corps](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insère une image dans le corps à l’emplacement spécifié.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insère une image incluse dans le contrôle de contenu, à l’emplacement spécifié.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete__)|Supprime l’image insérée du document.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertBreak_breakType__insertLocation_)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertFileFromBase64_base64File__insertLocation_)|Insère un document à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertHtml_html__insertLocation_)|Insère du code HTML à l’emplacement spécifié.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insère une image insérée à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertOoxml_ooxml__insertLocation_)|Insère du code OOXML à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertParagraph_paragraphText__insertLocation_)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertText_text__insertLocation_)|Insère du texte à l’emplacement spécifié.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Obtient le paragraphe parent qui contient l’image insérée.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#select_selectionMode_)|Sélectionne l’image insérée.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insère une image à l’emplacement spécifié.|
||[inlinePictures](/javascript/api/word/word.range#inlinePictures)|Obtient la collection d’objets image insérée de la plage.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
