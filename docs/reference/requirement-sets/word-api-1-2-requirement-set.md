---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1.2
description: Détails sur l’ensemble de conditions requises WordApi 1.2.
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: 1a5af83615786b241c43ecb07ee0d23b3758cfc8
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744218"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Nouveautés de l’API JavaScript 1.2 pour Word

WordApi 1.2 a ajouté la prise en charge des images en ligne.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de l’ensemble de conditions requises de l’API JavaScript pour Word 1.2. Pour afficher la documentation de référence de l’API pour toutes les API pris en charge par l’ensemble de conditions requises de l’API JavaScript pour Word 1.2 ou une version antérieure, voir API Word dans l’ensemble de conditions requises [1.2](/javascript/api/word?view=word-js-1.2&preserve-view=true) ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Corps](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertinlinepicturefrombase64-member(1))|Insère une image dans le corps à l’emplacement spécifié.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertinlinepicturefrombase64-member(1))|Insère une image incluse dans le contrôle de contenu, à l’emplacement spécifié.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-delete-member(1))|Supprime l’image insérée du document.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertbreak-member(1))|Insère un saut à l’emplacement spécifié du document principal.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertfilefrombase64-member(1))|Insère un document à l’emplacement spécifié.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserthtml-member(1))|Insère du code HTML à l’emplacement spécifié.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertinlinepicturefrombase64-member(1))|Insère une image insérée à l’emplacement spécifié.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertooxml-member(1))|Insère du code OOXML à l’emplacement spécifié.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertparagraph-member(1))|Insère un paragraphe à l’emplacement spécifié.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserttext-member(1))|Insère du texte à l’emplacement spécifié.|
||[paragraph](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-paragraph-member)|Obtient le paragraphe parent qui contient l’image insérée.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-select-member(1))|Sélectionne l’image insérée.|
|[Range](/javascript/api/word/word.range)|[inlinePictures](/javascript/api/word/word.range#word-word-range-inlinepictures-member)|Obtient la collection d’objets image insérée de la plage.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertinlinepicturefrombase64-member(1))|Insère une image à l’emplacement spécifié.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
