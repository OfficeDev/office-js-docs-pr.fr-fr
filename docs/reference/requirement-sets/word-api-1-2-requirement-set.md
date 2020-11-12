---
title: Ensemble de conditions requises de l’API JavaScript pour Word 1,2
description: Détails sur l’ensemble de conditions requises WordApi 1,2
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: ee9bf60a3a944a3a01a2ca5aa10d01958e3d3475
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996423"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Nouveautés de l’API JavaScript 1.2 pour Word

WordApi 1,2 Ajout de la prise en charge des images insérées.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Word 1,2. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’API JavaScript pour Word, ensemble de conditions requises 1,2 ou antérieure, voir [API Word dans l’ensemble de conditions requises 1,2 ou version antérieure](/javascript/api/word?view=word-js-1.2&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[Corps](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.body#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insère une image dans le corps à l’emplacement spécifié.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insère une image incluse dans le contrôle de contenu, à l’emplacement spécifié.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete--)|Supprime l’image insérée du document.|
||[insertBreak (breakType : Word. BreakType, insertLocation : Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertbreak-breaktype--insertlocation-)|Insère un saut à l’emplacement spécifié du document principal.|
||[insertFileFromBase64 (base64File : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertfilefrombase64-base64file--insertlocation-)|Insère un document à l’emplacement spécifié.|
||[insertHtml (HTML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserthtml-html--insertlocation-)|Insère du code HTML à l’emplacement spécifié.|
||[insertInlinePictureFromBase64 (base64EncodedImage : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insère une image insérée à l’emplacement spécifié.|
||[insertOoxml (OOXML : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertooxml-ooxml--insertlocation-)|Insère du code OOXML à l’emplacement spécifié.|
||[insertParagraph (paragraphText : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertparagraph-paragraphtext--insertlocation-)|Insère un paragraphe à l’emplacement spécifié.|
||[insertText (Text : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserttext-text--insertlocation-)|Insère du texte à l’emplacement spécifié.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Obtient le paragraphe parent qui contient l’image insérée.|
||[Select (selectionMode ?: Word. SelectionMode)](/javascript/api/word/word.inlinepicture#select-selectionmode-)|Sélectionne l’image insérée.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64 (base64EncodedImage : chaîne, insertLocation : Word. InsertLocation)](/javascript/api/word/word.range#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insère une image à l’emplacement spécifié.|
||[inlinePictures](/javascript/api/word/word.range#inlinepictures)|Obtient la collection d’objets image insérée de la plage.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
