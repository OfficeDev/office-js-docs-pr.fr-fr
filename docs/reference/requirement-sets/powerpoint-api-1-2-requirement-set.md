---
title: PowerPoint Ensemble de conditions requises de l’API JavaScript 1.2
description: Détails sur l’ensemble de conditions requises PowerPointApi 1.2.
ms.date: 01/27/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: fac472e9b88b78f52fe939f883d88cded8b1702c
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671610"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>Nouveautés de l PowerPoint API JavaScript 1.2

PowerPointApi 1.2 a ajouté la prise en charge de l’insertion de diapositives d’une autre présentation dans la présentation actuelle et de la suppression des diapositives.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Insérer et supprimer des diapositives](../../powerpoint/insert-slides-into-presentation.md) | Permet l’insertion de diapositives existantes dans la présentation actuelle à partir d’une autre présentation, ainsi que la possibilité de supprimer des diapositives. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les PowerPoint conditions requises de l’API JavaScript 1.2. Pour obtenir la liste complète de toutes les API JavaScript PowerPoint (y compris les API de prévisualisation et les API publiées précédemment), voir toutes les API [JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)PowerPoint.

| Classe | Champs | Description |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[mise en forme](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Spécifie la mise en forme à utiliser lors de l’insertion d’une diapositive.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceSlideIds)|Spécifie les diapositives de la présentation source qui seront insérées dans la présentation actuelle.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetSlideId)|Spécifie l’endroit où seront insérées les nouvelles diapositives dans la présentation.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_)|Insère les diapositives spécifiées d’une présentation dans la présentation actuelle.|
||[diapositives](/javascript/api/powerpoint/powerpoint.presentation#slides)|Renvoie une collection ordonnée de diapositives dans la présentation.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete__)|Supprime la diapositive de la présentation.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Obtient l’ID unique de la diapositive.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getCount__)|Obtient le nombre de diapositives de la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItem_key_)|Obtient une diapositive à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)|Obtient une diapositive à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemOrNullObject_id_)|Obtient une diapositive à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [PowerPoint Documentation de référence de l’API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](powerpoint-api-requirement-sets.md)
