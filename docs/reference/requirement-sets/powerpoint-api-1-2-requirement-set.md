---
title: Ensemble de conditions requises de l’API JavaScript pour PowerPoint 1.2
description: Détails sur l’ensemble de conditions requises PowerPointApi 1.2.
ms.date: 01/08/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 0f6d1e766de81fef5d071152f6116ab56613ec9d
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49841531"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>Nouveautés de l’API JavaScript 1.2 pour PowerPoint

PowerPointApi 1.2 a ajouté la prise en charge de l’insertion de diapositives d’une autre présentation dans la présentation actuelle et de la suppression des diapositives.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Insérer et supprimer des diapositives | Permet l’insertion de diapositives existantes dans la présentation actuelle à partir d’une autre présentation, ainsi que la possibilité de supprimer des diapositives. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie l’ensemble de conditions requises de l’API JavaScript pour PowerPoint 1.2. Pour obtenir la liste complète de toutes les API JavaScript PowerPoint (y compris les API de prévisualisation et les API publiées précédemment), voir toutes les API [JavaScript pour PowerPoint.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[mise en forme](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Spécifie la mise en forme à utiliser lors de l’insertion d’une diapositive.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Spécifie les diapositives de la présentation source qui seront insérées dans la présentation actuelle.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Spécifie l’endroit où seront insérées les nouvelles diapositives dans la présentation.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Insère les diapositives spécifiées d’une présentation dans la présentation actuelle.|
||[diapositives](/javascript/api/powerpoint/powerpoint.presentation#slides)|Renvoie une collection ordonnée de diapositives dans la présentation.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Supprime la diapositive de la présentation.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Obtient l’ID unique de la diapositive.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Obtient le nombre de diapositives de la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Obtient une diapositive à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Obtient une diapositive à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Obtient une diapositive à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](powerpoint-api-requirement-sets.md)
