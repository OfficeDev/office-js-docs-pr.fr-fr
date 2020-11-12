---
title: API d’aperçu JavaScript pour PowerPoint
description: Informations détaillées sur les API JavaScript JavaScript à venir.
ms.date: 11/09/2020
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: b53b6638b16b2028342003b9a77aa59e7406d5f3
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996521"
---
# <a name="powerpoint-javascript-preview-apis"></a>API d’aperçu JavaScript pour PowerPoint

Les nouvelles API JavaScript pour PowerPoint sont tout d’abord introduites dans « Preview » et par la suite, elles deviennent une partie d’un ensemble de conditions requises spécifiques, après un test suffisant, et les commentaires des utilisateurs sont acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Insérer et supprimer des diapositives | Permet l’insertion de diapositives existantes dans la présentation active à partir d’une autre présentation, ainsi que la possibilité de supprimer sildes. | [Slide. Delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation. insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour PowerPoint actuellement en version préliminaire. Pour obtenir la liste complète des API JavaScript pour PowerPoint (dont les API d’aperçu et les API précédemment publiées), voir [toutes les API JavaScript pour PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[mise en forme](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Indique la mise en forme à utiliser lors de l’insertion d’une diapositive.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Cette énumération spécifie les diapositives de la présentation source qui seront insérées dans la présentation en cours.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Indique où seront insérées les nouvelles diapositives dans la présentation.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (base64File : chaîne, Options ?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Insère les diapositives spécifiées à partir d’une présentation dans la présentation active.|
||[celles](/javascript/api/powerpoint/powerpoint.presentation#slides)|Renvoie une collection ordonnée de diapositives dans la présentation.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Supprime la diapositive de la présentation.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Obtient l’ID unique de la diapositive.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Obtient le nombre de diapositives de la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Obtient une diapositive à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Obtient une diapositive à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Obtient une diapositive à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](powerpoint-api-requirement-sets.md)
