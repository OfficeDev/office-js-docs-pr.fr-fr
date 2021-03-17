---
title: API d’aperçu JavaScript Pour PowerPoint
description: Détails sur les API JavaScript PowerPoint à venir.
ms.date: 01/27/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 042ce0c2b42b2c0dca9900982376cd568a4a3622
ms.sourcegitcommit: 929dcf2f415b94f42330a9035ed11a5cedad88f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/16/2021
ms.locfileid: "50830971"
---
# <a name="powerpoint-javascript-preview-apis"></a>API d’aperçu JavaScript Pour PowerPoint

Les nouvelles API JavaScript Pour PowerPoint sont d’abord introduites dans « aperçu », puis font partie d’un ensemble de conditions requises numérotées spécifiques une fois que des tests suffisants ont été effectués et que les commentaires des utilisateurs ont été acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Gestion des diapositives | Ajoute la prise en charge de l’ajout de diapositives, ainsi que de la gestion des mises en page des diapositives et des formes de base des diapositives. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| Formes | Ajoute la prise en charge de l’obtention de références aux formes dans une diapositive. | [Forme](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript PowerPoint actuellement en prévisualisation. Pour obtenir la liste complète de toutes les API JavaScript Pour PowerPoint (y compris les API de prévisualisation et les API publiées précédemment), voir toutes les API [JavaScript pour Excel.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|Spécifie l’ID d’une mise en page des diapositives à utiliser pour la nouvelle diapositive.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|Spécifie l’ID d’un curseur de diapositive à utiliser pour la nouvelle diapositive.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|Renvoie la collection `SlideMaster` d’objets qui se retrouvent dans la présentation.|
||[étiquettes](/javascript/api/powerpoint/powerpoint.presentation#tags)|Renvoie une collection de balises attachées à la présentation.|
|[Forme](/javascript/api/powerpoint/powerpoint.shape)|[id](/javascript/api/powerpoint/powerpoint.shape#id)|Obtient l’ID unique de la forme.|
||[étiquettes](/javascript/api/powerpoint/powerpoint.shape#tags)|Renvoie une collection de balises dans la forme.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|Obtient le nombre de formes dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|Obtient une forme à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|Obtient une forme à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|Obtient une forme à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[disposition](/javascript/api/powerpoint/powerpoint.slide#layout)|Obtient la mise en page de la diapositive.|
||[Formes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Renvoie une collection de formes dans la diapositive.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|Obtient `SlideMaster` l’objet qui représente le contenu par défaut de la diapositive.|
||[étiquettes](/javascript/api/powerpoint/powerpoint.slide#tags)|Renvoie une collection de balises dans la diapositive.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|Ajoute une nouvelle diapositive à la fin de la collection.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Obtient l’ID unique de la mise en page des diapositives.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Obtient le nom de la mise en page des diapositives.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|Obtient le nombre de dispositions dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|Obtient une disposition à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|Obtient une disposition à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|Obtient une disposition à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Obtient l’ID unique du curseur de diapositive.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Obtient la collection de mises en page fournies par le maître des diapositives pour les diapositives.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Obtient le nom unique du curseur de diapositive.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|Obtient le nombre de diapositives de la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|Obtient un curseur de diapositive à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|Obtient un curseur de diapositive à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|Obtient un curseur de diapositive à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Obtient l’ID unique de la balise.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Obtient la valeur de la balise.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add-key--value-)|Ajoute une nouvelle balise à la fin de la collection.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete-key-)|Supprime la balise avec la balise donnée `key` dans cette collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getcount--)|Obtient le nombre de balises dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitem-key-)|Obtient une balise à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemat-index-)|Obtient une balise à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemornullobject-key-)|Obtient une balise à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](powerpoint-api-requirement-sets.md)