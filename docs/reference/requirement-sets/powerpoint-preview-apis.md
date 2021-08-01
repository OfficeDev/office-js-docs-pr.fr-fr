---
title: PowerPoint API d’aperçu JavaScript
description: Détails sur les API JavaScript PowerPoint à venir.
ms.date: 01/27/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: af947919ad680864bf4a63ab29af33d0560aaaa0
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671603"
---
# <a name="powerpoint-javascript-preview-apis"></a>PowerPoint API d’aperçu JavaScript

Les nouvelles API JavaScript PowerPoint sont d’abord introduites dans « aperçu », puis font partie d’un ensemble de conditions requises numérotées spécifiques une fois que des tests suffisants ont eu lieu et que les commentaires des utilisateurs ont été acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Gestion des diapositives | Ajoute la prise en charge de l’ajout de diapositives, ainsi que de la gestion des mises en page des diapositives et des formes de base des diapositives. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| Formes | Ajoute la prise en charge de l’obtention de références aux formes dans une diapositive. | [Forme](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les POWERPOINT JavaScript actuellement en prévisualisation. Pour obtenir la liste complète de toutes les API JavaScript PowerPoint (y compris les API de prévisualisation et les API publiées précédemment), voir toutes Excel API [JavaScript.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|Spécifie l’ID d’une mise en page des diapositives à utiliser pour la nouvelle diapositive.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|Spécifie l’ID d’un curseur de diapositive à utiliser pour la nouvelle diapositive.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|Renvoie la collection `SlideMaster` d’objets qui se retrouvent dans la présentation.|
||[étiquettes](/javascript/api/powerpoint/powerpoint.presentation#tags)|Renvoie une collection de balises attachées à la présentation.|
|[Forme](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|Supprime la forme de la collection de formes.|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|Obtient l’ID unique de la forme.|
||[étiquettes](/javascript/api/powerpoint/powerpoint.shape#tags)|Renvoie une collection de balises dans la forme.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|Obtient le nombre de formes dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|Obtient une forme à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|Obtient une forme à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|Obtient une forme à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[disposition](/javascript/api/powerpoint/powerpoint.slide#layout)|Obtient la mise en page de la diapositive.|
||[Formes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Renvoie une collection de formes dans la diapositive.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|Obtient `SlideMaster` l’objet qui représente le contenu par défaut de la diapositive.|
||[étiquettes](/javascript/api/powerpoint/powerpoint.slide#tags)|Renvoie une collection de balises dans la diapositive.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|Ajoute une nouvelle diapositive à la fin de la collection.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Obtient l’ID unique de la mise en page des diapositives.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Obtient le nom de la mise en page des diapositives.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|Obtient le nombre de dispositions dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|Obtient une disposition à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|Obtient une disposition à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|Obtient une disposition à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Obtient l’ID unique du curseur de diapositive.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Obtient la collection de mises en page fournies par le maître des diapositives pour les diapositives.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Obtient le nom unique du curseur de diapositive.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|Obtient le nombre de diapositives de la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|Obtient un curseur de diapositive à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|Obtient un curseur de diapositive à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|Obtient un curseur de diapositive à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Obtient l’ID unique de la balise.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Obtient la valeur de la balise.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|Ajoute une nouvelle balise à la fin de la collection.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|Supprime la balise avec la balise donnée `key` dans cette collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|Obtient le nombre de balises dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|Obtient une balise à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|Obtient une balise à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|Obtient une balise à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [PowerPoint Documentation de référence de l’API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](powerpoint-api-requirement-sets.md)