---
title: PowerPoint l’ensemble de conditions requises de l’API JavaScript 1.3
description: Détails sur l’ensemble de conditions requises PowerPointApi 1.3.
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
---

# <a name="whats-new-in-powerpoint-javascript-api-13"></a>Nouveautés de l PowerPoint API JavaScript 1.3

PowerPointApi 1.3 a ajouté une prise en charge supplémentaire pour la gestion des diapositives et le marquage personnalisé.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Gestion des diapositives](../../powerpoint/add-slides.md) | Ajoute la prise en charge de l’ajout de diapositives, ainsi que de la gestion des mises en page des diapositives et des formes de base des diapositives. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| [Tags](../../powerpoint/tagging-presentations-slides-shapes.md) | Permet aux add-ins d’attacher des métadonnées personnalisées, sous la forme de paires clé-valeur. | [Tag](/javascript/api/powerpoint/powerpoint.tag) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie PowerPoint l’ensemble de conditions requises de l’API JavaScript 1.3. Pour obtenir la liste complète de toutes les API JavaScript PowerPoint (y compris les API de prévisualisation et les API publiées précédemment), consultez toutes les API [JavaScript PowerPoint de prévisualisation](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Classe | Champs | Description |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-layoutid-member)|Spécifie l’ID d’une mise en page des diapositives à utiliser pour la nouvelle diapositive.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-slidemasterid-member)|Spécifie l’ID d’un curseur de diapositive à utiliser pour la nouvelle diapositive.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slidemasters-member)|Renvoie la collection d’objets `SlideMaster` qui se retrouvent dans la présentation.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-tags-member)|Renvoie une collection de balises attachées à la présentation.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-delete-member(1))|Supprime la forme de la collection de formes.|
||[id](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-id-member)|Obtient l’ID unique de la forme.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-tags-member)|Renvoie une collection de balises dans la forme.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getcount-member(1))|Obtient le nombre de formes dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitem-member(1))|Obtient une forme à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemat-member(1))|Obtient une forme à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemornullobject-member(1))|Obtient une forme à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[disposition](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-layout-member)|Obtient la mise en page de la diapositive.|
||[Formes](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-shapes-member)|Renvoie une collection de formes dans la diapositive.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-slidemaster-member)|Obtient `SlideMaster` l’objet qui représente le contenu par défaut de la diapositive.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-tags-member)|Renvoie une collection de balises dans la diapositive.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1))|Ajoute une nouvelle diapositive à la fin de la collection.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-id-member)|Obtient l’ID unique de la mise en page des diapositives.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-name-member)|Obtient le nom de la mise en page des diapositives.|
||[Formes](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-shapes-member)|Renvoie une collection de formes dans la mise en page des diapositives.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getcount-member(1))|Obtient le nombre de dispositions dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitem-member(1))|Obtient une disposition à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemat-member(1))|Obtient une disposition à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemornullobject-member(1))|Obtient une disposition à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-id-member)|Obtient l’ID unique du curseur de diapositive.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-layouts-member)|Obtient la collection de mises en page fournies par le maître des diapositives pour les diapositives.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-name-member)|Obtient le nom unique du curseur de diapositive.|
||[Formes](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-shapes-member)|Renvoie une collection de formes dans le curseur de diapositive.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getcount-member(1))|Obtient le nombre de diapositives de la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitem-member(1))|Obtient un curseur de diapositive à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemat-member(1))|Obtient un curseur de diapositive à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemornullobject-member(1))|Obtient un curseur de diapositive à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-key-member)|Obtient l’ID unique de la balise.|
||[value](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-value-member)|Obtient la valeur de la balise.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1))|Ajoute une nouvelle balise à la fin de la collection.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-delete-member(1))|Supprime la balise avec la balise donnée `key` dans cette collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getcount-member(1))|Obtient le nombre de balises dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitem-member(1))|Obtient une balise à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemat-member(1))|Obtient une balise à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemornullobject-member(1))|Obtient une balise à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [PowerPoint documentation de référence de l’API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.3&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](powerpoint-api-requirement-sets.md)
