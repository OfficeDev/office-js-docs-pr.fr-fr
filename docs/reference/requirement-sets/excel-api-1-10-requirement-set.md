---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1.10
description: Détails sur l’ensemble de conditions requises ExcelApi 1.10.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1bafdd2064166019c5c3f22aa4da1a2d0ec73f08
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650820"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Nouveautés de l’API JavaScript 1.10 pour Excel

ExcelApi 1.10 a introduit des fonctionnalités clés, telles que les commentaires, les contours et les slicers. Il a également ajouté la prise en charge des événements pour le clic et le tri au niveau de la feuille de calcul.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Commentaires](../../excel/excel-add-ins-comments.md) | Ajouter, modifier et supprimer des listes. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Outlines](../../excel/excel-add-ins-ranges-group.md) | Groupez les lignes et les colonnes pour former des contours réductibles. | [Range](/javascript/api/excel/excel.range), [Worksheet](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | Insérer et configurez segments aux tableaux et tableaux croisés dynamiques. | [Segment](/javascript/api/excel/excel.slicer) |
| [Autres événements de feuille de calcul](../../excel/excel-add-ins-events.md) | Écouter les événements de clic et de tri dans la feuille de calcul. | [Feuille de calcul (événements)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de l’ensemble de conditions requises de l’API JavaScript pour Excel 1.10. Pour afficher la documentation de référence de l’API pour toutes les API pris en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1.10 ou une version antérieure, voir API Excel dans l’ensemble de conditions requises [1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Contenu du commentaire.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Supprime le commentaire et toutes les réponses connectées.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Obtient la cellule où se trouve ce commentaire.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Obtient le nom de l’auteur du commentaire.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Obtenir l’heure de création du commentaire.|
||[id](/javascript/api/excel/excel.comment#id)|Spécifie l’identificateur de commentaire.|
||[Réponses](/javascript/api/excel/excel.comment#replies)|Représente une collection de feuilles de calcul associées au classeur.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtient le nombre de commentaires de la collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Obtient un commentaire à partir de la collection de sites en fonction de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtient un commentaire en fonction de sa position dans la collection.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtient le commentaire à partir de la cellule spécifiée.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtient le commentaire auquel la réponse donnée est connectée.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Contenu de la réponse au commentaire.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Supprime la réponse de commentaire.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Obtient la cellule où se trouve cette réponse de commentaire.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtient le commentaire parent de cette réponse.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Obtient l’heure de création de la réponse à un commentaire.|
||[id](/javascript/api/excel/excel.commentreply#id)|Spécifie l’identificateur de réponse du commentaire.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse à un commentaire pour un commentaire.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtient le nombre de réponses aux commentaires de la collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Renvoie une réponse de commentaire identifié via son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtient une réponse de commentaire en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Supprime le tableau croisé dynamique.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Crée un doublon de cette PivotTableStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtient le nom du PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Spécifie si cet objet PivotTableStyle est accessible en lecture seule.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Crée un PivotTableStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Obtient le nombre de styles tableaux croisés dynamiques de la collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Obtient PivotTableStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Extrait un tableau croisé dynamique par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Extrait un tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Définit le PivotTableStyle par défaut pour la portée de l’objet parent portée.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|Groupe les colonnes et les lignes d’un plan.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Masquer les détails du groupe de lignes ou de colonnes.|
||[height](/javascript/api/excel/excel.range#height)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la plage au bord inférieur de la plage.|
||[left](/javascript/api/excel/excel.range#left)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la feuille de calcul au bord gauche de la plage.|
||[top](/javascript/api/excel/excel.range#top)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la feuille de calcul au bord supérieur de la plage.|
||[width](/javascript/api/excel/excel.range#width)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la plage au bord droit de la plage.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Afficher les détails du groupe de lignes ou de colonnes.|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Regroupe les colonnes et les lignes d’un plan.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copie et colle un objet de forme.|
||[placement](/javascript/api/excel/excel.shape#placement)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
|[Segment](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Représente la légende de segment.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Supprime tous les filtres appliqués actuellement sur le tableau.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Supprime le segment.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Renvoie une matrice de noms d’éléments sélectionnés.|
||[height](/javascript/api/excel/excel.slicer#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[left](/javascript/api/excel/excel.slicer#left)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicer#name)|Représente le nom de la forme.|
||[id](/javascript/api/excel/excel.slicer#id)|Représente l’id unique du segment.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|True si tous les filtres appliqués actuellement sur le segment sont effacés.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Représente la collection de SlicerItems qui font partie du segment.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Obtenir la feuille de calcul contenant la plage.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Sélectionne les éléments de slicer en fonction de leurs clés.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Représente l’ordre de tri des éléments dans le segment.|
||[style](/javascript/api/excel/excel.slicer#style)|Valeur de constante qui représente le style du tableau.|
||[top](/javascript/api/excel/excel.slicer#top)|Représente la distance, en points, du bord supérieur de la section à la partie droite de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicer#width)|Représente la largeur, en points, de la forme.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[Ajouter (slicerSource : chaîne \| tableau croisé dynamique \| Table, sourceField : chaîne \| PivotField \| nombre \| TableColumn, slicerDestination ? : chaîne \| feuille de calcul)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Ajoute un nouveau segment au classeur.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Renvoie le nombre de séries de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Obtient un slicer à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True si l’élément de slicer est sélectionné ; sinon False.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True si l’élément de segment comporte des données.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Représente la valeur unique représentant l’élément de segment.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Représente le titre affiché dans l’interface utilisateur.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Renvoie le nombre de segment les éléments dans le segment.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Obtient un segment objet de l’élément à l’aide de son nom ou clé.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Obtient un segment de l’élément à l’aide de son nom ou clé.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Supprime le SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Crée un doublon de cette SlicerStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtient le nom de la SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Spécifie si cet objet SlicerStyle est accessible en lecture seule.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Crée un SlicerStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Obtient le nombre de styles de slicer de la collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Obtient SlicerStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Obtient un SlicerStyle par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Obtient un SlicerStyle par nom.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Obtient le SlicerStyle par défaut pour la portée de l’objet parent portée.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Supprime le TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Crée un doublon de cette TableStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtient le nom du TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Spécifie si cet objet TableStyle est accessible en lecture seule.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Crée un TableStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Obtient le nombre de styles de tableaux de la collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Obtient le TableStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Extrait un TableStyle par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Extrait un TableStyle par nom.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Définit le TableStyle par défaut pour la portée de l’objet parent portée.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Supprime le TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Crée un doublon de cette TimelineStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtient le nom du TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Spécifie si cet objet TimelineStyle est accessible en lecture seule.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Crée un TimelineStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Obtient le nombre de styles de délai de la collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Obtient le TimelineStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Obtient un TimelineStyle par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Obtient un TimelineStyle par nom.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Définit le TimelineStyle par défaut pour la portée de l’objet parent portée.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtient le segment actif actuel du classeur.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtient le segment actif actuel du classeur.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Représente une collection de styles associés au classeur.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Représente une collection de PivotTableStyles associée au classeur.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Représente une collection de styles associés au classeur.|
||[Slicers](/javascript/api/excel/excel.workbook#slicers)|Représente une collection de styles associés au classeur.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Représente une collection de TableStyles associés au classeur.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Représente une collection de TimelineStyles associés au classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Renvoie une collection de tous les objets Lecteur sur l’ordinateur.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Se produit lorsqu’une ou plusieurs colonnes ont été triées.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Se produit lorsqu’une ou plusieurs lignes ont été triées.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Se produit lorsqu’une action clic gauche/clic se produit dans la feuille de calcul.|
||[Slicers](/javascript/api/excel/excel.worksheet#slicers)|Renvoie une collection de slicers qui font partie de la feuille de calcul.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Affiche les groupes de lignes ou de colonnes par niveaux de plan.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Se produit lorsqu’une ou plusieurs colonnes ont été triées.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Se produit lorsqu’une ou plusieurs lignes ont été triées.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Se produit lorsque l’opération clic gauche/clic se produit dans la collection de feuilles de calcul.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le tri s’est passé.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le tri s’est passé.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[adresse](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtient l’adresse qui représente la cellule sur laquelle vous avez fait un clic gauche/appuyé pour une feuille de calcul spécifique.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|Distance, en points, entre le point de clic gauche et le point de clic gauche vers le bord gauche (ou à droite pour les langues qui s’viennent de droite à gauche) du quadrillage de la cellule cliquée/tapée à gauche.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|La distance en points à partir du point sur lequel vous avez effectué un clic gauche/avez appuyé vers le bord supérieur du bord de la grille de la cellule clic gauche/tape.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la cellule a été sélectionnée par clic-gauche/tape.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)