---
title: Excel’ensemble de conditions requises de l’API JavaScript 1.10
description: Détails sur l’ensemble de conditions requises ExcelApi 1.10.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 53cf0ec55a26f02a615a3c5eee0b718b818790d0
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746340"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Nouveautés de l Excel API JavaScript 1.10

ExcelApi 1.10 a introduit des fonctionnalités clés, telles que les commentaires, les contours et les slicers. Il a également ajouté la prise en charge des événements pour le clic et le tri au niveau de la feuille de calcul.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Commentaires](../../excel/excel-add-ins-comments.md) | Ajouter, modifier et supprimer des listes. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Outlines](../../excel/excel-add-ins-ranges-group.md) | Groupez les lignes et les colonnes pour former des contours réductibles. | [Plage](/javascript/api/excel/excel.range), [feuille de calcul](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | Insérer et configurez segments aux tableaux et tableaux croisés dynamiques. | [Segment](/javascript/api/excel/excel.slicer) |
| [Autres événements de feuille de calcul](../../excel/excel-add-ins-events.md) | Écouter les événements de clic et de tri dans la feuille de calcul. | [Feuille de calcul (événements)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-events-member) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API Excel l’ensemble de conditions requises de l’API JavaScript 1.10. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.10 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true) ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[authorEmail](/javascript/api/excel/excel.comment#excel-excel-comment-authoremail-member)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/excel/excel.comment#excel-excel-comment-authorname-member)|Obtient le nom de l’auteur du commentaire.|
||[content](/javascript/api/excel/excel.comment#excel-excel-comment-content-member)|Contenu du commentaire.|
||[creationDate](/javascript/api/excel/excel.comment#excel-excel-comment-creationdate-member)|Obtenir l’heure de création du commentaire.|
||[delete()](/javascript/api/excel/excel.comment#excel-excel-comment-delete-member(1))|Supprime le commentaire et toutes les réponses connectées.|
||[getLocation()](/javascript/api/excel/excel.comment#excel-excel-comment-getlocation-member(1))|Obtient la cellule où se trouve ce commentaire.|
||[id](/javascript/api/excel/excel.comment#excel-excel-comment-id-member)|Spécifie l’identificateur de commentaire.|
||[Réponses](/javascript/api/excel/excel.comment#excel-excel-comment-replies-member)|Représente une collection de feuilles de calcul associées au classeur.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée.|
||[getCount()](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getcount-member(1))|Obtient le nombre de commentaires de la collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitem-member(1))|Obtient un commentaire à partir de la collection de sites en fonction de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemat-member(1))|Obtient un commentaire en fonction de sa position dans la collection.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembycell-member(1))|Obtient le commentaire à partir de la cellule spécifiée.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembyreplyid-member(1))|Obtient le commentaire auquel la réponse donnée est connectée.|
||[items](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[authorEmail](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authoremail-member)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authorname-member)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[content](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-content-member)|Contenu de la réponse au commentaire.|
||[creationDate](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-creationdate-member)|Obtient l’heure de création de la réponse à un commentaire.|
||[delete()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-delete-member(1))|Supprime la réponse de commentaire.|
||[getLocation()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getlocation-member(1))|Obtient la cellule où se trouve cette réponse de commentaire.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getparentcomment-member(1))|Obtient le commentaire parent de cette réponse.|
||[id](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-id-member)|Spécifie l’identificateur de réponse du commentaire.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|Crée une réponse de commentaire pour un commentaire.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getcount-member(1))|Obtient le nombre de réponses aux commentaires de la collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitem-member(1))|Renvoie une réponse de commentaire identifié via son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemat-member(1))|Obtient une réponse de commentaire en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-enablefieldlist-member)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-delete-member(1))|Supprime le style de tableau croisé dynamique.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-duplicate-member(1))|Crée une copie de ce style de tableau croisé dynamique avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-name-member)|Obtient le nom du style de tableau croisé dynamique.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-readonly-member)|Spécifie si cet objet `PivotTableStyle` est en lecture seule.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-add-member(1))|Crée un vide `PivotTableStyle` avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getcount-member(1))|Obtient le nombre de styles tableaux croisés dynamiques de la collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getdefault-member(1))|Obtient le style de tableau croisé dynamique par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitem-member(1))|Obtient une `PivotTableStyle` par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitemornullobject-member(1))|Obtient une `PivotTableStyle` par nom.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-setdefault-member(1))|Définit le style de tableau croisé dynamique par défaut à utiliser dans l’étendue de l’objet parent.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-group-member(1))|Groupe les colonnes et les lignes d’un plan.|
||[height](/javascript/api/excel/excel.range#excel-excel-range-height-member)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord supérieur de la plage et le bord inférieur de la plage.|
||[hideGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-hidegroupdetails-member(1))|Masque les détails du groupe de lignes ou de colonnes.|
||[left](/javascript/api/excel/excel.range#excel-excel-range-left-member)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord gauche de la feuille de calcul et le bord gauche de la plage.|
||[showGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-showgroupdetails-member(1))|Affiche les détails du groupe de lignes ou de colonnes.|
||[top](/javascript/api/excel/excel.range#excel-excel-range-top-member)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord supérieur de la feuille de calcul et le bord supérieur de la plage.|
||[ungroup(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1))|Regroupe les colonnes et les lignes d’un plan.|
||[width](/javascript/api/excel/excel.range#excel-excel-range-width-member)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord gauche de la plage et le bord droit de la plage.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#excel-excel-shape-copyto-member(1))|Copie et copie un `Shape` objet.|
||[placement](/javascript/api/excel/excel.shape#excel-excel-shape-placement-member)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
|[Segment](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#excel-excel-slicer-caption-member)|Représente la légende du slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#excel-excel-slicer-clearfilters-member(1))|Supprime tous les filtres appliqués actuellement sur le tableau.|
||[delete()](/javascript/api/excel/excel.slicer#excel-excel-slicer-delete-member(1))|Supprime le segment.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#excel-excel-slicer-getselecteditems-member(1))|Renvoie une matrice de noms d’éléments sélectionnés.|
||[height](/javascript/api/excel/excel.slicer#excel-excel-slicer-height-member)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[id](/javascript/api/excel/excel.slicer#excel-excel-slicer-id-member)|Représente l’ID unique du slicer.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#excel-excel-slicer-isfiltercleared-member)|La valeur est `true` si tous les filtres appliqués actuellement sur le slicer sont effacés.|
||[left](/javascript/api/excel/excel.slicer#excel-excel-slicer-left-member)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicer#excel-excel-slicer-name-member)|Représente le nom du slicer.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#excel-excel-slicer-selectitems-member(1))|Sélectionne les éléments de slicer en fonction de leurs clés.|
||[slicerItems](/javascript/api/excel/excel.slicer#excel-excel-slicer-sliceritems-member)|Représente la collection d’éléments de slicer qui font partie du slicer.|
||[sortBy](/javascript/api/excel/excel.slicer#excel-excel-slicer-sortby-member)|Représente l’ordre de tri des éléments dans le segment.|
||[style](/javascript/api/excel/excel.slicer#excel-excel-slicer-style-member)|Valeur constante qui représente le style de slicer.|
||[top](/javascript/api/excel/excel.slicer#excel-excel-slicer-top-member)|Représente la distance, en points, du bord supérieur de la section à la partie droite de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicer#excel-excel-slicer-width-member)|Représente la largeur, en points, de la forme.|
||[worksheet](/javascript/api/excel/excel.slicer#excel-excel-slicer-worksheet-member)|Obtenir la feuille de calcul contenant la plage.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[Ajouter (slicerSource : chaîne \| tableau croisé dynamique \| Table, sourceField : chaîne \| PivotField \| nombre \| TableColumn, slicerDestination ? : chaîne \| feuille de calcul)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-add-member(1))|Ajoute un nouveau segment au classeur.|
||[getCount()](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getcount-member(1))|Renvoie le nombre de séries de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitem-member(1))|Obtient un objet slicer à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemat-member(1))|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemornullobject-member(1))|Obtient un slicer à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[hasData](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-hasdata-member)|La valeur est `true` si l’élément de slicer possède des données.|
||[isSelected](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-isselected-member)|La valeur est `true` si l’élément de slicer est sélectionné.|
||[key](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-key-member)|Représente la valeur unique représentant l’élément de segment.|
||[name](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-name-member)|Représente le titre affiché dans l’interface Excel’utilisateur.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getcount-member(1))|Renvoie le nombre de segment les éléments dans le segment.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitem-member(1))|Obtient un segment objet de l’élément à l’aide de son nom ou clé.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemat-member(1))|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemornullobject-member(1))|Obtient un segment de l’élément à l’aide de son nom ou clé.|
||[items](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-delete-member(1))|Supprime le style de slicer.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-duplicate-member(1))|Crée une copie de ce style de slicer avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-name-member)|Obtient le nom du style de slicer.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-readonly-member)|Spécifie si cet objet `SlicerStyle` est en lecture seule.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-add-member(1))|Crée un style de slicer vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getcount-member(1))|Obtient le nombre de styles de slicer de la collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getdefault-member(1))|Obtient la valeur par `SlicerStyle` défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitem-member(1))|Obtient une `SlicerStyle` par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitemornullobject-member(1))|Obtient une `SlicerStyle` par nom.|
||[items](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-setdefault-member(1))|Définit le style de slicer par défaut à utiliser dans l’étendue de l’objet parent.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-delete-member(1))|Supprime le style de tableau.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-duplicate-member(1))|Crée une copie de ce style de tableau avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-name-member)|Obtient le nom du style de tableau.|
||[readOnly](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-readonly-member)|Spécifie si cet objet `TableStyle` est en lecture seule.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-add-member(1))|Crée un vide `TableStyle` avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getcount-member(1))|Obtient le nombre de styles de tableaux de la collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getdefault-member(1))|Obtient le style de tableau par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitem-member(1))|Obtient une `TableStyle` par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitemornullobject-member(1))|Obtient une `TableStyle` par nom.|
||[items](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-setdefault-member(1))|Définit le style de tableau par défaut à utiliser dans l’étendue de l’objet parent.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-delete-member(1))|Supprime le style de tableau.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-duplicate-member(1))|Crée un doublon de ce style de chronologie avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-name-member)|Obtient le nom du style de chronologie.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-readonly-member)|Spécifie si cet objet `TimelineStyle` est en lecture seule.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-add-member(1))|Crée un vide `TimelineStyle` avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getcount-member(1))|Obtient le nombre de styles de délai de la collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getdefault-member(1))|Obtient le style de chronologie par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitem-member(1))|Obtient une `TimelineStyle` par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitemornullobject-member(1))|Obtient une `TimelineStyle` par nom.|
||[items](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-setdefault-member(1))|Définit le style de chronologie par défaut à utiliser dans l’étendue de l’objet parent.|
|[Workbook](/javascript/api/excel/excel.workbook)|[comments](/javascript/api/excel/excel.workbook#excel-excel-workbook-comments-member)|Représente une collection de commentaires associés au workbook.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicer-member(1))|Obtient le segment actif actuel du classeur.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicerornullobject-member(1))|Obtient le segment actif actuel du classeur.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottablestyles-member)|Représente une collection de PivotTableStyles associée au classeur.|
||[slicerStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicerstyles-member)|Représente une collection de styles associés au classeur.|
||[Slicers](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicers-member)|Représente une collection de slicers associés au workbook.|
||[tableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-tablestyles-member)|Représente une collection de TableStyles associés au classeur.|
||[timelineStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-timelinestyles-member)|Représente une collection de TimelineStyles associés au classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-comments-member)|Renvoie une collection de tous les objets Lecteur sur l’ordinateur.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)|Se produit lorsqu’une ou plusieurs colonnes ont été triées.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member)|Se produit lorsqu’une ou plusieurs lignes ont été triées.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member)|Se produit lorsqu’une action clic gauche/clic se produit dans la feuille de calcul.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1))|Affiche les groupes de lignes ou de colonnes par niveaux de plan.|
||[Slicers](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-slicers-member)|Renvoie une collection de slicers qui font partie de la feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member)|Se produit lorsqu’une ou plusieurs colonnes ont été triées.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member)|Se produit lorsqu’une ou plusieurs lignes ont été triées.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member)|Se produit lorsque l’opération clic gauche/clic se produit dans la collection de feuilles de calcul.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-address-member)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-source-member)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle le tri s’est produit.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-address-member)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-source-member)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle le tri s’est produit.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[adresse](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-address-member)|Obtient l’adresse qui représente la cellule sur laquelle vous avez fait un clic gauche/appuyé pour une feuille de calcul spécifique.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsetx-member)|Distance, en points, entre le point de clic gauche et le point de clic gauche vers le bord gauche (ou à droite pour les langues qui s’viennent de droite à gauche) du quadrillage de la cellule cliquée/tapée à gauche.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsety-member)|La distance en points à partir du point sur lequel vous avez effectué un clic gauche/avez appuyé vers le bord supérieur du bord de la grille de la cellule clic gauche/tape.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle la cellule a été cliquée/tapée avec le bouton gauche.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)