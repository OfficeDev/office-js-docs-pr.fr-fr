---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.10
description: Détails sur l’ensemble de conditions requises ExcelApi 1.10.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 34c21ad0e90593352ae4042c2be148e607c63164aac1845357e9f96371104f6f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087212"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Nouveautés de l Excel API JavaScript 1.10

ExcelApi 1.10 a introduit des fonctionnalités clés, telles que les commentaires, les contours et les slicers. Il a également ajouté la prise en charge des événements pour le clic et le tri au niveau de la feuille de calcul.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Commentaires](../../excel/excel-add-ins-comments.md) | Ajouter, modifier et supprimer des listes. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Outlines](../../excel/excel-add-ins-ranges-group.md) | Groupez des lignes et des colonnes pour former des contours réductibles. | [Range](/javascript/api/excel/excel.range), [Worksheet](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | Insérer et configurez segments aux tableaux et tableaux croisés dynamiques. | [Segment](/javascript/api/excel/excel.slicer) |
| [Autres événements de feuille de calcul](../../excel/excel-add-ins-events.md) | Écouter les événements de clic et de tri dans la feuille de calcul. | [Feuille de calcul (événements)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.10. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.10 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Contenu du commentaire.|
||[delete()](/javascript/api/excel/excel.comment#delete__)|Supprime le commentaire et toutes les réponses connectées.|
||[getLocation()](/javascript/api/excel/excel.comment#getLocation__)|Obtient la cellule où se trouve ce commentaire.|
||[authorEmail](/javascript/api/excel/excel.comment#authorEmail)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/excel/excel.comment#authorName)|Obtient le nom de l’auteur du commentaire.|
||[creationDate](/javascript/api/excel/excel.comment#creationDate)|Obtenir l’heure de création du commentaire.|
||[id](/javascript/api/excel/excel.comment#id)|Spécifie l’identificateur de commentaire.|
||[Réponses](/javascript/api/excel/excel.comment#replies)|Représente une collection de feuilles de calcul associées au classeur.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getCount__)|Obtient le nombre de commentaires de la collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getItem_commentId_)|Obtient un commentaire à partir de la collection de sites en fonction de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getItemAt_index_)|Obtient un commentaire en fonction de sa position dans la collection.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getItemByCell_cellAddress_)|Obtient le commentaire à partir de la cellule spécifiée.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getItemByReplyId_replyId_)|Obtient le commentaire auquel la réponse donnée est connectée.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Contenu de la réponse au commentaire.|
||[delete()](/javascript/api/excel/excel.commentreply#delete__)|Supprime la réponse de commentaire.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getLocation__)|Obtient la cellule où se trouve cette réponse de commentaire.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getParentComment__)|Obtient le commentaire parent de cette réponse.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authorEmail)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreply#authorName)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationDate)|Obtient l’heure de création de la réponse à un commentaire.|
||[id](/javascript/api/excel/excel.commentreply#id)|Spécifie l’identificateur de réponse du commentaire.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|Crée une réponse de commentaire pour un commentaire.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getCount__)|Obtient le nombre de réponses aux commentaires de la collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItem_commentReplyId_)|Renvoie une réponse de commentaire identifié via son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getItemAt_index_)|Obtient une réponse de commentaire en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enableFieldList)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete__)|Supprime le style de tableau croisé dynamique.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate__)|Crée une copie de ce style de tableau croisé dynamique avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtient le nom du style de tableau croisé dynamique.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readOnly)|Spécifie si cet `PivotTableStyle` objet est en lecture seule.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add_name__makeUniqueName_)|Crée un vide `PivotTableStyle` avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getCount__)|Obtient le nombre de styles tableaux croisés dynamiques de la collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getDefault__)|Obtient le style de tableau croisé dynamique par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItem_name_)|Obtient `PivotTableStyle` une par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItemOrNullObject_name_)|Obtient `PivotTableStyle` une par nom.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.pivottablestylecollection#setDefault_newDefaultStyle_)|Définit le style de tableau croisé dynamique par défaut à utiliser dans l’étendue de l’objet parent.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#group_groupOption_)|Groupe les colonnes et les lignes d’un plan.|
||[hideGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#hideGroupDetails_groupOption_)|Masque les détails du groupe de lignes ou de colonnes.|
||[height](/javascript/api/excel/excel.range#height)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord supérieur de la plage et le bord inférieur de la plage.|
||[left](/javascript/api/excel/excel.range#left)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord gauche de la feuille de calcul et le bord gauche de la plage.|
||[top](/javascript/api/excel/excel.range#top)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord supérieur de la feuille de calcul et le bord supérieur de la plage.|
||[width](/javascript/api/excel/excel.range#width)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord gauche de la plage et le bord droit de la plage.|
||[showGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#showGroupDetails_groupOption_)|Affiche les détails du groupe de lignes ou de colonnes.|
||[ungroup(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#ungroup_groupOption_)|Regroupe les colonnes et les lignes d’un plan.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyTo_destinationSheet_)|Copie et copie un `Shape` objet.|
||[placement](/javascript/api/excel/excel.shape#placement)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
|[Segment](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Représente la légende du slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearFilters__)|Supprime tous les filtres appliqués actuellement sur le tableau.|
||[delete()](/javascript/api/excel/excel.slicer#delete__)|Supprime le segment.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getSelectedItems__)|Renvoie une matrice de noms d’éléments sélectionnés.|
||[height](/javascript/api/excel/excel.slicer#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[left](/javascript/api/excel/excel.slicer#left)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicer#name)|Représente le nom du slicer.|
||[id](/javascript/api/excel/excel.slicer#id)|Représente l’ID unique du slicer.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isFilterCleared)|La valeur est `true` si tous les filtres appliqués actuellement sur le slicer sont effacés.|
||[slicerItems](/javascript/api/excel/excel.slicer#slicerItems)|Représente la collection d’éléments de slicer qui font partie du slicer.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Obtenir la feuille de calcul contenant la plage.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectItems_items_)|Sélectionne les éléments de slicer en fonction de leurs clés.|
||[sortBy](/javascript/api/excel/excel.slicer#sortBy)|Représente l’ordre de tri des éléments dans le segment.|
||[style](/javascript/api/excel/excel.slicer#style)|Valeur constante qui représente le style de slicer.|
||[top](/javascript/api/excel/excel.slicer#top)|Représente la distance, en points, du bord supérieur de la section à la partie droite de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicer#width)|Représente la largeur, en points, de la forme.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[Ajouter (slicerSource : chaîne \| tableau croisé dynamique \| Table, sourceField : chaîne \| PivotField \| nombre \| TableColumn, slicerDestination ? : chaîne \| feuille de calcul)](/javascript/api/excel/excel.slicercollection#add_slicerSource__sourceField__slicerDestination_)|Ajoute un nouveau segment au classeur.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getCount__)|Renvoie le nombre de séries de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getItem_key_)|Obtient un objet slicer à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getItemAt_index_)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getItemOrNullObject_key_)|Obtient un slicer à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isSelected)|La valeur `true` est si l’élément de slicer est sélectionné.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasData)|La valeur est `true` si l’élément de slicer possède des données.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Représente la valeur unique représentant l’élément de segment.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Représente le titre affiché dans l’interface Excel’utilisateur.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getCount__)|Renvoie le nombre de segment les éléments dans le segment.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItem_key_)|Obtient un segment objet de l’élément à l’aide de son nom ou clé.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getItemAt_index_)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItemOrNullObject_key_)|Obtient un segment de l’élément à l’aide de son nom ou clé.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete__)|Supprime le style de slicer.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate__)|Crée une copie de ce style de slicer avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtient le nom du style de slicer.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readOnly)|Spécifie si cet `SlicerStyle` objet est en lecture seule.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add_name__makeUniqueName_)|Crée un style de slicer vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getCount__)|Obtient le nombre de styles de slicer de la collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getDefault__)|Obtient la valeur `SlicerStyle` par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItem_name_)|Obtient `SlicerStyle` une par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItemOrNullObject_name_)|Obtient `SlicerStyle` une par nom.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.slicerstylecollection#setDefault_newDefaultStyle_)|Définit le style de slicer par défaut à utiliser dans l’étendue de l’objet parent.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete__)|Supprime le style de tableau.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate__)|Crée une copie de ce style de tableau avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtient le nom du style de tableau.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readOnly)|Spécifie si cet `TableStyle` objet est en lecture seule.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add_name__makeUniqueName_)|Crée un vide `TableStyle` avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getCount__)|Obtient le nombre de styles de tableaux de la collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getDefault__)|Obtient le style de tableau par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getItem_name_)|Obtient `TableStyle` une par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getItemOrNullObject_name_)|Obtient `TableStyle` une par nom.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.tablestylecollection#setDefault_newDefaultStyle_)|Définit le style de tableau par défaut à utiliser dans l’étendue de l’objet parent.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete__)|Supprime le style de tableau.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate__)|Crée une copie de ce style de chronologie avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtient le nom du style de chronologie.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readOnly)|Spécifie si cet `TimelineStyle` objet est en lecture seule.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add_name__makeUniqueName_)|Crée un vide `TimelineStyle` avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getCount__)|Obtient le nombre de styles de délai de la collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getDefault__)|Obtient le style de chronologie par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItem_name_)|Obtient `TimelineStyle` une par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItemOrNullObject_name_)|Obtient `TimelineStyle` une par nom.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.timelinestylecollection#setDefault_newDefaultStyle_)|Définit le style de chronologie par défaut à utiliser dans l’étendue de l’objet parent.|
|[Classeur](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getActiveSlicer__)|Obtient le segment actif actuel du classeur.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getActiveSlicerOrNullObject__)|Obtient le segment actif actuel du classeur.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Représente une collection de commentaires associés au workbook.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivotTableStyles)|Représente une collection de PivotTableStyles associée au classeur.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerStyles)|Représente une collection de styles associés au classeur.|
||[Slicers](/javascript/api/excel/excel.workbook#slicers)|Représente une collection de slicers associés au workbook.|
||[tableStyles](/javascript/api/excel/excel.workbook#tableStyles)|Représente une collection de TableStyles associés au classeur.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelineStyles)|Représente une collection de TimelineStyles associés au classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Renvoie une collection de tous les objets Lecteur sur l’ordinateur.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#onColumnSorted)|Se produit lorsqu’une ou plusieurs colonnes ont été triées.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onRowSorted)|Se produit lorsqu’une ou plusieurs lignes ont été triées.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onSingleClicked)|Se produit lorsqu’une action clic gauche/clic se produit dans la feuille de calcul.|
||[Slicers](/javascript/api/excel/excel.worksheet#slicers)|Renvoie une collection de slicers qui font partie de la feuille de calcul.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showOutlineLevels_rowLevels__columnLevels_)|Affiche les groupes de lignes ou de colonnes par niveaux de plan.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#onColumnSorted)|Se produit lorsqu’une ou plusieurs colonnes ont été triées.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onRowSorted)|Se produit lorsqu’une ou plusieurs lignes ont été triées.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onSingleClicked)|Se produit lorsque l’opération clic gauche/clic se produit dans la collection de feuilles de calcul.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le tri s’est produit.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le tri s’est produit.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[adresse](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtient l’adresse qui représente la cellule sur laquelle vous avez fait un clic gauche/appuyé pour une feuille de calcul spécifique.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetX)|Distance, en points, entre le point de clic gauche et le point de clic gauche vers le bord gauche (ou de droite pour les langues qui s’viennent de droite à gauche) du quadrillage de la cellule cliquée/tapée à gauche.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetY)|La distance en points à partir du point sur lequel vous avez effectué un clic gauche/avez appuyé vers le bord supérieur du bord de la grille de la cellule clic gauche/tape.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle la cellule a été cliquée/tapée avec le bouton gauche.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)