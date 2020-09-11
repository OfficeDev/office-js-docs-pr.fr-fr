---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,10
description: Détails sur l’ensemble de conditions requises ExcelApi 1,10
ms.date: 10/22/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 890d198f238e29d39744d87d754381543ebcaf6a
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431233"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Nouveautés de l’API JavaScript pour Excel 1,10

Le ExcelApi 1,10 a introduit des fonctionnalités clés, telles que des commentaires, des plans et des segments. Elle a également ajouté la prise en charge des événements de clic et de tri au niveau de la feuille de calcul.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Commentaires](../../excel/excel-add-ins-comments.md) | Ajouter, modifier et supprimer des listes. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Décrit](../../excel/excel-add-ins-ranges-advanced.md#group-data-for-an-outline) | Grouper des lignes et des colonnes pour former des plans de développement/réduction. | [Plage](/javascript/api/excel/excel.range), [feuille de calcul](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#slicers) | Insérer et configurez segments aux tableaux et tableaux croisés dynamiques. | [Segment](/javascript/api/excel/excel.slicer) |
| [Autres événements de feuille de calcul](../../excel/excel-add-ins-events.md) | Écouter les événements Click et sort dans la feuille de calcul. | [Feuille de calcul (événements)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Excel 1,10. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1,10 ou antérieure, voir [API Excel dans l’ensemble de conditions requises 1,10 ou version antérieure](/javascript/api/excel?view=excel-js-1.10&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Obtient ou définit le contenu de commentaire. La chaîne est en texte brut.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Supprime le commentaire et toutes les réponses connectées.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Obtient la cellule dans laquelle se trouve ce commentaire.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Obtient le nom de l’auteur du commentaire.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Obtenir l’heure de création du commentaire. Renvoie la valeur null si le commentaire a été converti d’une note, étant donné que le commentaire ne dispose pas d’une date de création.|
||[id](/javascript/api/excel/excel.comment#id)|Représente l’identificateur de commentaire. En lecture seule.|
||[Réponses](/javascript/api/excel/excel.comment#replies)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (cellAddress : Range \| String, content : CommentRichContent \| String, ContentType ?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée. Une `InvalidArgument` erreur est générée si la plage fournie est plus grande qu’une cellule.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtient le nombre de commentaires de la collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Obtient un commentaire à partir de la collection de sites en fonction de son ID. En lecture seule.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtient un commentaire en fonction de sa position dans la collection.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtient le commentaire à partir de la cellule spécifiée.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtient le commentaire auquel la réponse donnée est connectée.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Obtient ou définit le contenu de la réponse au commentaire. La chaîne est en texte brut.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Supprime la réponse de commentaire.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Obtient la cellule où se trouve cette réponse de commentaire.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtient le commentaire parent de cette réponse.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Obtient l’heure de création de la réponse à un commentaire.|
||[id](/javascript/api/excel/excel.commentreply#id)|Représente l’identificateur de réponse du commentaire. En lecture seule.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Add (Content : CommentRichContent \| String, ContentType ?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse à un commentaire pour un commentaire.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtient le nombre de réponses aux commentaires de la collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Renvoie une réponse de commentaire identifié via son ID. En lecture seule.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtient une réponse de commentaire en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)||[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Supprime le tableau croisé dynamique.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Crée un doublon de cette PivotTableStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtient le nom du PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Indique si cet objet PivotTableStyle est en lecture seule. En lecture seule.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Crée un PivotTableStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Obtient le nombre de styles tableaux croisés dynamiques de la collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Obtient PivotTableStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Extrait un tableau croisé dynamique par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Définit le PivotTableStyle par défaut pour la portée de l’objet parent portée.|
|[Range](/javascript/api/excel/excel.range)|[Group (groupOption : Excel. GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|Regroupe les colonnes et les lignes d’un plan.|
||[hideGroupDetails (groupOption : Excel. GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Masque les détails du groupe de lignes ou de colonnes.|
||[height](/javascript/api/excel/excel.range#height)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la plage au bord inférieur de la plage. En lecture seule.|
||[left](/javascript/api/excel/excel.range#left)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la feuille de calcul au bord gauche de la plage. En lecture seule.|
||[top](/javascript/api/excel/excel.range#top)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la feuille de calcul au bord supérieur de la plage. En lecture seule.|
||[width](/javascript/api/excel/excel.range#width)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la plage au bord droit de la plage. En lecture seule.|
||[showGroupDetails (groupOption : Excel. GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Affiche les détails du groupe de lignes ou de colonnes.|
||[dissociation (groupOption : Excel. GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Dissocie les colonnes et les lignes d’un plan.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copie et colle un objet de forme.|
||[placement](/javascript/api/excel/excel.shape#placement)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
|[Segment](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Représente la légende de segment.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Supprime tous les filtres appliqués actuellement sur le tableau.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Supprime le segment.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Renvoie une matrice de noms d’éléments sélectionnés. En lecture seule.|
||[height](/javascript/api/excel/excel.slicer#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[left](/javascript/api/excel/excel.slicer#left)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicer#name)|Représente le nom de la forme.|
||[id](/javascript/api/excel/excel.slicer#id)|Représente l’id unique du segment. En lecture seule.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|True si tous les filtres appliqués actuellement sur le segment sont effacés.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Représente la collection de SlicerItems qui font partie du segment. En lecture seule.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Obtenir la feuille de calcul contenant la plage. En lecture seule.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Sélectionne les éléments du Slicer en fonction de leurs clés. Les sélections précédentes sont effacées.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Représente l’ordre de tri des éléments dans le segment. Les valeurs possibles sont : "DataSourceOrder", "ascending", "Descending".|
||[style](/javascript/api/excel/excel.slicer#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont : « SlicerStyleLight1 » à « SlicerStyleLight6 », « TableStyleOther1 » à « TableStyleOther2 », « SlicerStyleDark1 » et « SlicerStyleDark6 ». Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
||[top](/javascript/api/excel/excel.slicer#top)|Représente la distance, en points, du bord supérieur de la section à la partie droite de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicer#width)|Représente la largeur, en points, de la forme.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[Ajouter (slicerSource : chaîne \| tableau croisé dynamique \| Table, sourceField : chaîne \| PivotField \| nombre \| TableColumn, slicerDestination ? : chaîne \| feuille de calcul)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Ajoute un nouveau segment au classeur.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Renvoie le nombre de séries de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID. Si la feuille de calcul n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True si l’élément de slicer est sélectionné ; sinon False.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True si l’élément de segment comporte des données.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Représente la valeur unique représentant l’élément de segment.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Représente le titre affiché dans l’interface utilisateur.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Renvoie le nombre de segment les éléments dans le segment.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Obtient un segment objet de l’élément à l’aide de son nom ou clé.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Obtient un segment de l’élément à l’aide de son nom ou clé. Si le paramètre n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Supprime le SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Crée un doublon de cette SlicerStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtient le nom de la SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Indique si cet objet SlicerStyle est en lecture seule. En lecture seule.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Crée un SlicerStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Obtient le nombre de styles de slicer de la collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Obtient SlicerStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Obtient un SlicerStyle par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Obtient un SlicerStyle par nom. Si le SlicerStyle n’existe pas, il renvoie un objet null.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Obtient le SlicerStyle par défaut pour la portée de l’objet parent portée.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Supprime le TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Crée un doublon de cette TableStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtient le nom du TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Indique si cet objet TableStyle est en lecture seule. En lecture seule.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Crée un TableStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Obtient le nombre de styles de tableaux de la collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Obtient le TableStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Extrait un TableStyle par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Extrait un TableStyle par nom. Si le TableStyle n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Définit le TableStyle par défaut pour la portée de l’objet parent portée.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Supprime le TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Crée un doublon de cette TimelineStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtient le nom du TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Indique si cet objet TimelineStyle est en lecture seule. En lecture seule.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Crée un TimelineStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Obtient le nombre de styles de délai de la collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Obtient le TimelineStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Obtient un TimelineStyle par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Obtient un TimelineStyle par nom. Si leTimelineStyle n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Définit le TimelineStyle par défaut pour la portée de l’objet parent portée.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtient le segment actif actuel du classeur. S’il n’y a aucun segment actif, une `ItemNotFound` exception est générée.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtient le segment actif actuel du classeur. S’il n’existe aucun segment actif, un objet null est renvoyé.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Représente une collection de styles associés au classeur. En lecture seule.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Représente une collection de PivotTableStyles associée au classeur. En lecture seule.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Représente une collection de styles associés au classeur. En lecture seule.|
||[Slicers](/javascript/api/excel/excel.workbook#slicers)|Représente une collection de styles associés au classeur. En lecture seule.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Représente une collection de TableStyles associés au classeur. En lecture seule.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Représente une collection de TimelineStyles associés au classeur. En lecture seule.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Renvoie une collection de tous les objets Lecteur sur l’ordinateur. En lecture seule.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Se produit lorsqu’une ou plusieurs colonnes ont été triées. Ce problème se produit en raison de l’opération de tri de gauche à droite.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Se produit lorsqu’une ou plusieurs lignes ont été triées. Cela se produit en raison d’une opération de tri de haut en bas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Se produit lorsqu’une action de l’utilisateur clique sur la feuille de calcul. Cet événement ne sera pas déclenché lorsque vous cliquerez dans les cas suivants :|
||[Slicers](/javascript/api/excel/excel.worksheet#slicers)|Renvoie une collection de segments qui font partie de la feuille de calcul. En lecture seule.|
||[showOutlineLevels (rowLevels : nombre, columnLevels : nombre)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Affiche les groupes de lignes ou de colonnes en fonction de leurs niveaux hiérarchiques.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Se produit lorsqu’une ou plusieurs colonnes ont été triées. Ce problème se produit en raison de l’opération de tri de gauche à droite.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Se produit lorsqu’une ou plusieurs lignes ont été triées. Cela se produit en raison d’une opération de tri de haut en bas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Cet événement se produit lorsque l’utilisateur clique dessus ou a cliqué sur une opération dans la collection Worksheet. Cet événement ne sera pas déclenché lorsque vous cliquerez dans les cas suivants :|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique. Seules les colonnes modifiées à la suite de l’opération de tri sont renvoyées.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le tri s’est passé.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique. Seules les lignes modifiées à la suite de l’opération de tri sont renvoyées.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le tri s’est passé.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[adresse](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtient l’adresse qui représente la cellule sur laquelle vous avez fait un clic gauche/appuyé pour une feuille de calcul spécifique.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|Distance, en points, entre le point à gauche ou à gauche (ou à droite pour les langues se lisant de droite à gauche) et le bord gauche de la cellule qui a cliqué dessus.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|La distance en points à partir du point sur lequel vous avez effectué un clic gauche/avez appuyé vers le bord supérieur du bord de la grille de la cellule clic gauche/tape.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la cellule a été sélectionnée par clic-gauche/tape.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)