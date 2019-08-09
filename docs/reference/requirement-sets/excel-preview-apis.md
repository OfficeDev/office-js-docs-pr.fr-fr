---
title: Version d’évaluation API JavaScript Excel
description: Informations détaillées sur les API JavaScript pour Excel à venir
ms.date: 08/06/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4362a5e13e0031408236f34c718f0fcb3c4527e2
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268732"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

> [!NOTE]
> L’aperçu API peut être modifiés et n’est pas destinés à utiliser dans un environnement de production. Nous vous recommandons de les tester uniquement dans les environnements de test et de développement. N’utilisez pas un aperçu d’API dans un environnement de production ou dans les documents commerciaux importants.
>
> Pour utiliser l’aperçu API, vous devez référencer la bibliothèque**bêta**sur le CDN : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) et vous devrez également participer au programme Office Insider pour obtenir un build Office suffisamment récent.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Segment](../../excel/excel-add-ins-pivottables.md#slicers-preview) | Insérer et configurez segments aux tableaux et tableaux croisés dynamiques. | [Segment](/javascript/api/excel/excel.slicer) |
| [Commentaires](../../excel/excel-add-ins-workbooks.md#comments-preview) | Ajouter, modifier et supprimer des listes. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Classeur [enregistrer](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) et [fermer](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | Enregistrez et fermez ses classeurs.  | [Workbook](/javascript/api/excel/excel.workbook) |
| [Insérer le classeur](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insérer un classeur dans un autre.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Excel actuellement en version préliminaire. Pour afficher la liste complète de toutes les API JavaScript pour Excel (y compris les API d’aperçu et les API précédemment publiées), voir [toutes les API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview).

| Class | Champs | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Obtient ou définit le contenu de commentaire. La chaîne est en texte brut.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Supprime le thread de commentaires.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Obtient la cellule dans laquelle se trouve ce commentaire.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Obtient le nom de l’auteur du commentaire.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Obtenir l’heure de création du commentaire. Renvoie la valeur null si le commentaire a été converti d’une note, étant donné que le commentaire ne dispose pas d’une date de création.|
||[id](/javascript/api/excel/excel.comment#id)|Représente l’identificateur de commentaire. En lecture seule.|
||[Réponses](/javascript/api/excel/excel.comment#replies)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|
||[évaluation](/javascript/api/excel/excel.comment#resolved)|Obtient ou définit l’état du thread de commentaire. La valeur «true» signifie que le thread de commentaire est dans l’État résolu.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Crée un nouveau commentaire (thread de commentaire) avec le contenu donné sur la cellule donnée. Une `InvalidArgument` erreur est générée si la plage fournie est plus grande qu’une cellule.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtient le nombre de commentaires de la collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Obtient un commentaire à partir de la collection de sites en fonction de son ID. En lecture seule.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtient un commentaire en fonction de sa position dans la collection.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtient le commentaire à partir de la cellule spécifiée.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtient un commentaire lié à son ID dans la collection de réponse.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Obtient ou définit le contenu de la réponse au commentaire. La chaîne est en texte brut.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Supprime la réponse de commentaire.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Obtient la cellule où se trouve cette réponse de commentaire.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtient le commentaire parent de cette réponse.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Obtient l’heure de création de la réponse à un commentaire.|
||[id](/javascript/api/excel/excel.commentreply#id)|Représente l’identificateur de réponse du commentaire. En lecture seule.|
||[évaluation](/javascript/api/excel/excel.commentreply#resolved)|Obtient ou définit l’état de la réponse de commentaire. La valeur «true» signifie que la réponse au commentaire est dans l’État résolu.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse à un commentaire pour un commentaire.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtient le nombre de réponses aux commentaires de la collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Renvoie une réponse de commentaire identifié via son ID. En lecture seule.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtient une réponse de commentaire en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives. La cellule renvoyée est l’intersection de la ligne donnée et une colonne qui contient les données à partir de la hiérarchie donnée. Cette méthode est l’inverse de l’appel getPivotItems et getDataHierarchy sur une cellule particulière.|
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
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. En lecture seule.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. En lecture seule.|
||[Group (groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|Regroupe les colonnes et les lignes d’un plan.|
||[hideGroupDetails (groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Masque les détails du groupe de lignes ou de colonnes.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[height](/javascript/api/excel/excel.range#height)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la plage au bord inférieur de la plage. En lecture seule.|
||[left](/javascript/api/excel/excel.range#left)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la feuille de calcul au bord gauche de la plage. En lecture seule.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Représente si toutes les cellules sont enregistrées sous la forme d’une formule matricielle.|
||[top](/javascript/api/excel/excel.range#top)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la feuille de calcul au bord supérieur de la plage. En lecture seule.|
||[width](/javascript/api/excel/excel.range#width)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la plage au bord droit de la plage. En lecture seule.|
||[showGroupDetails (groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Affiche les détails du groupe de lignes ou de colonnes.|
||[dissociation (groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Dissocie les colonnes et les lignes d’un plan.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copie et colle un objet de forme.|
||[placement](/javascript/api/excel/excel.shape#placement)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul. Renvoie un objet Forme qui représente la nouvelle image.|
|[Segment](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Représente la légende de segment.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Supprime tous les filtres appliqués actuellement sur le tableau.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Supprime le segment.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Renvoie une matrice de noms d’éléments sélectionnés. En lecture seule.|
||[height](/javascript/api/excel/excel.slicer#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[left](/javascript/api/excel/excel.slicer#left)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicer#name)|Représente le nom de la forme.|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[id](/javascript/api/excel/excel.slicer#id)|Représente l’id unique du segment. En lecture seule.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|True si tous les filtres appliqués actuellement sur le segment sont effacés.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Représente la collection de SlicerItems qui font partie du segment. En lecture seule.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Obtenir la feuille de calcul contenant la plage. En lecture seule.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Éléments de segment à sélection multiple en fonction de leur nom. Sélection précédente est désactivée.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Représente l’ordre de tri des éléments dans le segment. Les valeurs possibles sont : DataSourceOrder par ordre croissant, par ordre décroissant.|
||[style](/javascript/api/excel/excel.slicer#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont: «SlicerStyleLight1» à «SlicerStyleLight6», «TableStyleOther1» à «TableStyleOther2», «SlicerStyleDark1» et «SlicerStyleDark6». Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
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
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Indique le nom de la table à laquelle le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Représente l’id de la feuille de calcul qui contient le tableau.|
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
|[Workbook](/javascript/api/excel/excel.workbook)|[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtient le segment actif actuel du classeur. S’il n’y a aucun segment actif, `ItemNotFound` une exception est générée.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtient le segment actif actuel du classeur. S’il n’existe aucun segment actif, un objet null est renvoyé.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Représente une collection de styles associés au classeur. En lecture seule.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Représente une collection de PivotTableStyles associée au classeur. En lecture seule.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Représente une collection de styles associés au classeur. En lecture seule.|
||[Slicers](/javascript/api/excel/excel.workbook#slicers)|Représente une collection de styles associés au classeur. En lecture seule.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Représente une collection de TableStyles associés au classeur. En lecture seule.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Représente une collection de TimelineStyles associés au classeur. En lecture seule.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Renvoie une collection de tous les objets Lecteur sur l’ordinateur. En lecture seule.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Se produit lors du tri des colonnes.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Se produit lorsque l’état de la ligne masquée d’une feuille de calcul spécifique est modifié.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Se produit lors du tri des lignes.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Se produit lorsque l’opération clic gauche/tape se produit dans la feuille de calcul. Cet événement ne sera pas déclenché lorsque vous cliquerez dans les cas suivants: [...]|
||[Slicers](/javascript/api/excel/excel.worksheet#slicers)|Renvoie une collection de graphiques qui font partie de la feuille de calcul. En lecture seule.|
||[showOutlineLevels (rowLevels: nombre, columnLevels: nombre)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Affiche les groupes de lignes ou de colonnes en fonction de leurs niveaux hiérarchiques.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Se produit lors du tri des colonnes.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Cet événement se produit lorsqu’un état masqué d’une feuille de calcul du classeur a été modifié.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Se produit lors du tri des lignes.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Cet événement se produit lorsque l’utilisateur clique dessus ou a cliqué sur une opération dans la collection Worksheet. Cet événement ne sera pas déclenché lorsque vous cliquerez dans les cas suivants: [...]|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le tri s’est passé.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Indique le nom du tableau auquel le filtre est appliqué.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtient le type de modification qui représente le mode de déclenchement de l’événement RowHiddenChanged. Pour plus d’informations, voir Excel. RowHiddenChangeType.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le tri s’est passé.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[adresse](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtient l’adresse qui représente la cellule sur laquelle vous avez fait un clic gauche/appuyé pour une feuille de calcul spécifique.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|Distance, en points, entre le point à gauche ou à gauche (ou à droite pour les langues se lisant de droite à gauche) et le bord gauche de la cellule qui a cliqué dessus.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|La distance en points à partir du point sur lequel vous avez effectué un clic gauche/avez appuyé vers le bord supérieur du bord de la grille de la cellule clic gauche/tape.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la cellule a été sélectionnée par clic-gauche/tape.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
