---
title: Version d’évaluation API JavaScript Excel
description: Informations détaillées sur les API JavaScript pour Excel à venir
ms.date: 07/25/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 200b187b059c1b03ae3713b5afa11b2152aba0da
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064850"
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
||[Set (propriétés: Excel. Comment)](/javascript/api/excel/excel.comment#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. CommentUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.comment#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Crée un nouveau commentaire (thread de commentaire) avec le contenu donné sur la cellule donnée. Une `InvalidArgument` erreur est générée si la plage fournie est plus grande qu’une cellule.|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Crée un nouveau commentaire (thread de commentaire) avec le contenu donné sur la cellule donnée. Une `InvalidArgument` erreur est générée si la plage fournie est plus grande qu’une cellule.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtient le nombre de commentaires de la collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Obtient un commentaire à partir de la collection de sites en fonction de son ID. En lecture seule.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtient un commentaire en fonction de sa position dans la collection.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtient le commentaire à partir de la cellule spécifiée.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtient un commentaire lié à son ID dans la collection de réponse.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentCollectionData](/javascript/api/excel/excel.commentcollectiondata)|[items](/javascript/api/excel/excel.commentcollectiondata#items)||
|[CommentCollectionLoadOptions](/javascript/api/excel/excel.commentcollectionloadoptions)|[$all](/javascript/api/excel/excel.commentcollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentcollectionloadoptions#authoremail)|Pour chaque élément de la collection: obtient le message électronique de l’auteur du commentaire.|
||[authorName](/javascript/api/excel/excel.commentcollectionloadoptions#authorname)|Pour chaque élément de la collection: obtient le nom de l’auteur du commentaire.|
||[content](/javascript/api/excel/excel.commentcollectionloadoptions#content)|Pour chaque élément de la collection: Obtient ou définit le contenu du commentaire. La chaîne est en texte brut.|
||[creationDate](/javascript/api/excel/excel.commentcollectionloadoptions#creationdate)|Pour chaque élément de la collection: obtient l’heure de création du commentaire. Renvoie la valeur null si le commentaire a été converti d’une note, étant donné que le commentaire ne dispose pas d’une date de création.|
||[id](/javascript/api/excel/excel.commentcollectionloadoptions#id)|Pour chaque élément de la collection: représente l’identificateur de commentaire. En lecture seule.|
|[CommentCollectionUpdateData](/javascript/api/excel/excel.commentcollectionupdatedata)|[items](/javascript/api/excel/excel.commentcollectionupdatedata#items)||
|[CommentData](/javascript/api/excel/excel.commentdata)|[authorEmail](/javascript/api/excel/excel.commentdata#authoremail)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/excel/excel.commentdata#authorname)|Obtient le nom de l’auteur du commentaire.|
||[content](/javascript/api/excel/excel.commentdata#content)|Obtient ou définit le contenu de commentaire. La chaîne est en texte brut.|
||[creationDate](/javascript/api/excel/excel.commentdata#creationdate)|Obtenir l’heure de création du commentaire. Renvoie la valeur null si le commentaire a été converti d’une note, étant donné que le commentaire ne dispose pas d’une date de création.|
||[id](/javascript/api/excel/excel.commentdata#id)|Représente l’identificateur de commentaire. En lecture seule.|
||[Réponses](/javascript/api/excel/excel.commentdata#replies)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|
|[CommentLoadOptions](/javascript/api/excel/excel.commentloadoptions)|[$all](/javascript/api/excel/excel.commentloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentloadoptions#authoremail)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/excel/excel.commentloadoptions#authorname)|Obtient le nom de l’auteur du commentaire.|
||[content](/javascript/api/excel/excel.commentloadoptions#content)|Obtient ou définit le contenu de commentaire. La chaîne est en texte brut.|
||[creationDate](/javascript/api/excel/excel.commentloadoptions#creationdate)|Obtenir l’heure de création du commentaire. Renvoie la valeur null si le commentaire a été converti d’une note, étant donné que le commentaire ne dispose pas d’une date de création.|
||[id](/javascript/api/excel/excel.commentloadoptions#id)|Représente l’identificateur de commentaire. En lecture seule.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Obtient ou définit le contenu de la réponse au commentaire. La chaîne est en texte brut.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Supprime la réponse de commentaire.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Obtient la cellule où se trouve cette réponse de commentaire.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtient le commentaire parent de cette réponse.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Obtient l’heure de création de la réponse à un commentaire.|
||[id](/javascript/api/excel/excel.commentreply#id)|Représente l’identificateur de réponse du commentaire. En lecture seule.|
||[Set (propriétés: Excel. CommentReply)](/javascript/api/excel/excel.commentreply#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. CommentReplyUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.commentreply#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse à un commentaire pour un commentaire.|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse à un commentaire pour un commentaire.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtient le nombre de réponses aux commentaires de la collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Renvoie une réponse de commentaire identifié via son ID. En lecture seule.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtient une réponse de commentaire en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentReplyCollectionData](/javascript/api/excel/excel.commentreplycollectiondata)|[items](/javascript/api/excel/excel.commentreplycollectiondata#items)||
|[CommentReplyCollectionLoadOptions](/javascript/api/excel/excel.commentreplycollectionloadoptions)|[$all](/javascript/api/excel/excel.commentreplycollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplycollectionloadoptions#authoremail)|Pour chaque élément de la collection: obtient le message électronique de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreplycollectionloadoptions#authorname)|Pour chaque élément de la collection: obtient le nom de l’auteur de la réponse au commentaire.|
||[content](/javascript/api/excel/excel.commentreplycollectionloadoptions#content)|Pour chaque élément de la collection: Obtient ou définit le contenu de la réponse de commentaire. La chaîne est en texte brut.|
||[creationDate](/javascript/api/excel/excel.commentreplycollectionloadoptions#creationdate)|Pour chaque élément de la collection: obtient l’heure de création de la réponse au commentaire.|
||[id](/javascript/api/excel/excel.commentreplycollectionloadoptions#id)|Pour chaque élément de la collection: représente l’identificateur de réponse de commentaire. En lecture seule.|
|[CommentReplyCollectionUpdateData](/javascript/api/excel/excel.commentreplycollectionupdatedata)|[items](/javascript/api/excel/excel.commentreplycollectionupdatedata#items)||
|[CommentReplyData](/javascript/api/excel/excel.commentreplydata)|[authorEmail](/javascript/api/excel/excel.commentreplydata#authoremail)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreplydata#authorname)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[content](/javascript/api/excel/excel.commentreplydata#content)|Obtient ou définit le contenu de la réponse au commentaire. La chaîne est en texte brut.|
||[creationDate](/javascript/api/excel/excel.commentreplydata#creationdate)|Obtient l’heure de création de la réponse à un commentaire.|
||[id](/javascript/api/excel/excel.commentreplydata#id)|Représente l’identificateur de réponse du commentaire. En lecture seule.|
|[CommentReplyLoadOptions](/javascript/api/excel/excel.commentreplyloadoptions)|[$all](/javascript/api/excel/excel.commentreplyloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplyloadoptions#authoremail)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreplyloadoptions#authorname)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[content](/javascript/api/excel/excel.commentreplyloadoptions#content)|Obtient ou définit le contenu de la réponse au commentaire. La chaîne est en texte brut.|
||[creationDate](/javascript/api/excel/excel.commentreplyloadoptions#creationdate)|Obtient l’heure de création de la réponse à un commentaire.|
||[id](/javascript/api/excel/excel.commentreplyloadoptions#id)|Représente l’identificateur de réponse du commentaire. En lecture seule.|
|[CommentReplyUpdateData](/javascript/api/excel/excel.commentreplyupdatedata)|[content](/javascript/api/excel/excel.commentreplyupdatedata#content)|Obtient ou définit le contenu de la réponse au commentaire. La chaîne est en texte brut.|
|[CommentUpdateData](/javascript/api/excel/excel.commentupdatedata)|[content](/javascript/api/excel/excel.commentupdatedata#content)|Obtient ou définit le contenu de commentaire. La chaîne est en texte brut.|
|[GroupShapeCollectionLoadOptions](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[placement](/javascript/api/excel/excel.groupshapecollectionloadoptions#placement)|Pour chaque élément de la collection: représente la manière dont l’objet est attaché aux cellules qui se trouvent en-dessous.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives. La cellule renvoyée est l’intersection de la ligne donnée et une colonne qui contient les données à partir de la hiérarchie donnée. Cette méthode est l’inverse de l’appel getPivotItems et getDataHierarchy sur une cellule particulière.|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutdata#enablefieldlist)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutloadoptions#enablefieldlist)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutupdatedata#enablefieldlist)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Supprime le tableau croisé dynamique.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Crée un doublon de cette PivotTableStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtient le nom du PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Spécifie si cet objet PivotTableStyle est accessible en lecture seule. En lecture seule.|
||[Set (propriétés: Excel. PivotTableStyle)](/javascript/api/excel/excel.pivottablestyle#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. PivotTableStyleUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.pivottablestyle#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Crée un PivotTableStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Obtient le nombre de styles tableaux croisés dynamiques de la collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Obtient PivotTableStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Extrait un tableau croisé dynamique par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Définit le PivotTableStyle par défaut pour la portée de l’objet parent portée.|
|[PivotTableStyleCollectionData](/javascript/api/excel/excel.pivottablestylecollectiondata)|[items](/javascript/api/excel/excel.pivottablestylecollectiondata#items)||
|[PivotTableStyleCollectionLoadOptions](/javascript/api/excel/excel.pivottablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#name)|Pour chaque élément de la collection: obtient le nom du PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#readonly)|Pour chaque élément de la collection: indique si cet objet PivotTableStyle est en lecture seule. En lecture seule.|
|[PivotTableStyleCollectionUpdateData](/javascript/api/excel/excel.pivottablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.pivottablestylecollectionupdatedata#items)||
|[PivotTableStyleData](/javascript/api/excel/excel.pivottablestyledata)|[name](/javascript/api/excel/excel.pivottablestyledata#name)|Obtient le nom du PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyledata#readonly)|Spécifie si cet objet PivotTableStyle est accessible en lecture seule. En lecture seule.|
|[PivotTableStyleLoadOptions](/javascript/api/excel/excel.pivottablestyleloadoptions)|[$all](/javascript/api/excel/excel.pivottablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestyleloadoptions#name)|Obtient le nom du PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyleloadoptions#readonly)|Spécifie si cet objet PivotTableStyle est accessible en lecture seule. En lecture seule.|
|[PivotTableStyleUpdateData](/javascript/api/excel/excel.pivottablestyleupdatedata)|[name](/javascript/api/excel/excel.pivottablestyleupdatedata#name)|Obtient le nom du PivotTableStyle.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. En lecture seule.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. En lecture seule.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[height](/javascript/api/excel/excel.range#height)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la plage au bord inférieur de la plage. En lecture seule.|
||[left](/javascript/api/excel/excel.range#left)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la feuille de calcul au bord gauche de la plage. En lecture seule.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Représente si toutes les cellules sont enregistrées sous la forme d’une formule matricielle.|
||[top](/javascript/api/excel/excel.range#top)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la feuille de calcul au bord supérieur de la plage. En lecture seule.|
||[width](/javascript/api/excel/excel.range#width)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la plage au bord droit de la plage. En lecture seule.|
|[RangeCollectionLoadOptions](/javascript/api/excel/excel.rangecollectionloadoptions)|[hasSpill](/javascript/api/excel/excel.rangecollectionloadoptions#hasspill)|Pour chaque élément de la collection: représente si toutes les cellules ont une bordure de renversement.|
||[height](/javascript/api/excel/excel.rangecollectionloadoptions#height)|Pour chaque élément de la collection: renvoie la distance en points, pour un zoom de 100%, entre le bord supérieur de la plage et le bord inférieur de la plage. En lecture seule.|
||[left](/javascript/api/excel/excel.rangecollectionloadoptions#left)|Pour chaque élément de la collection: renvoie la distance en points, pour un zoom de 100%, entre le bord gauche de la feuille de calcul et le bord gauche de la plage. En lecture seule.|
||[savedAsArray](/javascript/api/excel/excel.rangecollectionloadoptions#savedasarray)|Pour chaque élément de la collection: représente si toutes les cellules sont enregistrées sous la forme d’une formule matricielle.|
||[top](/javascript/api/excel/excel.rangecollectionloadoptions#top)|Pour chaque élément de la collection: renvoie la distance en points, pour un zoom de 100%, entre le bord supérieur de la feuille de calcul et le bord supérieur de la plage. En lecture seule.|
||[width](/javascript/api/excel/excel.rangecollectionloadoptions#width)|Pour chaque élément de la collection: renvoie la distance en points, pour un zoom de 100%, du bord gauche de la plage au bord droit de la plage. En lecture seule.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[hasSpill](/javascript/api/excel/excel.rangedata#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[height](/javascript/api/excel/excel.rangedata#height)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la plage au bord inférieur de la plage. En lecture seule.|
||[left](/javascript/api/excel/excel.rangedata#left)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la feuille de calcul au bord gauche de la plage. En lecture seule.|
||[savedAsArray](/javascript/api/excel/excel.rangedata#savedasarray)|Représente si toutes les cellules sont enregistrées sous la forme d’une formule matricielle.|
||[top](/javascript/api/excel/excel.rangedata#top)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la feuille de calcul au bord supérieur de la plage. En lecture seule.|
||[width](/javascript/api/excel/excel.rangedata#width)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la plage au bord droit de la plage. En lecture seule.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[hasSpill](/javascript/api/excel/excel.rangeloadoptions#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[height](/javascript/api/excel/excel.rangeloadoptions#height)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la plage au bord inférieur de la plage. En lecture seule.|
||[left](/javascript/api/excel/excel.rangeloadoptions#left)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la feuille de calcul au bord gauche de la plage. En lecture seule.|
||[savedAsArray](/javascript/api/excel/excel.rangeloadoptions#savedasarray)|Représente si toutes les cellules sont enregistrées sous la forme d’une formule matricielle.|
||[top](/javascript/api/excel/excel.rangeloadoptions#top)|Renvoie la distance en points pour zoom 100 %, à partir du bord supérieur de la feuille de calcul au bord supérieur de la plage. En lecture seule.|
||[width](/javascript/api/excel/excel.rangeloadoptions#width)|Renvoie la distance en points pour zoom 100 %, à partir du bord gauche de la plage au bord droit de la plage. En lecture seule.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copie et colle un objet de forme.|
||[placement](/javascript/api/excel/excel.shape#placement)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul. Renvoie un objet Forme qui représente la nouvelle image.|
|[ShapeCollectionLoadOptions](/javascript/api/excel/excel.shapecollectionloadoptions)|[placement](/javascript/api/excel/excel.shapecollectionloadoptions#placement)|Pour chaque élément de la collection: représente la manière dont l’objet est attaché aux cellules qui se trouvent en-dessous.|
|[ShapeData](/javascript/api/excel/excel.shapedata)|[placement](/javascript/api/excel/excel.shapedata#placement)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
|[ShapeLoadOptions](/javascript/api/excel/excel.shapeloadoptions)|[placement](/javascript/api/excel/excel.shapeloadoptions#placement)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
|[ShapeUpdateData](/javascript/api/excel/excel.shapeupdatedata)|[placement](/javascript/api/excel/excel.shapeupdatedata#placement)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
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
||[Set (propriétés: Excel. Slicer)](/javascript/api/excel/excel.slicer#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. SlicerUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.slicer#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
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
|[SlicerCollectionData](/javascript/api/excel/excel.slicercollectiondata)|[items](/javascript/api/excel/excel.slicercollectiondata#items)||
|[SlicerCollectionLoadOptions](/javascript/api/excel/excel.slicercollectionloadoptions)|[$all](/javascript/api/excel/excel.slicercollectionloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicercollectionloadoptions#caption)|Pour chaque élément de la collection: représente la légende du Slicer.|
||[height](/javascript/api/excel/excel.slicercollectionloadoptions#height)|Pour chaque élément de la collection: représente la hauteur, exprimée en points, du segment.|
||[id](/javascript/api/excel/excel.slicercollectionloadoptions#id)|Pour chaque élément de la collection: représente l’ID unique de Slicer. En lecture seule.|
||[isFilterCleared](/javascript/api/excel/excel.slicercollectionloadoptions#isfiltercleared)|Pour chaque élément de la collection: true si tous les filtres actuellement appliqués sur le segment sont effacés.|
||[left](/javascript/api/excel/excel.slicercollectionloadoptions#left)|Pour chaque élément de la collection: représente la distance, en points, entre le côté gauche du Slicer et le bord gauche de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicercollectionloadoptions#name)|Pour chaque élément de la collection: représente le nom du Slicer.|
||[nameInFormula](/javascript/api/excel/excel.slicercollectionloadoptions#nameinformula)|Pour chaque élément de la collection: représente le nom du segment utilisé dans la formule.|
||[sortBy](/javascript/api/excel/excel.slicercollectionloadoptions#sortby)|Pour chaque élément de la collection: représente l’ordre de tri des éléments dans le Slicer. Les valeurs possibles sont : DataSourceOrder par ordre croissant, par ordre décroissant.|
||[style](/javascript/api/excel/excel.slicercollectionloadoptions#style)|Pour chaque élément de la collection: valeur de constante qui représente le style de segment. Les valeurs possibles sont: «SlicerStyleLight1» à «SlicerStyleLight6», «TableStyleOther1» à «TableStyleOther2», «SlicerStyleDark1» et «SlicerStyleDark6». Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
||[top](/javascript/api/excel/excel.slicercollectionloadoptions#top)|Pour chaque élément de la collection: représente la distance, en points, entre le bord supérieur du Slicer et le haut de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicercollectionloadoptions#width)|Pour chaque élément de la collection: représente la largeur, exprimée en points, du segment.|
||[worksheet](/javascript/api/excel/excel.slicercollectionloadoptions#worksheet)|Pour chaque élément de la collection: représente la feuille de calcul contenant le Slicer.|
|[SlicerCollectionUpdateData](/javascript/api/excel/excel.slicercollectionupdatedata)|[items](/javascript/api/excel/excel.slicercollectionupdatedata#items)||
|[SlicerData](/javascript/api/excel/excel.slicerdata)|[caption](/javascript/api/excel/excel.slicerdata#caption)|Représente la légende de segment.|
||[height](/javascript/api/excel/excel.slicerdata#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[id](/javascript/api/excel/excel.slicerdata#id)|Représente l’id unique du segment. En lecture seule.|
||[isFilterCleared](/javascript/api/excel/excel.slicerdata#isfiltercleared)|True si tous les filtres appliqués actuellement sur le segment sont effacés.|
||[left](/javascript/api/excel/excel.slicerdata#left)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicerdata#name)|Représente le nom de la forme.|
||[nameInFormula](/javascript/api/excel/excel.slicerdata#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[slicerItems](/javascript/api/excel/excel.slicerdata#sliceritems)|Représente la collection de SlicerItems qui font partie du segment. En lecture seule.|
||[sortBy](/javascript/api/excel/excel.slicerdata#sortby)|Représente l’ordre de tri des éléments dans le segment. Les valeurs possibles sont : DataSourceOrder par ordre croissant, par ordre décroissant.|
||[style](/javascript/api/excel/excel.slicerdata#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont: «SlicerStyleLight1» à «SlicerStyleLight6», «TableStyleOther1» à «TableStyleOther2», «SlicerStyleDark1» et «SlicerStyleDark6». Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
||[top](/javascript/api/excel/excel.slicerdata#top)|Représente la distance, en points, du bord supérieur de la section à la partie droite de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicerdata#width)|Représente la largeur, en points, de la forme.|
||[worksheet](/javascript/api/excel/excel.slicerdata#worksheet)|Obtenir la feuille de calcul contenant la plage. En lecture seule.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True si l’élément de slicer est sélectionné ; sinon False.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True si l’élément de segment comporte des données.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Représente la valeur unique représentant l’élément de segment.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Représente le titre affiché dans l’interface utilisateur.|
||[Set (propriétés: Excel. SlicerItem)](/javascript/api/excel/excel.sliceritem#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. SlicerItemUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.sliceritem#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Renvoie le nombre de segment les éléments dans le segment.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Obtient un segment objet de l’élément à l’aide de son nom ou clé.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Obtient un segment de l’élément à l’aide de son nom ou clé. Si le paramètre n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerItemCollectionData](/javascript/api/excel/excel.sliceritemcollectiondata)|[items](/javascript/api/excel/excel.sliceritemcollectiondata#items)||
|[SlicerItemCollectionLoadOptions](/javascript/api/excel/excel.sliceritemcollectionloadoptions)|[$all](/javascript/api/excel/excel.sliceritemcollectionloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemcollectionloadoptions#hasdata)|Pour chaque élément de la collection: true si l’élément de Slicer contient des données.|
||[isSelected](/javascript/api/excel/excel.sliceritemcollectionloadoptions#isselected)|Pour chaque élément de la collection: true si l’élément de Slicer est sélectionné.|
||[key](/javascript/api/excel/excel.sliceritemcollectionloadoptions#key)|Pour chaque élément de la collection: représente la valeur unique représentant l’élément de Slicer.|
||[name](/javascript/api/excel/excel.sliceritemcollectionloadoptions#name)|Pour chaque élément de la collection: représente le titre affiché dans l’interface utilisateur.|
|[SlicerItemCollectionUpdateData](/javascript/api/excel/excel.sliceritemcollectionupdatedata)|[items](/javascript/api/excel/excel.sliceritemcollectionupdatedata#items)||
|[SlicerItemData](/javascript/api/excel/excel.sliceritemdata)|[hasData](/javascript/api/excel/excel.sliceritemdata#hasdata)|True si l’élément de segment comporte des données.|
||[isSelected](/javascript/api/excel/excel.sliceritemdata#isselected)|True si l’élément de slicer est sélectionné ; sinon False.|
||[key](/javascript/api/excel/excel.sliceritemdata#key)|Représente la valeur unique représentant l’élément de segment.|
||[name](/javascript/api/excel/excel.sliceritemdata#name)|Représente le titre affiché dans l’interface utilisateur.|
|[SlicerItemLoadOptions](/javascript/api/excel/excel.sliceritemloadoptions)|[$all](/javascript/api/excel/excel.sliceritemloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemloadoptions#hasdata)|True si l’élément de segment comporte des données.|
||[isSelected](/javascript/api/excel/excel.sliceritemloadoptions#isselected)|True si l’élément de slicer est sélectionné ; sinon False.|
||[key](/javascript/api/excel/excel.sliceritemloadoptions#key)|Représente la valeur unique représentant l’élément de segment.|
||[name](/javascript/api/excel/excel.sliceritemloadoptions#name)|Représente le titre affiché dans l’interface utilisateur.|
|[SlicerItemUpdateData](/javascript/api/excel/excel.sliceritemupdatedata)|[isSelected](/javascript/api/excel/excel.sliceritemupdatedata#isselected)|True si l’élément de slicer est sélectionné ; sinon False.|
|[SlicerLoadOptions](/javascript/api/excel/excel.slicerloadoptions)|[$all](/javascript/api/excel/excel.slicerloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicerloadoptions#caption)|Représente la légende de segment.|
||[height](/javascript/api/excel/excel.slicerloadoptions#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[id](/javascript/api/excel/excel.slicerloadoptions#id)|Représente l’id unique du segment. En lecture seule.|
||[isFilterCleared](/javascript/api/excel/excel.slicerloadoptions#isfiltercleared)|True si tous les filtres appliqués actuellement sur le segment sont effacés.|
||[left](/javascript/api/excel/excel.slicerloadoptions#left)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicerloadoptions#name)|Représente le nom de la forme.|
||[nameInFormula](/javascript/api/excel/excel.slicerloadoptions#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[sortBy](/javascript/api/excel/excel.slicerloadoptions#sortby)|Représente l’ordre de tri des éléments dans le segment. Les valeurs possibles sont : DataSourceOrder par ordre croissant, par ordre décroissant.|
||[style](/javascript/api/excel/excel.slicerloadoptions#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont: «SlicerStyleLight1» à «SlicerStyleLight6», «TableStyleOther1» à «TableStyleOther2», «SlicerStyleDark1» et «SlicerStyleDark6». Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
||[top](/javascript/api/excel/excel.slicerloadoptions#top)|Représente la distance, en points, du bord supérieur de la section à la partie droite de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicerloadoptions#width)|Représente la largeur, en points, de la forme.|
||[worksheet](/javascript/api/excel/excel.slicerloadoptions#worksheet)|Obtenir la feuille de calcul contenant la plage.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Supprime le SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Crée un doublon de cette SlicerStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtient le nom de la SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Spécifie si cet objet SlicerStyle est accessible en lecture seule. En lecture seule.|
||[Set (propriétés: Excel. SlicerStyle)](/javascript/api/excel/excel.slicerstyle#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. SlicerStyleUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.slicerstyle#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Crée un SlicerStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Obtient le nombre de styles de slicer de la collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Obtient SlicerStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Obtient un SlicerStyle par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Obtient un SlicerStyle par nom. Si le SlicerStyle n’existe pas, il renvoie un objet null.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Obtient le SlicerStyle par défaut pour la portée de l’objet parent portée.|
|[SlicerStyleCollectionData](/javascript/api/excel/excel.slicerstylecollectiondata)|[items](/javascript/api/excel/excel.slicerstylecollectiondata#items)||
|[SlicerStyleCollectionLoadOptions](/javascript/api/excel/excel.slicerstylecollectionloadoptions)|[$all](/javascript/api/excel/excel.slicerstylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstylecollectionloadoptions#name)|Pour chaque élément de la collection: obtient le nom du SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstylecollectionloadoptions#readonly)|Pour chaque élément de la collection: indique si cet objet SlicerStyle est en lecture seule. En lecture seule.|
|[SlicerStyleCollectionUpdateData](/javascript/api/excel/excel.slicerstylecollectionupdatedata)|[items](/javascript/api/excel/excel.slicerstylecollectionupdatedata#items)||
|[SlicerStyleData](/javascript/api/excel/excel.slicerstyledata)|[name](/javascript/api/excel/excel.slicerstyledata#name)|Obtient le nom de la SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyledata#readonly)|Spécifie si cet objet SlicerStyle est accessible en lecture seule. En lecture seule.|
|[SlicerStyleLoadOptions](/javascript/api/excel/excel.slicerstyleloadoptions)|[$all](/javascript/api/excel/excel.slicerstyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstyleloadoptions#name)|Obtient le nom de la SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyleloadoptions#readonly)|Spécifie si cet objet SlicerStyle est accessible en lecture seule. En lecture seule.|
|[SlicerStyleUpdateData](/javascript/api/excel/excel.slicerstyleupdatedata)|[name](/javascript/api/excel/excel.slicerstyleupdatedata#name)|Obtient le nom de la SlicerStyle.|
|[SlicerUpdateData](/javascript/api/excel/excel.slicerupdatedata)|[caption](/javascript/api/excel/excel.slicerupdatedata#caption)|Représente la légende de segment.|
||[height](/javascript/api/excel/excel.slicerupdatedata#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[left](/javascript/api/excel/excel.slicerupdatedata#left)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicerupdatedata#name)|Représente le nom de la forme.|
||[nameInFormula](/javascript/api/excel/excel.slicerupdatedata#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[sortBy](/javascript/api/excel/excel.slicerupdatedata#sortby)|Représente l’ordre de tri des éléments dans le segment. Les valeurs possibles sont : DataSourceOrder par ordre croissant, par ordre décroissant.|
||[style](/javascript/api/excel/excel.slicerupdatedata#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont: «SlicerStyleLight1» à «SlicerStyleLight6», «TableStyleOther1» à «TableStyleOther2», «SlicerStyleDark1» et «SlicerStyleDark6». Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
||[top](/javascript/api/excel/excel.slicerupdatedata#top)|Représente la distance, en points, du bord supérieur de la section à la partie droite de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicerupdatedata#width)|Représente la largeur, en points, de la forme.|
||[worksheet](/javascript/api/excel/excel.slicerupdatedata#worksheet)|Obtenir la feuille de calcul contenant la plage.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Indique le nom de la table à laquelle le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Représente l’id de la feuille de calcul qui contient le tableau.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Supprime le TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Crée un doublon de cette TableStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtient le nom du TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Spécifie si cet objet TableStyle est accessible en lecture seule. En lecture seule.|
||[Set (propriétés: Excel. TableStyle)](/javascript/api/excel/excel.tablestyle#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. TableStyleUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.tablestyle#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Crée un TableStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Obtient le nombre de styles de tableaux de la collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Obtient le TableStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Extrait un TableStyle par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Extrait un TableStyle par nom. Si le TableStyle n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Définit le TableStyle par défaut pour la portée de l’objet parent portée.|
|[TableStyleCollectionData](/javascript/api/excel/excel.tablestylecollectiondata)|[items](/javascript/api/excel/excel.tablestylecollectiondata#items)||
|[TableStyleCollectionLoadOptions](/javascript/api/excel/excel.tablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestylecollectionloadoptions#name)|Pour chaque élément de la collection: obtient le nom de TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestylecollectionloadoptions#readonly)|Pour chaque élément de la collection: indique si cet objet TableStyle est en lecture seule. En lecture seule.|
|[TableStyleCollectionUpdateData](/javascript/api/excel/excel.tablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.tablestylecollectionupdatedata#items)||
|[TableStyleData](/javascript/api/excel/excel.tablestyledata)|[name](/javascript/api/excel/excel.tablestyledata#name)|Obtient le nom du TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyledata#readonly)|Spécifie si cet objet TableStyle est accessible en lecture seule. En lecture seule.|
|[TableStyleLoadOptions](/javascript/api/excel/excel.tablestyleloadoptions)|[$all](/javascript/api/excel/excel.tablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestyleloadoptions#name)|Obtient le nom du TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyleloadoptions#readonly)|Spécifie si cet objet TableStyle est accessible en lecture seule. En lecture seule.|
|[TableStyleUpdateData](/javascript/api/excel/excel.tablestyleupdatedata)|[name](/javascript/api/excel/excel.tablestyleupdatedata#name)|Obtient le nom du TableStyle.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Supprime le TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Crée un doublon de cette TimelineStyle avec une copie de tous les éléments de style.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtient le nom du TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Spécifie si cet objet TimelineStyle est accessible en lecture seule. En lecture seule.|
||[Set (propriétés: Excel. TimelineStyle)](/javascript/api/excel/excel.timelinestyle#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. TimelineStyleUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.timelinestyle#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Crée un TimelineStyle vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Obtient le nombre de styles de délai de la collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Obtient le TimelineStyle par défaut pour la portée de l’objet parent portée.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Obtient un TimelineStyle par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Obtient un TimelineStyle par nom. Si leTimelineStyle n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Définit le TimelineStyle par défaut pour la portée de l’objet parent portée.|
|[TimelineStyleCollectionData](/javascript/api/excel/excel.timelinestylecollectiondata)|[items](/javascript/api/excel/excel.timelinestylecollectiondata#items)||
|[TimelineStyleCollectionLoadOptions](/javascript/api/excel/excel.timelinestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.timelinestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestylecollectionloadoptions#name)|Pour chaque élément de la collection: obtient le nom du TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestylecollectionloadoptions#readonly)|Pour chaque élément de la collection: indique si cet objet TimelineStyle est en lecture seule. En lecture seule.|
|[TimelineStyleCollectionUpdateData](/javascript/api/excel/excel.timelinestylecollectionupdatedata)|[items](/javascript/api/excel/excel.timelinestylecollectionupdatedata#items)||
|[TimelineStyleData](/javascript/api/excel/excel.timelinestyledata)|[name](/javascript/api/excel/excel.timelinestyledata#name)|Obtient le nom du TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyledata#readonly)|Spécifie si cet objet TimelineStyle est accessible en lecture seule. En lecture seule.|
|[TimelineStyleLoadOptions](/javascript/api/excel/excel.timelinestyleloadoptions)|[$all](/javascript/api/excel/excel.timelinestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestyleloadoptions#name)|Obtient le nom du TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyleloadoptions#readonly)|Spécifie si cet objet TimelineStyle est accessible en lecture seule. En lecture seule.|
|[TimelineStyleUpdateData](/javascript/api/excel/excel.timelinestyleupdatedata)|[name](/javascript/api/excel/excel.timelinestyleupdatedata#name)|Obtient le nom du TimelineStyle.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtient le segment actif actuel du classeur. S’il n’y a aucun segment actif, `ItemNotFound` une exception est générée.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtient le segment actif actuel du classeur. S’il n’existe aucun segment actif, un objet null est renvoyé.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Représente une collection de styles associés au classeur. En lecture seule.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Représente une collection de PivotTableStyles associée au classeur. En lecture seule.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Représente une collection de styles associés au classeur. En lecture seule.|
||[Slicers](/javascript/api/excel/excel.workbook#slicers)|Représente une collection de styles associés au classeur. En lecture seule.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Représente une collection de TableStyles associés au classeur. En lecture seule.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Représente une collection de TimelineStyles associés au classeur. En lecture seule.|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[comments](/javascript/api/excel/excel.workbookdata#comments)|Représente une collection de styles associés au classeur. En lecture seule.|
||[pivotTableStyles](/javascript/api/excel/excel.workbookdata#pivottablestyles)|Représente une collection de PivotTableStyles associée au classeur. En lecture seule.|
||[slicerStyles](/javascript/api/excel/excel.workbookdata#slicerstyles)|Représente une collection de styles associés au classeur. En lecture seule.|
||[Slicers](/javascript/api/excel/excel.workbookdata#slicers)|Représente une collection de styles associés au classeur. En lecture seule.|
||[tableStyles](/javascript/api/excel/excel.workbookdata#tablestyles)|Représente une collection de TableStyles associés au classeur. En lecture seule.|
||[timelineStyles](/javascript/api/excel/excel.workbookdata#timelinestyles)|Représente une collection de TimelineStyles associés au classeur. En lecture seule.|
||[use1904DateSystem](/javascript/api/excel/excel.workbookdata#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[use1904DateSystem](/javascript/api/excel/excel.workbookloadoptions#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[use1904DateSystem](/javascript/api/excel/excel.workbookupdatedata#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Renvoie une collection de tous les objets Lecteur sur l’ordinateur. En lecture seule.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Se produit lors du tri des colonnes.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Se produit lorsque l’état de la ligne masquée d’une feuille de calcul spécifique est modifié.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Se produit lors du tri des lignes.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Se produit lorsque l’opération clic gauche/tape se produit dans la feuille de calcul.|
||[Slicers](/javascript/api/excel/excel.worksheet#slicers)|Renvoie une collection de graphiques qui font partie de la feuille de calcul. En lecture seule.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64 (base64File : chaîne, sheetNamesToInsert ? : [] chaîne, positionType ? : « None » \| « Avant les caractères » \| « Après » \| « Depuis » \| « Fin », relativeTo ? : feuille de calcul \| chaîne)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Se produit lors du tri des colonnes.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Cet événement se produit lorsqu’un état masqué d’une feuille de calcul du classeur a été modifié.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Se produit lors du tri des lignes.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Cet événement se produit lorsque l’utilisateur clique dessus ou a cliqué sur une opération dans la collection Worksheet.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le tri s’est passé.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[comments](/javascript/api/excel/excel.worksheetdata#comments)|Renvoie une collection de tous les objets Lecteur sur l’ordinateur. En lecture seule.|
||[Slicers](/javascript/api/excel/excel.worksheetdata#slicers)|Renvoie une collection de graphiques qui font partie de la feuille de calcul. En lecture seule.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Indique le nom du tableau auquel le filtre est appliqué.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont est déclenché l’événement modifié. Pour plus d’informations, voir Excel. RowHiddenChangeType.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le tri s’est passé.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[adresse](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtient l’adresse qui représente la cellule sur laquelle vous avez fait un clic gauche/appuyé pour une feuille de calcul spécifique.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|La distance en points à partir du point sur lequel vous avez effectué un clic gauche/avez appuyé vers le bord gauche (droite pour DÀG) du bord de la grille de la cellule clic gauche/tape.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|La distance en points à partir du point sur lequel vous avez effectué un clic gauche/avez appuyé vers le bord supérieur du bord de la grille de la cellule clic gauche/tape.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la cellule a été sélectionnée par clic-gauche/tape.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
