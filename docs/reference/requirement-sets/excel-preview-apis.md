---
title: Version d’évaluation API JavaScript Excel
description: Informations détaillées sur les API JavaScript pour Excel à venir
ms.date: 10/22/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: dc0a2a3b23fbf4ccffb5de3b0689b0de0ed08b75
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682542"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Mentions de commentaires](../../excel/excel-add-ins-comments.md#mentions-preview) | Indiquez d’autres personnes dans les commentaires pour envoyer des notifications. | [Commentaire](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| [Insérer le classeur](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insérer un classeur dans un autre.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Classeur [enregistrer](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) et [fermer](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | Enregistrez et fermez ses classeurs.  | [Workbook](/javascript/api/excel/excel.workbook) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Excel actuellement en version préliminaire. Pour afficher la liste complète de toutes les API JavaScript pour Excel (y compris les API d’aperçu et les API précédemment publiées), voir [toutes les API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview).

| Class | Champs | Description |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension : Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtient les valeurs d’une dimension unique de la série de graphiques. Il peut s’agir de valeurs de catégorie ou de valeurs de données, en fonction de la dimension spécifiée et de la façon dont les données sont mappées pour la série de graphiques.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Obtient les entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Obtient le contenu de commentaire enrichi (par exemple, les mentions dans les commentaires). Cette chaîne n’est pas destinée à être affichée aux utilisateurs finaux. Votre complément doit uniquement l’utiliser pour analyser le contenu de commentaire enrichi.|
||[évaluation](/javascript/api/excel/excel.comment#resolved)|Obtient ou définit l’état du thread de commentaire. La valeur « true » signifie que le thread de commentaire est dans l’État résolu.|
||[updateMentions (contentWithMentions : Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Met à jour le contenu de commentaire avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Obtient ou définit l’adresse de messagerie de l’entité mentionnée dans Comment.|
||[id](/javascript/api/excel/excel.commentmention#id)|Obtient ou définit l’ID de l’entité. Elle est alignée sur les informations d' `CommentRichContent.richContent`ID dans.|
||[name](/javascript/api/excel/excel.commentmention#name)|Obtient ou définit le nom de l’entité mentionnée dans Comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Obtient les entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[évaluation](/javascript/api/excel/excel.commentreply#resolved)|Obtient ou définit l’état de la réponse de commentaire. La valeur « true » signifie que la réponse au commentaire est dans l’État résolu.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Obtient le contenu de commentaire enrichi (par exemple, les mentions dans les commentaires). Cette chaîne n’est pas destinée à être affichée aux utilisateurs finaux. Votre complément doit uniquement l’utiliser pour analyser le contenu de commentaire enrichi.|
||[updateMentions (contentWithMentions : Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Met à jour le contenu de commentaire avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Tableau contenant toutes les entités (par exemple, les personnes) mentionnées dans le commentaire.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives. La cellule renvoyée est l’intersection de la ligne donnée et une colonne qui contient les données à partir de la hiérarchie donnée. Cette méthode est l’inverse de l’appel getPivotItems et getDataHierarchy sur une cellule particulière.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. Échoue si appliqué à une plage comportant plusieurs cellules. Lecture seule.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. Lecture seule.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. Échoue si appliqué à une plage comportant plusieurs cellules. Lecture seule.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. En lecture seule.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Représente si toutes les cellules sont enregistrées sous la forme d’une formule matricielle.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (montant : nombre)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Ajuste la mise en retrait de la plage de mise en forme. La valeur de retrait est comprise entre 0 et 250.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul. Renvoie un objet Forme qui représente la nouvelle image.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Représente l’ID de la table dans laquelle le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Représente l’id de la feuille de calcul qui contient le tableau.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Indique le nom du tableau auquel le filtre est appliqué.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont l’événement a été déclenché. Pour `Excel.RowHiddenChangeType` plus d’informations, voir.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
