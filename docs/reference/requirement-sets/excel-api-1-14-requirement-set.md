---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.14
description: Détails sur l’ensemble de conditions requises ExcelApi 1.14.
ms.date: 10/13/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 3b968cf40fa35921df1be0aca508041e72cddf23
ms.sourcegitcommit: e4d98eb90e516b9c90e3832f3212caf48691acf6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/22/2021
ms.locfileid: "60537713"
---
# <a name="whats-new-in-excel-javascript-api-114"></a>Nouveautés de l Excel API JavaScript 1.14

ExcelApi 1.14 a ajouté des objets pour contrôler la fonctionnalité de table de données d’un graphique, une méthode pour localiser toutes les cellules précédentes d’une formule et des événements de protection de feuille de calcul pour suivre les modifications apportées à l’état de protection d’une feuille de calcul. Il a également ajouté plusieurs méthodes pour des objets tels que , et pour améliorer [`getItemOrNullObject`](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties) `CommentCollection` la gestion des `ShapeCollection` `StyleCollection` erreurs.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Tables de données de graphique | Contrôler l’apparence, la mise en forme et la visibilité des tables de données sur les graphiques. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Antécédents de formule | Renvoyer toutes les cellules précédentes d’une formule. | [Range](/javascript/api/excel/excel.range) |
| Requêtes | Récupérer des attributs Power Query tels que le nom, la date d’actualisation et le nombre de requêtes. | [Query](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| Événements de protection de feuille de calcul | Suivre les modifications apportées à l’état de protection d’une feuille de calcul et à la source de ces modifications. | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [Worksheet](/javascript/api/excel/excel.worksheet), [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.14. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.14 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.14](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Représente la direction (par exemple, vers le haut ou vers la gauche) vers le haut ou vers la gauche que les cellules restantes déplacent lorsqu’une ou plusieurs cellules sont supprimées.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Représente la direction (par exemple, vers le bas ou vers la droite) que les cellules existantes déplacent lorsqu’une ou plusieurs nouvelles cellules sont insérées.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Obtient la table de données du graphique.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Obtient la table de données du graphique.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Représente le format d’un tableau de données de graphique, qui inclut le format de remplissage, de police et de bordure.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Spécifie s’il faut afficher la bordure horizontale de la table de données.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Spécifie s’il faut afficher le clé de légende de la table de données.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Spécifie s’il faut afficher la bordure de plan de la table de données.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Spécifie s’il faut afficher la bordure verticale de la table de données.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Spécifie s’il faut afficher la table de données du graphique.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[bordure](/javascript/api/excel/excel.chartdatatableformat#border)|Représente le format de bordure de la table de données du graphique, qui inclut la couleur, le style de trait et l’pondération.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan.|
||[police](/javascript/api/excel/excel.chartdatatableformat#font)|Représente les attributs de police (tels que le nom de la police, la taille de police et la couleur) de l’objet actuel.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Obtient un commentaire à partir de la collection de sites en fonction de son ID.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Renvoie une réponse de commentaire identifié via son ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Renvoie un format conditionnel identifié par son ID.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|Obtient une forme à l’aide de son nom ou de son ID.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Obtient le message d’erreur de requête à partir de la dernière actualisation de la requête.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Obtient la requête chargée dans le type d’objet.|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Spécifie si la requête a été chargée dans le modèle de données.|
||[name](/javascript/api/excel/excel.query#name)|Obtient le nom de la requête.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Obtient la date et l’heure de la dernière actualisation de la requête.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Obtient le nombre de lignes qui ont été chargées lors de la dernière actualisation de la requête.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Obtient le nombre de requêtes dans le workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Obtient une requête de la collection en fonction de son nom.|
||[items](/javascript/api/excel/excel.querycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Renvoie un objet qui représente la plage contenant tous les antécédents d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Obtient une forme à l’aide de son nom ou de son ID.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|Obtient un style par nom.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|Obtient un tableau à l’aide de son nom ou de son ID.|
|[Workbook](/javascript/api/excel/excel.workbook)|[requêtes](/javascript/api/excel/excel.workbook#queries)|Renvoie une collection de requêtes Power Query qui font partie du manuel.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Renvoie une valeur représentant cette feuille de calcul qui peut être lue par Open Office XML.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Représente une modification du sens de déplacement des cellules d’une feuille de calcul lorsqu’une ou plusieurs cellules sont supprimées ou insérées.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Représente la source du déclencheur de l’événement.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Obtient l’état de protection actuel de la feuille de calcul.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle l’état de protection est modifié.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
