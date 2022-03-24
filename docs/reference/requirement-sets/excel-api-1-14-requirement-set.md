---
title: Excel conditions requises de l’API JavaScript 1.14
description: Détails sur l’ensemble de conditions requises ExcelApi 1.14.
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 93b1690a3c03e51dadb2110ec6382ca6ee86cfe1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747016"
---
# <a name="whats-new-in-excel-javascript-api-114"></a>Nouveautés de l’API JavaScript Excel 1.14

ExcelApi 1.14 a ajouté des objets pour contrôler la fonctionnalité de table de données d’un graphique, une méthode pour localiser toutes les cellules précédentes d’une formule et des événements de protection de feuille de calcul pour suivre les modifications apportées à l’état de protection d’une feuille de calcul. Il a également ajouté plusieurs méthodes [`getItemOrNullObject`](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties) pour des objets tels `CommentCollection`que , et `ShapeCollection`pour `StyleCollection` améliorer la gestion des erreurs.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Tables de données de graphique](../../excel/excel-add-ins-charts.md#add-and-format-a-chart-data-table) | Contrôler l’apparence, la mise en forme et la visibilité des tables de données sur les graphiques. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| [Antécédents de formule](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-precedents-of-a-formula) | Renvoyer toutes les cellules précédentes d’une formule. | [Range](/javascript/api/excel/excel.range) |
| Requêtes | Récupérer les attributs Power Query tels que le nom, la date d’actualisation et le nombre de requêtes. | [Requête](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| [Événements de protection de feuille de calcul](../../excel/excel-add-ins-worksheets.md#detect-changes-to-the-worksheet-protection-state) | Suivre les modifications apportées à l’état de protection d’une feuille de calcul et à la source de ces modifications. | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [Worksheet](/javascript/api/excel/excel.worksheet), [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.14. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.14 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.14](/javascript/api/excel?view=excel-js-1.14&preserve-view=true) ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcolumncriteria-member(1))|Cette fonction permet d’effacer les critères de filtrage des colonnes du filtre automatique.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-deleteshiftdirection-member)|Représente la direction (par exemple, vers le haut ou vers la gauche) vers le haut ou vers la gauche que les cellules restantes déplacent lorsqu’une ou plusieurs cellules sont supprimées.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-insertshiftdirection-member)|Représente la direction (par exemple, vers le bas ou vers la droite) vers le bas ou vers la droite que les cellules existantes déplacent lorsqu’une ou plusieurs nouvelles cellules sont insérées.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatable-member(1))|Obtient la table de données du graphique.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatableornullobject-member(1))|Obtient la table de données du graphique.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-format-member)|Représente le format d’un tableau de données de graphique, qui inclut le format de remplissage, de police et de bordure.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showhorizontalborder-member)|Spécifie s’il faut afficher la bordure horizontale de la table de données.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showlegendkey-member)|Spécifie s’il faut afficher le clé de légende de la table de données.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showoutlineborder-member)|Spécifie s’il faut afficher la bordure de plan de la table de données.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showverticalborder-member)|Spécifie s’il faut afficher la bordure verticale de la table de données.|
||[visible](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-visible-member)|Spécifie s’il faut afficher la table de données du graphique.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[bordure](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-border-member)|Représente le format de bordure du tableau de données de graphique, qui inclut la couleur, le style de trait et l’pondération.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-fill-member)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan.|
||[police](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-font-member)|Représente les attributs de police (tels que le nom de la police, la taille de police et la couleur) de l’objet actuel.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemornullobject-member(1))|Obtient un commentaire à partir de la collection de sites en fonction de son ID.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemornullobject-member(1))|Renvoie une réponse de commentaire identifié via son ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitemornullobject-member(1))|Renvoie un format conditionnel identifié par son ID.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemornullobject-member(1))|Obtient une forme à l’aide de son nom ou de son ID.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#excel-excel-query-error-member)|Obtient le message d’erreur de requête à partir de la dernière actualisation de la requête.|
||[loadedTo](/javascript/api/excel/excel.query#excel-excel-query-loadedto-member)|Obtient la requête chargée dans le type d’objet.|
||[loadedToDataModel](/javascript/api/excel/excel.query#excel-excel-query-loadedtodatamodel-member)|Spécifie si la requête a été chargée dans le modèle de données.|
||[name](/javascript/api/excel/excel.query#excel-excel-query-name-member)|Obtient le nom de la requête.|
||[refreshDate](/javascript/api/excel/excel.query#excel-excel-query-refreshdate-member)|Obtient la date et l’heure de la dernière actualisation de la requête.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#excel-excel-query-rowsloadedcount-member)|Obtient le nombre de lignes qui ont été chargées lors de la dernière actualisation de la requête.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getcount-member(1))|Obtient le nombre de requêtes dans le manuel.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getitem-member(1))|Obtient une requête de la collection en fonction de son nom.|
||[items](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1))|Renvoie un `WorkbookRangeAreas` objet qui représente la plage contenant tous les antécédents d’une cellule dans la même feuille de calcul ou dans plusieurs feuilles de calcul.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemornullobject-member(1))|Obtient une forme à l’aide de son nom ou de son ID.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemornullobject-member(1))|Obtient un style par nom.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitemornullobject-member(1))|Obtient un tableau à l’aide de son nom ou de son ID.|
|[Workbook](/javascript/api/excel/excel.workbook)|[requêtes](/javascript/api/excel/excel.workbook#excel-excel-workbook-queries-member)|Renvoie une collection de requêtes Power Query qui font partie du manuel.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onProtectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
||[tabId](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabid-member)|Renvoie une valeur représentant cette feuille de calcul qui peut être lue par Open Office XML.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changedirectionstate-member)|Représente une modification du sens de déplacement des cellules d’une feuille de calcul lorsqu’une ou plusieurs cellules sont supprimées ou insérées.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-triggersource-member)|Représente la source du déclencheur de l’événement.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-isprotected-member)|Obtient l’état de protection actuel de la feuille de calcul.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-source-member)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle l’état de protection est modifié.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
