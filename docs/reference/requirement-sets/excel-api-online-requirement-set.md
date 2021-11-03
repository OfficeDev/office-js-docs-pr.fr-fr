---
title: Excel Ensemble de conditions requises de l’API JavaScript en ligne uniquement
description: Détails sur l’ensemble de conditions requises ExcelApiOnline.
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ae338b6bd361113ee04ae3dd9076df6c66125345
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681492"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel Ensemble de conditions requises de l’API JavaScript en ligne uniquement

L’ensemble de conditions requises est un ensemble de conditions requises spécial qui inclut des fonctionnalités qui ne sont disponibles que `ExcelApiOnline` pour Excel sur le Web. Les API de cet ensemble de conditions requises sont considérées comme des API de production (non sujettes à des modifications comportementales ou structurelles nondocumentées) pour l Excel sur le Web application. `ExcelApiOnline`Les API sont considérées comme des API « d’aperçu » pour d’autres plateformes (Windows, Mac, iOS) et peuvent ne pas être pris en charge par l’une de ces plateformes.

Lorsque les API de l’ensemble de conditions requises sont pris en charge sur toutes les plateformes, elles sont ajoutées à l’ensemble de conditions requises `ExcelApiOnline` publié suivant ( `ExcelApi 1.[NEXT]` ). Une fois que cette nouvelle exigence est publique, ces API sont supprimées de `ExcelApiOnline` . Il s’agit d’un processus de promotion similaire à une API passant de la version d’évaluation à la publication.

> [!IMPORTANT]
> `ExcelApiOnline` est un sur-ensemble de l’ensemble de conditions requises numérotées le plus récent.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` est la seule version des API en ligne uniquement. Cela est dû au Excel sur le Web’une seule version disponible pour les utilisateurs qui est la dernière version.

Le tableau suivant fournit un résumé concis des API, tandis que le tableau de liste [d’API](#api-list) suivant fournit une liste détaillée des `ExcelApiOnline` API actuelles.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Workbooks liés | Gérez les liens entre les workbooks, notamment la prise en charge de l’actualisation et de la rupture des liens de ces derniers. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Vues de feuille nommée | Permet de contrôler par programme les affichages de feuille de calcul par utilisateur. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| Événements de déplacement de feuille de calcul | Détecter le moment où les feuilles de calcul sont déplacées dans une collection, la position de la feuille de calcul et la source de la modification. | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection), [WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |

## <a name="recommended-usage"></a>Utilisation recommandée

Étant donné que les API sont uniquement Excel sur le Web, votre add-in doit vérifier si l’ensemble de conditions requises est pris en charge avant d’appeler `ExcelApiOnline` ces API. Cela évite d’appeler une API en ligne uniquement sur une plateforme différente.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Une fois que l’API se trouve dans un ensemble de conditions requises sur plusieurs plateformes, vous devez supprimer ou modifier la `isSetSupported` vérification. Cela activera la fonctionnalité de votre add-in sur d’autres plateformes. N’oubliez pas de tester la fonctionnalité sur ces plateformes lors de cette modification.

> [!IMPORTANT]
> Votre manifeste ne peut pas `ExcelApiOnline 1.1` spécifier comme condition d’activation. Il ne s’agit pas d’une valeur valide à utiliser dans [l’élément Set](../manifest/set.md).

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie Excel api JavaScript actuellement incluses dans l’ensemble `ExcelApiOnline` de conditions requises. Pour obtenir la liste complète de toutes les API JavaScript Excel (y compris les API et les API publiées précédemment), consultez toutes les API `ExcelApiOnline` [JavaScript Excel.](/javascript/api/excel?view=excel-js-online&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|Effectue une demande pour rompre les liens pointant vers le workbook lié.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|URL d’origine pointant vers le workbook lié.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|Effectue une demande d’actualisation des données récupérées à partir du workbook lié.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|Rompt tous les liens vers les workbooks liés.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|Obtient des informations sur un workbook lié par son URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|Obtient des informations sur un workbook lié par son URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|Effectue une demande d’actualisation de tous les liens dubook.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|Représente le mode de mise à jour des liens du workbook.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate__)|Active cette vue de feuille.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete__)|Supprime l’affichage Feuille de la feuille de calcul.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate_name_)|Crée une copie de cette vue de feuille.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtient ou définit le nom de l’affichage Feuille.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add_name_)|Crée un affichage feuille avec le nom donné.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#enterTemporary__)|Crée et active un nouvel affichage de feuille temporaire.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit__)|Quitte l’affichage feuille actif.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getActive__)|Obtient la vue de feuille de calcul active.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getCount__)|Obtient le nombre d’affichages de feuille dans cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItem_key_)|Obtient une vue de feuille à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getItemAt_index_)|Obtient une vue de feuille par son index dans la collection.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|Supprimez plusieurs lignes d’un tableau.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|Supprimez un nombre spécifié de lignes d’un tableau, en commençant à un index donné.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Renvoie une collection de workbooks liés.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Renvoie une collection d’affichages de feuille présents dans la feuille de calcul.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#onNameChanged)|Se produit lorsque le nom de la feuille de calcul est modifié.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#onVisibilityChanged)|Se produit lorsque la visibilité de la feuille de calcul est modifiée.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#onMoved)|Se produit lorsqu’une feuille de calcul est déplacée par un utilisateur dans un workbook.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#onNameChanged)|Se produit lorsque le nom de la feuille de calcul est modifié dans la collection de feuilles de calcul.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#onVisibilityChanged)|Se produit lorsque la visibilité de la feuille de calcul est modifiée dans la collection de feuilles de calcul.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#positionAfter)|Obtient la nouvelle position de la feuille de calcul, après le déplacement.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#positionBefore)|Obtient la position précédente de la feuille de calcul, avant le déplacement.|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul qui a été déplacée.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameAfter)|Obtient le nouveau nom de la feuille de calcul, après la modification du nom.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameBefore)|Obtient le nom précédent de la feuille de calcul, avant que le nom ne soit modifié.|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul avec le nouveau nom.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#type)|Obtient le type de l’événement.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityAfter)|Obtient le nouveau paramètre de visibilité de la feuille de calcul, après la modification de la visibilité.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityBefore)|Obtient le paramètre de visibilité précédent de la feuille de calcul, avant la modification de la visibilité.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dont la visibilité a changé.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Version d’évaluation API JavaScript Excel](excel-preview-apis.md)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
