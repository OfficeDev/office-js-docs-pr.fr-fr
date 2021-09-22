---
title: Excel Ensemble de conditions requises de l’API JavaScript en ligne uniquement
description: Détails sur l’ensemble de conditions requises ExcelApiOnline.
ms.date: 09/16/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9b8d326e1a756a873fc19b3d78f795ebf04e5f4e
ms.sourcegitcommit: a854a2fd2ad9f379a3ef712f307e0b1bb9b5b00d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/22/2021
ms.locfileid: "59474335"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel Ensemble de conditions requises de l’API JavaScript en ligne uniquement

L’ensemble de conditions requises est un ensemble de conditions requises spécial qui inclut des fonctionnalités qui ne sont disponibles que `ExcelApiOnline` pour Excel sur le Web. Les API de cet ensemble de conditions requises sont considérées comme des API de production (non sujettes à des modifications comportementales ou structurelles nondocumentées) pour l’application Excel sur le Web de production. `ExcelApiOnline`Les API sont considérées comme des API « d’aperçu » pour d’autres plateformes (Windows, Mac, iOS) et peuvent ne pas être pris en charge par l’une de ces plateformes.

Lorsque les API de l’ensemble de conditions requises sont pris en charge sur toutes les plateformes, elles sont ajoutées à l’ensemble de conditions requises `ExcelApiOnline` publié suivant ( `ExcelApi 1.[NEXT]` ). Une fois que cette nouvelle exigence est publique, ces API sont supprimées de `ExcelApiOnline` . Il s’agit d’un processus de promotion similaire à une API passant de la version d’évaluation à la publication.

> [!IMPORTANT]
> `ExcelApiOnline` est un sur-ensemble de l’ensemble de conditions requises numérotées le plus récent.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` est la seule version des API en ligne uniquement. Cela est dû au Excel sur le Web’une seule version disponible pour les utilisateurs qui est la dernière version.

Le tableau suivant fournit un résumé concis des API, tandis que le tableau de liste [d’API](#api-list) suivant fournit une liste détaillée des `ExcelApiOnline` API actuelles.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Workbooks liés | Gérer les liens entre les workbooks, y compris la prise en charge de l’actualisation et de la rupture des liens de ces derniers. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Vues de feuille nommée | Permet de contrôler par programme les affichages de feuille de calcul par utilisateur. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |

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
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|Cette fonction permet d’effacer les critères de filtrage des colonnes du filtre automatique.|
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
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Renvoie une collection de workbooks liés.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Renvoie une collection d’affichages de feuille présents dans la feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Version d’évaluation API JavaScript Excel](excel-preview-apis.md)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
