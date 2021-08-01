---
title: Excel Ensemble de conditions requises de l’API JavaScript en ligne uniquement
description: Détails sur l’ensemble de conditions requises ExcelApiOnline.
ms.date: 07/23/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 8e49d2a9d655e7bc1bf21233bdaf23321d1e48d1
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671876"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel Ensemble de conditions requises de l’API JavaScript en ligne uniquement

L’ensemble de conditions requises est un ensemble de conditions requises spécial qui inclut des fonctionnalités qui ne sont disponibles que `ExcelApiOnline` pour Excel sur le Web. Les API de cet ensemble de conditions requises sont considérées comme des API de production (non sujettes à des modifications comportementales ou structurelles nondocumentées) pour l Excel sur le Web application. `ExcelApiOnline`Les API sont considérées comme des API « d’aperçu » pour d’autres plateformes (Windows, Mac, iOS) et peuvent ne pas être pris en charge par l’une de ces plateformes.

Lorsque les API de l’ensemble de conditions requises sont pris en charge sur toutes les plateformes, elles sont ajoutées à l’ensemble de conditions requises `ExcelApiOnline` publié suivant ( `ExcelApi 1.[NEXT]` ). Une fois que cette nouvelle exigence est publique, ces API sont supprimées de `ExcelApiOnline` . Il s’agit d’un processus de promotion similaire à une API passant de la version d’évaluation à la publication.

> [!IMPORTANT]
> `ExcelApiOnline` est un sur-ensemble de l’ensemble de conditions requises numérodé le plus récent.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` est la seule version des API en ligne uniquement. Cela est dû au Excel sur le Web’une seule version disponible pour les utilisateurs qui est la dernière version.

Le tableau suivant fournit un résumé concis des API, tandis que le tableau de liste [d’API](#api-list) suivant fournit une liste détaillée des `ExcelApiOnline` API actuelles.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Vues de feuille nommée | Permet de contrôler par programme les affichages de feuille de calcul par utilisateur. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

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

Le tableau suivant répertorie les Excel api JavaScript actuellement incluses dans l’ensemble `ExcelApiOnline` de conditions requises. Pour obtenir la liste complète de toutes les API JavaScript Excel (y compris les API et les API publiées précédemment), consultez toutes les API `ExcelApiOnline` [JavaScript Excel.](/javascript/api/excel?view=excel-js-online&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|Cette fonction permet d’effacer les critères de filtrage des colonnes du filtre automatique.|
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
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Renvoie une collection d’affichages de feuille présents dans la feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Version d’évaluation API JavaScript Excel](excel-preview-apis.md)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
