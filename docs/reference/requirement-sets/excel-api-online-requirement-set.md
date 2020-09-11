---
title: Ensemble de conditions requises de l’API JavaScript pour Excel en ligne uniquement
description: Détails sur l’ensemble de conditions requises pour ExcelApiOnline
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: aa497ff97533ff3a414905547a949fa8430c3efe
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430813"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Ensemble de conditions requises de l’API JavaScript pour Excel en ligne uniquement

L' `ExcelApiOnline` ensemble de conditions requises est un ensemble de conditions requises spéciales qui inclut des fonctionnalités qui sont disponibles uniquement pour Excel sur le Web. Les API de cet ensemble de conditions requises sont considérées comme des API de production (non soumises à des modifications structurelles ou comportementales non documentées) pour l’application Excel sur le Web. `ExcelApiOnline` sont considérés comme des API de « préversion » pour les autres plateformes (Windows, Mac, iOS) et ne sont peut-être pas pris en charge par aucune de ces plateformes.

Lorsque les API dans l' `ExcelApiOnline` ensemble de conditions requises sont prises en charge sur toutes les plateformes, elles seront ajoutées à l’ensemble de conditions requises publié suivant ( `ExcelApi 1.[NEXT]` ). Une fois que cette nouvelle exigence est publique, ces API seront supprimées de `ExcelApiOnline` . Imaginez qu’il s’agit d’un processus de promotion similaire, qui passe de l’aperçu à la version Release.

> [!IMPORTANT]
> `ExcelApiOnline` est un sur-ensemble du jeu de conditions requises le plus récent.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` est la seule version des API en ligne uniquement. En effet, Excel sur le Web disposera toujours d’une seule version disponible pour les utilisateurs qui est la version la plus récente.

## <a name="recommended-usage"></a>Utilisation recommandée

Étant donné que `ExcelApiOnline` les API sont uniquement prises en charge par Excel sur le Web, votre complément doit vérifier si l’ensemble de conditions requises est pris en charge avant d’appeler ces API. Cela évite d’appeler une API en ligne uniquement sur une autre plateforme.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Une fois que l’API se trouve dans un ensemble de conditions requises entre plateformes, vous devez supprimer ou modifier la `isSetSupported` vérification. Cette opération active la fonctionnalité de votre complément sur d’autres plateformes. Veillez à tester la fonctionnalité sur ces plateformes lors de l’exécution de cette modification.

> [!IMPORTANT]
> Votre manifeste ne peut pas spécifier `ExcelApiOnline 1.1` comme condition d’activation. Il ne s’agit pas d’une valeur valide à utiliser dans l' [élément Set](../manifest/set.md).

## <a name="api-list"></a>Liste des API

Les API suivantes sont actuellement disponibles pour Excel sur le Web dans le cadre de l' `ExcelApiOnline 1.1` ensemble de conditions requises.

| Class | Champs | Description |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Cette énumération spécifie l’angle auquel le texte est orienté pour le titre de l’axe du graphique. La valeur doit être un entier compris entre-90 et 90 ou l’entier 180 pour le texte orienté verticalement.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtient le nombre de tableaux croisés dynamiques dans la collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtient le premier tableau croisé dynamique de la collection. Les tableaux croisés dynamiques de la collection sont triés de haut en bas et de gauche à droite, de sorte que le tableau supérieur gauche est le premier tableau croisé dynamique de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtient un tableau croisé dynamique par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/excel/excel.range)|[getPivotTables (fullyContained ?: booléen)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtient une collection d’étendues de tableaux croisés dynamiques qui se chevauchent avec la plage.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Version d’évaluation API JavaScript Excel](./excel-preview-apis.md)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)