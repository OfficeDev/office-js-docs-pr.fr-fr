---
title: Ensemble de conditions requises de l’API JavaScript pour Excel en ligne uniquement
description: Détails sur l’ensemble de conditions requises pour ExcelApiOnline.
ms.date: 09/15/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 29f5826ba2adbf18b79033b83254b046210015fe
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819804"
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

Il n’existe actuellement aucune API dans l' `ExcelApiOnline` ensemble de conditions requises. Toutes les API qui faisaient auparavant partie de cet ensemble ont été graduées en un ensemble de conditions requises et sont disponibles sur toutes les plateformes.

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Version d’évaluation API JavaScript Excel](excel-preview-apis.md)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
