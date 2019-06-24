---
ms.date: 06/18/2019
description: Gérez les erreurs dans vos fonctions personnalisées Excel.
title: Gestion des erreurs liées aux fonctions personnalisées dans Excel
localization_priority: Priority
ms.openlocfilehash: 3818d33121ed26bb7d65c56bf6c504f2fb049c72
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127918"
---
# <a name="error-handling-within-custom-functions"></a>Gestion des erreurs dans des fonctions personnalisées

Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.

```js
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [résoudre les problèmes liés à vos fonctions personnalisées](custom-functions-troubleshooting.md).

## <a name="see-also"></a>Voir aussi

* [Débogage des fonctions personnalisées](custom-functions-debugging.md)
* [Configuration requise de fonctions personnalisées](custom-functions-requirements.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
