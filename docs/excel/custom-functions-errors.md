---
ms.date: 05/03/2019
description: Gérez les erreurs dans vos fonctions personnalisées Excel.
title: Gestion des erreurs liées aux fonctions personnalisées dans Excel
localization_priority: Priority
ms.openlocfilehash: 188ece6c77bc2cafad6f22448fb698e0c0370ef8
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628157"
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
