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
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="3c67d-103">Gestion des erreurs dans des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="3c67d-103">Error handling within custom functions</span></span>

<span data-ttu-id="3c67d-104">Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="3c67d-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="3c67d-105">La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="3c67d-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="3c67d-106">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="3c67d-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="3c67d-107">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="3c67d-107">Next steps</span></span>
<span data-ttu-id="3c67d-108">Découvrez comment [résoudre les problèmes liés à vos fonctions personnalisées](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="3c67d-108">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3c67d-109">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3c67d-109">See also</span></span>

* [<span data-ttu-id="3c67d-110">Débogage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="3c67d-110">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="3c67d-111">Configuration requise de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="3c67d-111">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="3c67d-112">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="3c67d-112">Create custom functions in Excel</span></span>](custom-functions-overview.md)
