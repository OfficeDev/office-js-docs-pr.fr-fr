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
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="f3f93-103">Gestion des erreurs dans des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f3f93-103">Error handling within custom functions</span></span>

<span data-ttu-id="f3f93-104">Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="f3f93-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="f3f93-105">La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="f3f93-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f3f93-106">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="f3f93-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="f3f93-107">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="f3f93-107">Next steps</span></span>
<span data-ttu-id="f3f93-108">Découvrez comment [résoudre les problèmes liés à vos fonctions personnalisées](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="f3f93-108">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f3f93-109">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f3f93-109">See also</span></span>

* [<span data-ttu-id="f3f93-110">Débogage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f3f93-110">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="f3f93-111">Configuration requise de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f3f93-111">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="f3f93-112">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="f3f93-112">Create custom functions in Excel</span></span>](custom-functions-overview.md)
