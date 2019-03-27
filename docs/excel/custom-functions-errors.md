---
ms.date: 02/08/2019
description: Gérez les erreurs dans vos fonctions personnalisées Excel.
title: Gestion des erreurs pour des fonctions personnalisées dans Excel (aperçu)
localization_priority: Priority
ms.openlocfilehash: 6c1c7f780aea125977510e4eb0e320933cd6ed9c
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870008"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="a19b2-103">Gestion des erreurs dans des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a19b2-103">Error handling within custom functions</span></span>

<span data-ttu-id="a19b2-104">Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="a19b2-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="a19b2-105">La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="a19b2-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="a19b2-106">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="a19b2-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
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

## <a name="see-also"></a><span data-ttu-id="a19b2-107">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a19b2-107">See also</span></span>

* [<span data-ttu-id="a19b2-108">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="a19b2-108">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="a19b2-109">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a19b2-109">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="a19b2-110">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="a19b2-110">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="a19b2-111">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a19b2-111">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="a19b2-112">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="a19b2-112">Custom functions changelog</span></span>](custom-functions-changelog.md)
