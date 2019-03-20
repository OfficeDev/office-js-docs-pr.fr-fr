---
ms.date: 02/08/2019
description: Gérez les erreurs dans vos fonctions personnalisées Excel.
title: Gestion des erreurs pour des fonctions personnalisées dans Excel (aperçu)
localization_priority: Priority
ms.openlocfilehash: 170da03331663d6779bed7bf0bf5a9b75b908b3f
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30632694"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="5f233-103">Gestion des erreurs dans des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5f233-103">Error handling within custom functions</span></span>

<span data-ttu-id="5f233-104">Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="5f233-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="5f233-105">La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="5f233-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="5f233-106">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="5f233-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="see-also"></a><span data-ttu-id="5f233-107">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5f233-107">See also</span></span>

* [<span data-ttu-id="5f233-108">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="5f233-108">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="5f233-109">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5f233-109">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="5f233-110">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="5f233-110">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="5f233-111">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5f233-111">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="5f233-112">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="5f233-112">Custom functions changelog</span></span>](custom-functions-changelog.md)
