---
ms.date: 06/18/2019
description: Créez une boîte de dialogue via des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Afficher la boîte de dialogue d’une fonction personnalisée
localization_priority: Priority
ms.openlocfilehash: b8a2a6a68f23121fe4d3931c13fd45af6034c7d2
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940632"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a><span data-ttu-id="2afcd-103">Afficher la boîte de dialogue d’une fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="2afcd-103">Display a dialog box from a custom function</span></span>

<span data-ttu-id="2afcd-104">Si votre fonction personnalisée doit interagir avec l’utilisateur, vous pouvez créer une boîte de dialogue à l’aide de l’[objet`Office.Dialog`](/javascript/api/office-runtime/officeruntime.dialog?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="2afcd-104">If your custom function needs to interact with the user, you can create a dialog box using the [`Office.Dialog` object](/javascript/api/office-runtime/officeruntime.dialog?view=office-js).</span></span> <span data-ttu-id="2afcd-105">Un scénario classique pour l’utilisation de la boîte de dialogue consiste à authentifier un utilisateur afin que votre fonction personnalisée puisse accéder à un service web.</span><span class="sxs-lookup"><span data-stu-id="2afcd-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="2afcd-106">Pour plus d’informations sur l’authentification de fonctions personnalisées, voir[authentification des fonctions personnalisées](./custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="2afcd-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> <span data-ttu-id="2afcd-107">L’objet `Office.Dialog` fait partie de l’exécution de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="2afcd-107">The `Office.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="2afcd-108">Les volets Office n’utilisent pas l’objet `Dialog`.</span><span class="sxs-lookup"><span data-stu-id="2afcd-108">Task panes don't use the `Dialog` object.</span></span> <span data-ttu-id="2afcd-109">Pour créer une boîte de dialogue à partir d’un volet de tâches, consultez [API de boîte de dialogue](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span><span class="sxs-lookup"><span data-stu-id="2afcd-109">To create a dialog box from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="2afcd-110">Exemple d’API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="2afcd-110">dialog box API example</span></span>

<span data-ttu-id="2afcd-111">Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la fonction `Dialog` de l’API `displayWebDialogOptions` pour afficher une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="2afcd-111">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API’s `displayWebDialogOptions` function to display a dialog box.</span></span>

```js
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once, wait for previous dialog box's token
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      OfficeRuntime.displayWebDialog(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## <a name="next-steps"></a><span data-ttu-id="2afcd-112">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="2afcd-112">Next steps</span></span>
<span data-ttu-id="2afcd-113">Découvrez comment [rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="2afcd-113">Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="2afcd-114">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2afcd-114">See also</span></span>

* [<span data-ttu-id="2afcd-115">Authentification des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="2afcd-115">Custom functions authentication</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="2afcd-116">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="2afcd-116">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="2afcd-117">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="2afcd-117">Create custom functions in Excel</span></span>](custom-functions-overview.md)
