---
ms.date: 03/21/2019
description: Créer des boîtes de dialogue via des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Les boîtes de dialogue fonctions personnalisées (aperçu)
localization_priority: Priority
ms.openlocfilehash: 0f596825a7a32525a68ef45656f1390196146706
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30926656"
---
# <a name="display-a-dialog-box-in-custom-functions"></a><span data-ttu-id="3899c-103">Affichent une boîte de dialogue dans les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="3899c-103">Display a dialog box in custom functions</span></span>

<span data-ttu-id="3899c-104">Si votre fonction personnalisée doit interagir avec l’utilisateur, vous pouvez créer une boîte de dialogue à l’aide de l’`OfficeRuntime.Dialog` objet.</span><span class="sxs-lookup"><span data-stu-id="3899c-104">If your custom function needs to interact with the user, you can create a dialog box using the `OfficeRuntime.Dialog` object.</span></span> <span data-ttu-id="3899c-105">Un scénario classique pour l’utilisation de la boîte de dialogue consiste à authentifier un utilisateur afin que votre fonction personnalisée puisse accéder à un service web.</span><span class="sxs-lookup"><span data-stu-id="3899c-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="3899c-106">Pour plus d’informations sur l’authentification de fonctions personnalisées, voir[authentification des fonctions personnalisées](./custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="3899c-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

<span data-ttu-id="3899c-107">Remarque : L’`OfficeRuntime.Dialog`objet fait partie de l’exécution de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="3899c-107">Note: The `OfficeRuntime.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="3899c-108">Il ne peut être utilisé à partir du contexte d’un volet de tâche.</span><span class="sxs-lookup"><span data-stu-id="3899c-108">It cannot be used from the context of a task pane.</span></span> <span data-ttu-id="3899c-109">Pour créer une boîte de dialogue à partir d’un volet de tâche, voir [Boîte de dialogue API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span><span class="sxs-lookup"><span data-stu-id="3899c-109">To create a dialog from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span></span>

## <a name="dialog-api-example"></a><span data-ttu-id="3899c-110">Exemple d’API Boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="3899c-110">Dialog API example</span></span>

<span data-ttu-id="3899c-111">Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la fonction `displayWebDialog` de l’API Boîte de dialogue pour afficher une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="3899c-111">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialog` function to display a dialog box.</span></span>

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {

      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {

        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else {
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }

  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
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
}
```

## <a name="see-also"></a><span data-ttu-id="3899c-112">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3899c-112">See also</span></span>

* [<span data-ttu-id="3899c-113">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="3899c-113">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="3899c-114">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="3899c-114">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="3899c-115">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="3899c-115">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="3899c-116">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="3899c-116">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="3899c-117">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="3899c-117">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
