---
ms.date: 05/17/2020
description: Authentifier les utilisateurs à l’aide de fonctions personnalisées dans Excel qui n’utilisent pas le volet Office.
title: Authentification pour les fonctions personnalisées sans interface utilisateur
localization_priority: Normal
ms.openlocfilehash: 93073fb23f3f4d30c36faf4927a3aebdafbc887d
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278377"
---
# <a name="authentication-for-ui-less-custom-functions"></a><span data-ttu-id="6d223-103">Authentification pour les fonctions personnalisées sans interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="6d223-103">Authentication for UI-less custom functions</span></span>

<span data-ttu-id="6d223-104">Dans certains scénarios, votre fonction personnalisée qui n’utilise pas de volet de tâches ou d’autres éléments de l’interface utilisateur (fonction personnalisée sans interface utilisateur) doit authentifier l’utilisateur afin d’accéder aux ressources protégées.</span><span class="sxs-lookup"><span data-stu-id="6d223-104">In some scenarios your custom function that does not use a task pane or other user interface elements (UI-less custom function) will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="6d223-105">N’oubliez pas que les fonctions personnalisées sans interface utilisateur s’exécutent dans un Runtime JavaScript uniquement.</span><span class="sxs-lookup"><span data-stu-id="6d223-105">Be aware that UI-less custom functions run in a JavaScript-only runtime.</span></span> <span data-ttu-id="6d223-106">Pour cette raison, vous devez transmettre les données entre le runtime JavaScript uniquement et le runtime du moteur de navigateur standard utilisé par la plupart des compléments à l’aide de l' `OfficeRuntime.storage` objet et de l’API de dialogue.</span><span class="sxs-lookup"><span data-stu-id="6d223-106">Because of this, you'll need to pass data back and forth between the JavaScript-only runtime and the typical browser engine runtime used by most add-ins using the `OfficeRuntime.storage` object and the Dialog API.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a><span data-ttu-id="6d223-107">Objet OfficeRuntime.storage</span><span class="sxs-lookup"><span data-stu-id="6d223-107">OfficeRuntime.storage object</span></span>

<span data-ttu-id="6d223-108">Le runtime JavaScript uniquement utilisé par des fonctions personnalisées sans interface utilisateur ne dispose pas d’un `localStorage` objet disponible dans la fenêtre globale, dans laquelle vous stockez généralement les données.</span><span class="sxs-lookup"><span data-stu-id="6d223-108">The JavaScript-only runtime used by UI-less custom functions doesn't have a `localStorage` object available on the global window, where you typically store data.</span></span> <span data-ttu-id="6d223-109">Au lieu de cela, vous devez partager des données entre des fonctions personnalisées sans interface utilisateur et des volets de tâches à l’aide de [OfficeRuntime. Storage](/javascript/api/office-runtime/officeruntime.storage) pour définir et obtenir des données.</span><span class="sxs-lookup"><span data-stu-id="6d223-109">Instead, you should share data between UI-less custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="6d223-110">Utilisation suggérée</span><span class="sxs-lookup"><span data-stu-id="6d223-110">Suggested usage</span></span>

<span data-ttu-id="6d223-111">Lorsque vous devez vous authentifier à partir d’une fonction personnalisée sans interface utilisateur, vérifiez `storage` si le jeton d’accès a déjà été acquis.</span><span class="sxs-lookup"><span data-stu-id="6d223-111">When you need to authenticate from a UI-less custom function, check `storage` to see if the access token was already acquired.</span></span> <span data-ttu-id="6d223-112">Si ce n’est pas le cas, utilisez l’API de boîte de dialogue pour authentifier l’utilisateur, récupérer le jeton d’accès, puis stocker le jeton dans `storage`pour une utilisation ultérieure.</span><span class="sxs-lookup"><span data-stu-id="6d223-112">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="6d223-113">API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="6d223-113">Dialog API</span></span>

<span data-ttu-id="6d223-114">Si un jeton n’existe pas, vous devez utiliser l’API de boîte de dialogue pour demander à l’utilisateur de se connecter.</span><span class="sxs-lookup"><span data-stu-id="6d223-114">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="6d223-115">Une fois qu’un utilisateur a entré ses informations d’identification, le jeton d’accès résultant peut être stocké dans `storage`.</span><span class="sxs-lookup"><span data-stu-id="6d223-115">After a user enters their credentials, the resulting access token can be stored in `storage`.</span></span>

> [!NOTE]
> <span data-ttu-id="6d223-116">Le runtime JavaScript uniquement utilise un objet Dialog qui est légèrement différent de l’objet Dialog dans le runtime du moteur du navigateur utilisé par les volets des tâches.</span><span class="sxs-lookup"><span data-stu-id="6d223-116">The JavaScript-only runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="6d223-117">Ils sont tous deux appelés « API de dialogue », mais utilisent `OfficeRuntime.Dialog` pour authentifier les utilisateurs dans le runtime JavaScript uniquement.</span><span class="sxs-lookup"><span data-stu-id="6d223-117">They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the JavaScript-only runtime.</span></span>

<span data-ttu-id="6d223-118">Le diagramme suivant décrit ce processus de base.</span><span class="sxs-lookup"><span data-stu-id="6d223-118">The following diagram outlines this basic process.</span></span> <span data-ttu-id="6d223-119">La ligne pointillée indique que les fonctions personnalisées sans interface utilisateur et le volet Office de votre complément font partie de votre complément dans son intégralité, même s’ils utilisent des runtimes distincts.</span><span class="sxs-lookup"><span data-stu-id="6d223-119">The dotted line indicates that UI-less custom functions and your add-in's task pane are both part of your add-in as a whole, though they use separate runtimes.</span></span>

1. <span data-ttu-id="6d223-120">Vous émettez un appel de fonction personnalisée sans interface utilisateur à partir d’une cellule dans un classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="6d223-120">You issue a UI-less custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="6d223-121">La fonction personnalisée sans interface utilisateur utilise `Dialog` pour transmettre vos informations d’identification d’utilisateur à un site Web.</span><span class="sxs-lookup"><span data-stu-id="6d223-121">The UI-less custom function uses `Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="6d223-122">Ce site Web renvoie ensuite un jeton d’accès à la fonction personnalisée sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6d223-122">This website then returns an access token to the UI-less custom function.</span></span>
4. <span data-ttu-id="6d223-123">Votre fonction personnalisée sans interface utilisateur définit ensuite le jeton d’accès sur `storage` .</span><span class="sxs-lookup"><span data-stu-id="6d223-123">Your UI-less custom function then sets this access token to the `storage`.</span></span>
5. <span data-ttu-id="6d223-124">Le volet de tâches de votre complément accède au jeton à partir de`storage`.</span><span class="sxs-lookup"><span data-stu-id="6d223-124">Your add-in's task pane accesses the token from `storage`.</span></span>

<span data-ttu-id="6d223-125">![Diagramme de la fonction personnalisée à l’aide de l’API de boîte de dialogue pour obtenir le jeton d’accès, puis partager le jeton avec le volet de tâches via l’API OfficeRuntime. Storage.](../images/authentication-diagram.png "Diagramme d’authentification.")</span><span class="sxs-lookup"><span data-stu-id="6d223-125">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="6d223-126">Stockage du jeton</span><span class="sxs-lookup"><span data-stu-id="6d223-126">Storing the token</span></span>

<span data-ttu-id="6d223-127">Les exemples suivants s’appliquent à partir de l’exemple de code[utilisation d’OfficeRuntime.storage dans les fonctions personnalisées](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage).</span><span class="sxs-lookup"><span data-stu-id="6d223-127">The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="6d223-128">Pour obtenir un exemple complet de partage de données entre des fonctions personnalisées sans interface utilisateur et le volet Office, reportez-vous à cet exemple de code.</span><span class="sxs-lookup"><span data-stu-id="6d223-128">Refer to this code sample for a complete example of sharing data between UI-less custom functions and the task pane.</span></span>

<span data-ttu-id="6d223-129">Si la fonction personnalisée sans interface utilisateur s’authentifie, elle reçoit le jeton d’accès et doit le stocker dans `storage` .</span><span class="sxs-lookup"><span data-stu-id="6d223-129">If the UI-less custom function authenticates, then it receives the access token and will need to store it in `storage`.</span></span> <span data-ttu-id="6d223-130">L’exemple de code suivant montre comment appeler la méthode`storage.setItem` pour stocker une valeur.</span><span class="sxs-lookup"><span data-stu-id="6d223-130">The following code sample shows how to call the `storage.setItem` method to store a value.</span></span> <span data-ttu-id="6d223-131">La `storeValue` fonction est une fonction personnalisée sans interface utilisateur qui, par exemple, stocke une valeur de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6d223-131">The `storeValue` function is a UI-less custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="6d223-132">Vous pouvez modifier cette valeur pour stocker les valeurs de jeton dont vous avez besoin.</span><span class="sxs-lookup"><span data-stu-id="6d223-132">You can modify this to store any token value you need.</span></span>

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

<span data-ttu-id="6d223-133">Lorsque le volet de tâches a besoin du jeton d’accès, il peut récupérer le jeton à partir de `storage`.</span><span class="sxs-lookup"><span data-stu-id="6d223-133">When the task pane needs the access token, it can retrieve the token from `storage`.</span></span> <span data-ttu-id="6d223-134">L’exemple de code suivant montre comment utiliser la méthode`storage.getItem` pour récupérer le jeton.</span><span class="sxs-lookup"><span data-stu-id="6d223-134">The following code sample shows how to use the `storage.getItem` method to retrieve the token.</span></span>

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a><span data-ttu-id="6d223-135">Instructions générales</span><span class="sxs-lookup"><span data-stu-id="6d223-135">General guidance</span></span>

<span data-ttu-id="6d223-136">Les compléments Office sont basés sur le Web et vous pouvez utiliser n’importe quelle technique d’authentification Web.</span><span class="sxs-lookup"><span data-stu-id="6d223-136">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="6d223-137">Il n’existe pas de modèle ni de méthode particulier à respecter pour implémenter votre propre authentification avec des fonctions personnalisées sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6d223-137">There is no particular pattern or method you must follow to implement your own authentication with UI-less custom functions.</span></span> <span data-ttu-id="6d223-138">Vous pouvez consulter la documentation relative à différents modèles d’authentification, en commençant par[cet article sur l’autorisation d’accès via les services externes](../develop/auth-external-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="6d223-138">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](../develop/auth-external-add-ins.md).</span></span>  

<span data-ttu-id="6d223-139">Évitez d’utiliser les emplacements suivants pour stocker des données lors du développement de fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="6d223-139">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="6d223-140">`localStorage`: Les fonctions personnalisées sans interface utilisateur n’ont pas accès à l' `window` objet global et, par conséquent, n’ont pas accès aux données stockées dans `localStorage` .</span><span class="sxs-lookup"><span data-stu-id="6d223-140">`localStorage`: UI-less custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.</span></span>
- <span data-ttu-id="6d223-141">`Office.context.document.settings`: Cet emplacement n’est pas sécurisé et les informations peuvent être extraites par toute personne utilisant le complément.</span><span class="sxs-lookup"><span data-stu-id="6d223-141">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="6d223-142">Exemple d’API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="6d223-142">Dialog box API example</span></span>

<span data-ttu-id="6d223-143">Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la `Dialog` fonction de l’API `displayWebDialogOptions` pour afficher une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="6d223-143">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API's `displayWebDialogOptions` function to display a dialog box.</span></span> <span data-ttu-id="6d223-144">Cet exemple est fourni pour afficher les fonctionnalités de l' `Dialog` objet, ne pas montrer comment s’authentifier.</span><span class="sxs-lookup"><span data-stu-id="6d223-144">This sample is provided to show the capabilities of the `Dialog` object, not demonstrate how to authenticate.</span></span>

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
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

## <a name="next-steps"></a><span data-ttu-id="6d223-145">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="6d223-145">Next steps</span></span>
<span data-ttu-id="6d223-146">Découvrez comment [Déboguer des fonctions personnalisées sans interface utilisateur](custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="6d223-146">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6d223-147">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6d223-147">See also</span></span>

* [<span data-ttu-id="6d223-148">Runtime pour les fonctions personnalisées Excel sans interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="6d223-148">Runtime for UI-less Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="6d223-149">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="6d223-149">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
