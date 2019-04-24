---
ms.date: 04/15/2019
description: Authentifier les utilisateurs à l'aide de fonctions personnalisées dans Excel.
title: Authentification pour les fonctions personnalisées
ms.openlocfilehash: 75ffb82c0dc9350c35b22b1d1676990598ea0c44
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449318"
---
# <a name="authentication"></a><span data-ttu-id="d927a-103">Authentification</span><span class="sxs-lookup"><span data-stu-id="d927a-103">Authentication</span></span>

<span data-ttu-id="d927a-104">Dans certains scénarios, votre fonction personnalisée doit authentifier l'utilisateur afin d'accéder aux ressources protégées.</span><span class="sxs-lookup"><span data-stu-id="d927a-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="d927a-105">Bien que les fonctions personnalisées ne nécessitent pas de méthode d'authentification spécifique, vous devez savoir que les fonctions personnalisées s'exécutent dans un Runtime distinct à partir du volet Office et d'autres éléments d'interface utilisateur de votre complément.</span><span class="sxs-lookup"><span data-stu-id="d927a-105">While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="d927a-106">Pour cette raison, vous devez transmettre les données entre les deux runtimes à l'aide de l' `AsyncStorage` objet et de l'API Dialog.</span><span class="sxs-lookup"><span data-stu-id="d927a-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `AsyncStorage` object and the Dialog API.</span></span>
  
## <a name="asyncstorage-object"></a><span data-ttu-id="d927a-107">Objet Dansasyncstorage</span><span class="sxs-lookup"><span data-stu-id="d927a-107">AsyncStorage object</span></span>

<span data-ttu-id="d927a-108">Le runtime des fonctions personnalisées ne `localStorage` dispose pas d'un objet disponible dans la fenêtre globale, dans laquelle vous pouvez généralement stocker des données.</span><span class="sxs-lookup"><span data-stu-id="d927a-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="d927a-109">Au lieu de cela, vous devez partager des données entre des fonctions personnalisées et des volets Office à l'aide de [OfficeRuntime. dansasyncstorage](/javascript/api/office-runtime/officeruntime.asyncstorage) pour définir et obtenir des données.</span><span class="sxs-lookup"><span data-stu-id="d927a-109">Instead, you should share data between custom functions and task panes by using [OfficeRuntime.AsyncStorage](/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.</span></span>

<span data-ttu-id="d927a-110">Par ailleurs, il est intéressant d'utiliser `AsyncStorage`; Il utilise un environnement de bac à sable (sandbox) sécurisé afin que les autres compléments ne puissent pas accéder à vos données.</span><span class="sxs-lookup"><span data-stu-id="d927a-110">Additionally, there is a benefit to using `AsyncStorage`; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="d927a-111">Utilisation suggérée</span><span class="sxs-lookup"><span data-stu-id="d927a-111">Suggested usage</span></span>

<span data-ttu-id="d927a-112">Lorsque vous devez vous authentifier à partir du volet Office ou d'une fonction personnalisée `AsyncStorage` , vérifiez si le jeton d'accès a déjà été acquis.</span><span class="sxs-lookup"><span data-stu-id="d927a-112">When you need to authenticate either from the task pane or a custom function, check `AsyncStorage` to see if the access token was already acquired.</span></span> <span data-ttu-id="d927a-113">Si ce n'est pas le cas, utilisez l'API de boîte de dialogue pour authentifier l'utilisateur, récupérer le `AsyncStorage` jeton d'accès, puis stocker le jeton en vue d'une utilisation ultérieure.</span><span class="sxs-lookup"><span data-stu-id="d927a-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `AsyncStorage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="d927a-114">API de dialogue</span><span class="sxs-lookup"><span data-stu-id="d927a-114">Dialog API</span></span>

<span data-ttu-id="d927a-115">Si un jeton n'existe pas, vous devez utiliser l'API de boîte de dialogue pour demander à l'utilisateur de se connecter.</span><span class="sxs-lookup"><span data-stu-id="d927a-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="d927a-116">Une fois qu'un utilisateur a entré ses informations d'identification, le jeton d'accès `AsyncStorage`résultant peut être stocké dans.</span><span class="sxs-lookup"><span data-stu-id="d927a-116">After a user enters their credentials, the resulting access token can be stored in `AsyncStorage`.</span></span>

> [!NOTE]
> <span data-ttu-id="d927a-117">Le runtime des fonctions personnalisées utilise un objet Dialog légèrement différent de l'objet Dialog dans le moteur d'exécution du moteur de navigateur utilisé par les volets des tâches.</span><span class="sxs-lookup"><span data-stu-id="d927a-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="d927a-118">Ils sont tous deux appelés «API de dialogue», mais utilisent `Officeruntime.Dialog` pour authentifier les utilisateurs dans le runtime des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="d927a-118">They're both referred to as the "Dialog API", but use `Officeruntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="d927a-119">Pour plus d'informations sur l'utilisation `OfficeRuntime.Dialog`de l', voir [Custom Functions Dialog](/office/dev/add-ins/excel/custom-functions-dialog).</span><span class="sxs-lookup"><span data-stu-id="d927a-119">For information on how to use the `OfficeRuntime.Dialog`, see [Custom Functions dialog](/office/dev/add-ins/excel/custom-functions-dialog).</span></span>

<span data-ttu-id="d927a-120">Lors de l'identification de l'ensemble du processus d'authentification, il peut s'avérer utile de considérer le volet des tâches et les éléments de l'interface utilisateur de votre complément, ainsi que les fonctions personnalisées de votre complément en tant qu'entités distinctes pouvant communiquer `AsyncStorage`les uns avec les autres.</span><span class="sxs-lookup"><span data-stu-id="d927a-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `AsyncStorage`.</span></span>

<span data-ttu-id="d927a-121">Le diagramme suivant décrit ce processus de base.</span><span class="sxs-lookup"><span data-stu-id="d927a-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="d927a-122">Notez que la ligne pointillée indique que lorsqu'ils effectuent des actions distinctes, les fonctions personnalisées et le volet Office de votre complément font partie de votre complément dans son intégralité.</span><span class="sxs-lookup"><span data-stu-id="d927a-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.</span></span>

1. <span data-ttu-id="d927a-123">Vous émettez un appel de fonction personnalisée à partir d'une cellule dans un classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="d927a-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="d927a-124">La fonction personnalisée utilise `Officeruntime.Dialog` pour transmettre les informations d'identification de votre utilisateur à un site Web.</span><span class="sxs-lookup"><span data-stu-id="d927a-124">The custom function uses `Officeruntime.Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="d927a-125">Ce site Web renvoie ensuite un jeton d'accès à la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="d927a-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="d927a-126">Votre fonction personnalisée définit ensuite le jeton d'accès sur `AsyncStorage`le.</span><span class="sxs-lookup"><span data-stu-id="d927a-126">Your custom function then sets this access token to the `AsyncStorage`.</span></span>
5. <span data-ttu-id="d927a-127">Le volet Office de votre complément accède au jeton à partir de `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="d927a-127">Your add-in's task pane accesses the token from `AsyncStorage`.</span></span>

<span data-ttu-id="d927a-128">![Diagramme de la fonction personnalisée à l'aide de l'API de boîte de dialogue pour obtenir le jeton d'accès, puis partager le jeton avec le volet de tâches via l'API dansasyncstorage.] (../images/authentication-diagram.png "Diagramme d'authentification.")</span><span class="sxs-lookup"><span data-stu-id="d927a-128">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the AsyncStorage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="d927a-129">Stockage du jeton</span><span class="sxs-lookup"><span data-stu-id="d927a-129">Storing the token</span></span>

<span data-ttu-id="d927a-130">Les exemples suivants sont tirés de l'exemple de code [utilisant dansasyncstorage dans des fonctions personnalisées](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) .</span><span class="sxs-lookup"><span data-stu-id="d927a-130">The following examples are from the [Using AsyncStorage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="d927a-131">Pour obtenir un exemple complet de partage de données entre des fonctions personnalisées et le volet Office, rePortez-vous à cet exemple de code.</span><span class="sxs-lookup"><span data-stu-id="d927a-131">Refer to this code sample for a complete example of sharing data between custom functions and the task pane.</span></span>

<span data-ttu-id="d927a-132">Si la fonction personnalisée s'authentifie, elle reçoit le jeton d'accès et le stocke dans `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="d927a-132">If the custom function authenticates, then it receives the access token and will need to store it in `AsyncStorage`.</span></span> <span data-ttu-id="d927a-133">L'exemple de code suivant montre comment appeler la `AsyncStorage.setItem` méthode pour stocker une valeur.</span><span class="sxs-lookup"><span data-stu-id="d927a-133">The following code sample shows how to call the `AsyncStorage.setItem` method to store a value.</span></span> <span data-ttu-id="d927a-134">La `StoreValue` fonction est une fonction personnalisée qui, à titre d'exemple, stocke une valeur de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d927a-134">The `StoreValue` function is a custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="d927a-135">Vous pouvez modifier cette valeur pour stocker les valeurs de jeton dont vous avez besoin.</span><span class="sxs-lookup"><span data-stu-id="d927a-135">You can modify this to store any token value you need.</span></span>

```javascript
function StoreValue(key, value) {
  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

<span data-ttu-id="d927a-136">Lorsque le volet Office a besoin du jeton d'accès, il peut récupérer le `AsyncStorage`jeton à partir de.</span><span class="sxs-lookup"><span data-stu-id="d927a-136">When the task pane needs the access token, it can retrieve the token from `AsyncStorage`.</span></span> <span data-ttu-id="d927a-137">L'exemple de code suivant montre comment utiliser la `AsyncStorage.getItem` méthode pour récupérer le jeton.</span><span class="sxs-lookup"><span data-stu-id="d927a-137">The following code sample shows how to use the `AsyncStorage.getItem` method to retrieve the token.</span></span>

```javascript
function ReceiveTokenFromCustomFunction() {
   var key = "token";
   var tokenSendStatus = document.getElementById('tokenSendStatus');
   OfficeRuntime.AsyncStorage.getItem(key).then(function (result) {
      tokenSendStatus.value = "Success: Item with key '" + key + "' read from AsyncStorage.";
      document.getElementById('tokenTextBox2').value = result;
   }, function (error) {
      tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from AsyncStorage. " + error;
   });
}
```

## <a name="general-guidance"></a><span data-ttu-id="d927a-138">Conseils généraux</span><span class="sxs-lookup"><span data-stu-id="d927a-138">General guidance</span></span>

<span data-ttu-id="d927a-139">Les compléments Office sont basés sur le Web et vous pouvez utiliser n'importe quelle technique d'authentification Web.</span><span class="sxs-lookup"><span data-stu-id="d927a-139">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="d927a-140">Il n'existe pas de modèle ni de méthode particulier à respecter pour implémenter votre propre authentification avec des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="d927a-140">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="d927a-141">Vous pouvez consulter la documentation sur les différents modèles d'authentification, en commençant par [cet article sur la création via des services externes](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="d927a-141">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span></span>  

<span data-ttu-id="d927a-142">Évitez d'utiliser les emplacements suivants pour stocker des données lors du développement de fonctions personnalisées:</span><span class="sxs-lookup"><span data-stu-id="d927a-142">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="d927a-143">`localStorage`: Les fonctions personnalisées n'ont pas accès à `window` l'objet global et, par conséquent, n'ont `localStorage`pas accès aux données stockées dans.</span><span class="sxs-lookup"><span data-stu-id="d927a-143">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data     stored in `localStorage`.</span></span>
- <span data-ttu-id="d927a-144">`Office.context.document.settings`: Cet emplacement n'est pas sécurisé et les informations peuvent être extraites par quiconque utilisant le complément.</span><span class="sxs-lookup"><span data-stu-id="d927a-144">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the     add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="d927a-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d927a-145">See also</span></span>

* [<span data-ttu-id="d927a-146">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d927a-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d927a-147">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="d927a-147">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="d927a-148">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d927a-148">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="d927a-149">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="d927a-149">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
