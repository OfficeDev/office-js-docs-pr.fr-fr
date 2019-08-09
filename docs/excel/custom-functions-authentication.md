---
ms.date: 07/09/2019
description: Authentifiez les utilisateurs à l’aide de fonctions personnalisées dans Excel.
title: Authentification des fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: f746947122da7ef3d54a0dd3b4f90dd059e5830f
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268137"
---
# <a name="authentication-for-custom-functions"></a><span data-ttu-id="0347f-103">Authentification des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="0347f-103">Authentication for custom functions</span></span>

<span data-ttu-id="0347f-104">Dans certains scénarios, votre fonction personnalisée doit authentifier l’utilisateur pour accéder aux ressources protégées.</span><span class="sxs-lookup"><span data-stu-id="0347f-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="0347f-105">Bien que les fonctions personnalisées ne nécessitent pas de méthode spécifique d’authentification, sachez que les fonctions personnalisées s’exécutent dans un autre temps d’exécution, à partir du volet Office et d’autres éléments d’interface utilisateur de votre complément.</span><span class="sxs-lookup"><span data-stu-id="0347f-105">While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="0347f-106">Pour cette raison, vous devez transférer les données entre les deux exécutions à l’aide de l'objet`OfficeRuntime.storage` et de l’API de boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="0347f-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `OfficeRuntime.storage` object and the Dialog API.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="officeruntimestorage-object"></a><span data-ttu-id="0347f-107">Objet OfficeRuntime.storage</span><span class="sxs-lookup"><span data-stu-id="0347f-107">OfficeRuntime.storage object</span></span>

<span data-ttu-id="0347f-108">L’exécution des fonctions personnalisées n'a pas d’objet`localStorage`disponible dans la fenêtre globale, dans laquelle vous pouvez généralement stocker des données.</span><span class="sxs-lookup"><span data-stu-id="0347f-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="0347f-109">Au lieu de cela, vous devez partager les données entre les fonctions personnalisées et les volets de tâches à l’aide de [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) pour configurer et obtenir les données.</span><span class="sxs-lookup"><span data-stu-id="0347f-109">Instead, you should share data between custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.</span></span>

<span data-ttu-id="0347f-110">De plus, l’utilisation de l'objet `storage`est avantageuse ; il utilise un environnement sandbox sécurisé pour que vos données ne soient pas accessibles aux autres compléments.</span><span class="sxs-lookup"><span data-stu-id="0347f-110">Additionally, there is a benefit to using the `storage` object; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="0347f-111">Utilisation suggérée</span><span class="sxs-lookup"><span data-stu-id="0347f-111">Suggested usage</span></span>

<span data-ttu-id="0347f-112">Lorsque vous devez vous authentifier à partir du volet de tâches ou d’une fonction personnalisée, vérifiez `storage` pour voir si le jeton d’accès a déjà été acquis.</span><span class="sxs-lookup"><span data-stu-id="0347f-112">When you need to authenticate either from the task pane or a custom function, check `storage` to see if the access token was already acquired.</span></span> <span data-ttu-id="0347f-113">Si ce n’est pas le cas, utilisez l’API de boîte de dialogue pour authentifier l’utilisateur, récupérer le jeton d’accès, puis stocker le jeton dans `storage`pour une utilisation ultérieure.</span><span class="sxs-lookup"><span data-stu-id="0347f-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="0347f-114">API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="0347f-114">Dialog API scenarios</span></span>

<span data-ttu-id="0347f-115">Si un jeton n’existe pas, vous devez utiliser l’API de boîte de dialogue pour demander à l’utilisateur de se connecter.</span><span class="sxs-lookup"><span data-stu-id="0347f-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="0347f-116">Une fois qu’un utilisateur a entré ses informations d’identification, le jeton d’accès résultant peut être stocké dans `storage`.</span><span class="sxs-lookup"><span data-stu-id="0347f-116">After a user enters their credentials, the resulting access token can be stored in `storage`.</span></span>

> [!NOTE]
> <span data-ttu-id="0347f-117">Le runtime des fonctions personnalisées utilise un objet de boîte de dialogue qui est légèrement différent de l’objet de boîte de dialogue dans le moteur d’exécution du moteur d’exploration utilisé par les volets de tâches.</span><span class="sxs-lookup"><span data-stu-id="0347f-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="0347f-118">Ils sont tous deux appelés «API de boîte de dialogue», mais utilisent `OfficeRuntime.Dialog`pour authentifier les utilisateurs dans le runtime de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0347f-118">They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="0347f-119">Pour plus d’informations sur l’utilisation de l’objet `Dialog`, voir [boîte de dialogue fonctions personnalisées](/office/dev/add-ins/excel/custom-functions-dialog).</span><span class="sxs-lookup"><span data-stu-id="0347f-119">For information on how to use the `Dialog` object, see [Custom Functions dialog](/office/dev/add-ins/excel/custom-functions-dialog).</span></span>

<span data-ttu-id="0347f-120">Lorsque vous envisagez l’intégralité du processus d’authentification, il peut être utile de considérer les éléments du volet de tâches et de l’interface utilisateur de votre complément, ainsi que les fonctions personnalisées de votre complément en tant qu’entités distinctes pouvant communiquer entre eux via`OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="0347f-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `OfficeRuntime.storage`.</span></span>

<span data-ttu-id="0347f-121">Le diagramme suivant décrit ce processus de base.</span><span class="sxs-lookup"><span data-stu-id="0347f-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="0347f-122">Notez que la ligne pointillée indique qu’en effectuant des actions distinctes, les fonctions personnalisées et le volet de tâches de votre complément sont tous deux inclus dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="0347f-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.</span></span>

1. <span data-ttu-id="0347f-123">Vous émettez un appel de fonction personnalisée à partir d’une cellule d’un classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="0347f-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="0347f-124">La fonction personnalisée utilise `Dialog` pour transmettre vos informations d’identification d’utilisateur à un site Web.</span><span class="sxs-lookup"><span data-stu-id="0347f-124">The custom function uses `Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="0347f-125">Ce site Web renvoie ensuite un jeton d’accès à la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="0347f-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="0347f-126">Votre fonction personnalisée définit ensuite ce jeton d’accès sur `storage`.</span><span class="sxs-lookup"><span data-stu-id="0347f-126">Your custom function then sets this access token to the `storage`.</span></span>
5. <span data-ttu-id="0347f-127">Le volet de tâches de votre complément accède au jeton à partir de`storage`.</span><span class="sxs-lookup"><span data-stu-id="0347f-127">Your add-in's task pane accesses the token from `storage`.</span></span>

<span data-ttu-id="0347f-128">![Diagramme de la fonction personnalisée à l’aide de l’API de boîte de dialogue pour obtenir un jeton d’accès, puis partagez le jeton avec le volet de tâches via l’API OfficeRuntime.storage.](../images/authentication-diagram.png " Diagramme d’authentification.")</span><span class="sxs-lookup"><span data-stu-id="0347f-128">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="0347f-129">Stockage du jeton</span><span class="sxs-lookup"><span data-stu-id="0347f-129">Storing the token</span></span>

<span data-ttu-id="0347f-130">Les exemples suivants s’appliquent à partir de l’exemple de code[utilisation d’OfficeRuntime.storage dans les fonctions personnalisées](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage).</span><span class="sxs-lookup"><span data-stu-id="0347f-130">The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="0347f-131">Pour obtenir un exemple complet de partage de données entre les fonctions personnalisées et le volet de tâches, consultez cet exemple de code.</span><span class="sxs-lookup"><span data-stu-id="0347f-131">Refer to this code sample for a complete example of sharing data between custom functions and the task pane.</span></span>

<span data-ttu-id="0347f-132">Si la fonction personnalisée s’authentifie, elle reçoit le jeton d’accès et doit la stocker dans `storage`.</span><span class="sxs-lookup"><span data-stu-id="0347f-132">If the custom function authenticates, then it receives the access token and will need to store it in `storage`.</span></span> <span data-ttu-id="0347f-133">L’exemple de code suivant montre comment appeler la méthode`storage.setItem` pour stocker une valeur.</span><span class="sxs-lookup"><span data-stu-id="0347f-133">The following code sample shows how to call the `storage.setItem` method to store a value.</span></span> <span data-ttu-id="0347f-134">La fonction `storeValue`est une fonction personnalisée qui, à titre d’exemple, stocke une valeur de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0347f-134">The `storeValue` function is a custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="0347f-135">Vous pouvez modifier cette valeur pour stocker les valeurs de jeton dont vous avez besoin.</span><span class="sxs-lookup"><span data-stu-id="0347f-135">You can modify this to store any token value you need.</span></span>

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

<span data-ttu-id="0347f-136">Lorsque le volet de tâches a besoin du jeton d’accès, il peut récupérer le jeton à partir de `storage`.</span><span class="sxs-lookup"><span data-stu-id="0347f-136">When the task pane needs the access token, it can retrieve the token from `storage`.</span></span> <span data-ttu-id="0347f-137">L’exemple de code suivant montre comment utiliser la méthode`storage.getItem` pour récupérer le jeton.</span><span class="sxs-lookup"><span data-stu-id="0347f-137">The following code sample shows how to use the `storage.getItem` method to retrieve the token.</span></span>

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

## <a name="general-guidance"></a><span data-ttu-id="0347f-138">Instructions générales</span><span class="sxs-lookup"><span data-stu-id="0347f-138">General Guidance</span></span>

<span data-ttu-id="0347f-139">Les compléments Office sont basés sur le Web et vous pouvez utiliser n’importe quelle technique d’authentification Web.</span><span class="sxs-lookup"><span data-stu-id="0347f-139">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="0347f-140">Il n’existe pas de modèle ou de méthode spécifique que vous devez suivre pour implémenter votre propre authentification avec des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0347f-140">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="0347f-141">Vous pouvez consulter la documentation relative à différents modèles d’authentification, en commençant par[cet article sur l’autorisation d’accès via les services externes](/office/dev/add-ins/develop/auth-external-add-ins).</span><span class="sxs-lookup"><span data-stu-id="0347f-141">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](/office/dev/add-ins/develop/auth-external-add-ins).</span></span>  

<span data-ttu-id="0347f-142">Évitez d’utiliser les emplacements suivants pour stocker des données lors du développement de fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="0347f-142">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="0347f-143">`localStorage`: Les fonctions personnalisées n’ont pas accès à l’objet `window`global et n’ont par conséquent aucun accès aux données stockées dans`localStorage`.</span><span class="sxs-lookup"><span data-stu-id="0347f-143">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.</span></span>
- <span data-ttu-id="0347f-144">`Office.context.document.settings`: Cet emplacement n’est pas sécurisé et les informations peuvent être extraites par toute personne utilisant le complément.</span><span class="sxs-lookup"><span data-stu-id="0347f-144">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0347f-145">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="0347f-145">Next steps</span></span>
<span data-ttu-id="0347f-146">En savoir plus sur[l’API de boîte de dialogue pour les fonctions personnalisées](custom-functions-dialog.md).</span><span class="sxs-lookup"><span data-stu-id="0347f-146">Learn about the [dialog API for custom functions](custom-functions-dialog.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0347f-147">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0347f-147">See also</span></span>

* [<span data-ttu-id="0347f-148">Architecture des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="0347f-148">Custom functions architecture</span></span>](custom-functions-architecture.md)
* [<span data-ttu-id="0347f-149">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="0347f-149">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="0347f-150">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="0347f-150">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0347f-151">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="0347f-151">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
