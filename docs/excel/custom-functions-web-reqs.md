---
ms.date: 04/29/2020
description: Demander, flux de données et annuler la diffusion en continu de données externes à votre classeur avec des fonctions personnalisées dans Excel
title: Recevoir et gérer des données à l’aide de fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: c53ad94c798f787447ab353201a245cd4f20d463
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610460"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="68413-103">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="68413-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="68413-104">L’une des façon dont les fonctions personnalisées améliorent la puissance d’Excel est qu’elles reçoivent des données en provenance d’emplacements autres que le classeur, par exemple, le web ou un serveur (via WebSockets).</span><span class="sxs-lookup"><span data-stu-id="68413-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="68413-105">Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme[`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)Récupérer ou à l’aide de`XmlHttpRequest` [ (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.</span><span class="sxs-lookup"><span data-stu-id="68413-105">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![Image GIF d'une fonction personnalisée diffusant le temps en continu à partir d'une API](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="68413-107">Fonctions qui retournent des données provenant de sources externes</span><span class="sxs-lookup"><span data-stu-id="68413-107">Functions that return data from external sources</span></span>

<span data-ttu-id="68413-108">Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :</span><span class="sxs-lookup"><span data-stu-id="68413-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="68413-109">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="68413-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="68413-110">Résoudre la promesse avec la valeur finale à l’aide de la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="68413-110">Resolve the Promise with the final value using the callback function.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="68413-111">Exemple de récupération</span><span class="sxs-lookup"><span data-stu-id="68413-111">Fetch example</span></span>

<span data-ttu-id="68413-112">Dans l’exemple de code suivant, la `webRequest` fonction accède à l’API « nombre de personnes dans l’espace contoso », qui effectue le suivi du nombre de personnes actuellement présentes sur la station internationale.</span><span class="sxs-lookup"><span data-stu-id="68413-112">In the following code sample, the `webRequest` function reaches out to the hypothetical Contoso "Number of People in Space" API, which tracks the number of people currently on the International Space Station.</span></span> <span data-ttu-id="68413-113">La fonction renvoie une promesse JavaScript et utilise la récupération pour demander des informations à l’API.</span><span class="sxs-lookup"><span data-stu-id="68413-113">The function returns a JavaScript Promise and uses fetch to request information from the API.</span></span> <span data-ttu-id="68413-114">Les données obtenues sont transformées en JSON et la `names` propriété est convertie en chaîne, ce qui permet de résoudre la promesse.</span><span class="sxs-lookup"><span data-stu-id="68413-114">The resulting data is transformed into JSON and the `names` property is converted into a string, which is used to resolve the Promise.</span></span>

<span data-ttu-id="68413-115">Lorsque vous développez vos propres fonctions, vous souhaitez peut-être effectuer une action si la requête Web ne se termine pas en temps voulu ou envisager de [regrouper plusieurs demandes API](./custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="68413-115">When developing your own functions, you may want to perform an action if the web request does not complete in a timely manner or consider [batching up multiple API requests](./custom-functions-batching.md).</span></span>

```JS
/**
 * Requests the names of the people currently on the International Space Station from a hypothetical API.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace";
  return new Promise(function (resolve, reject) {
    fetch(url)
      .then(function (response){
        return response.json();
        }
      )
      .then(function (json) {
        resolve(JSON.stringify(json.names));
      })
  })
}
```

>[!NOTE]
><span data-ttu-id="68413-116">L’utilisation de `Fetch` permet d’éviter les rappels imbriqués et peut être préférable à XHR dans certains cas.</span><span class="sxs-lookup"><span data-stu-id="68413-116">Using `Fetch` avoids nested callbacks and may be preferable to XHR in some cases.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="68413-117">Exemple avec XHR</span><span class="sxs-lookup"><span data-stu-id="68413-117">XHR example</span></span>

<span data-ttu-id="68413-118">Dans l’exemple de code suivant, la `getStarCount` fonction appelle l’API GitHub pour découvrir la quantité d’étoiles donnée au référentiel d’un utilisateur particulier.</span><span class="sxs-lookup"><span data-stu-id="68413-118">In the following code sample, the `getStarCount` function calls the Github API to discover the amount of stars given to a particular user's repository.</span></span> <span data-ttu-id="68413-119">Il s’agit d’une fonction asynchrone qui renvoie une promesse JavaScript.</span><span class="sxs-lookup"><span data-stu-id="68413-119">This is an asynchronous function which returns a JavaScript Promise.</span></span> <span data-ttu-id="68413-120">Lorsque des données sont obtenues à partir de l’appel Web, la promesse est résolue et renvoie les données à la cellule.</span><span class="sxs-lookup"><span data-stu-id="68413-120">When data is obtained from the web call, the Promise is resolved which returns the data to the cell.</span></span>

```TS
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param userName string name of organization or user.
 * @param repoName string name of the repository.
 * @return number of stars.
 */

async function getStarCount(userName: string, repoName: string) {

  const url = "https://api.github.com/repos/" + userName + "/" + repoName;

  let xhttp = new XMLHttpRequest();

  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;

      if (xhttp.status == 200) {
        resolve(JSON.parse(xhttp.responseText).watchers_count);
      } else {
        reject({
          status: xhttp.status,

          statusText: xhttp.statusText
        });
      }
    };

    xhttp.open("GET", url, true);

    xhttp.send();
  });
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="68413-121">Créer une fonction de diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="68413-121">Make a streaming function</span></span>

<span data-ttu-id="68413-122">Les fonctions personnalisées de diffusion vous aident à copier des données vers des cellules à plusieurs reprises, sans exiger qu’un utilisateur actualise explicitement quoi que ce soit.</span><span class="sxs-lookup"><span data-stu-id="68413-122">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="68413-123">Cela peut s’avérer utile pour vérifier les données actives d’un service en ligne, comme la fonction dans le [didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="68413-123">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="68413-124">Pour déclarer une fonction de diffusion en continu, vous pouvez utiliser l’une des méthodes suivantes :</span><span class="sxs-lookup"><span data-stu-id="68413-124">To declare a streaming function, you can use either:</span></span>

- <span data-ttu-id="68413-125">La `@streaming` balise.</span><span class="sxs-lookup"><span data-stu-id="68413-125">The `@streaming` tag.</span></span>
- <span data-ttu-id="68413-126">Le `CustomFunctions.StreamingInvocation` paramètre invocation.</span><span class="sxs-lookup"><span data-stu-id="68413-126">The `CustomFunctions.StreamingInvocation` invocation parameter.</span></span>

<span data-ttu-id="68413-127">L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde.</span><span class="sxs-lookup"><span data-stu-id="68413-127">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="68413-128">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="68413-128">Note the following about this code:</span></span>

- <span data-ttu-id="68413-129">Excel affiche chaque nouvelle valeur automatiquement à l’aide de la méthode `setResult`.</span><span class="sxs-lookup"><span data-stu-id="68413-129">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="68413-130">Le deuxième paramètre d’entrée, l’invocation, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.</span><span class="sxs-lookup"><span data-stu-id="68413-130">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="68413-131">Le rappel `onCanceled` définit la fonction qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="68413-131">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>
- <span data-ttu-id="68413-132">La diffusion en continu n’est pas nécessairement liée à la création d’une requête Web : dans ce cas, la fonction ne crée pas de requête Web, mais continue d’obtenir des données à intervalles définis, de sorte qu’elle nécessite l’utilisation du paramètre `invocation` de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="68413-132">Streaming isn't necessarily tied to making a web request: in this case, the function isn't making a web request but is still getting data at set intervals, so it requires the use of the streaming `invocation` parameter.</span></span>

```JS
/**
 * Increments a value once a second.
 * @customfunction INC increment
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="68413-133">Annulation d’une fonction</span><span class="sxs-lookup"><span data-stu-id="68413-133">Canceling a function</span></span>

<span data-ttu-id="68413-134">Excel annule l’exécution d’une fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="68413-134">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="68413-135">L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="68413-135">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="68413-136">Un des arguments (entrées) de la fonction est modifié.</span><span class="sxs-lookup"><span data-stu-id="68413-136">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="68413-137">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="68413-137">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="68413-138">L’utilisateur déclenche manuellement le recalcul.</span><span class="sxs-lookup"><span data-stu-id="68413-138">When the user triggers recalculation manually.</span></span> <span data-ttu-id="68413-139">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="68413-139">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="68413-140">Vous pouvez également définir une valeur de diffusion en continu par défaut pour gérer les cas lorsqu’une demande est effectuée, mais que vous êtes en mode hors connexion.</span><span class="sxs-lookup"><span data-stu-id="68413-140">You can also consider setting a default streaming value to handle cases when a request is made but you are offline.</span></span>

<span data-ttu-id="68413-141">Notez qu’il existe également une catégorie de fonctions appelée fonctions annulables, qui ne sont _pas_ liées à des fonctions de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="68413-141">Note that there are also a category of functions called cancelable functions, which are _not_ related to streaming functions.</span></span> <span data-ttu-id="68413-142">Seules les fonctions personnalisées asynchrones qui retournent une valeur peuvent être annulées.</span><span class="sxs-lookup"><span data-stu-id="68413-142">Only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="68413-143">Les fonctions annulables permettent de mettre fin à une requête web au milieu d’une demande, en utilisant une commande [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) pour décider de l’action à effectuer lors de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="68413-143">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="68413-144">Déclarez une fonction annulable à l’aide de la balise `@cancelable`.</span><span class="sxs-lookup"><span data-stu-id="68413-144">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="68413-145">Utilisation d’un paramètre d’appel</span><span class="sxs-lookup"><span data-stu-id="68413-145">Using an invocation parameter</span></span>

<span data-ttu-id="68413-146">Par défaut, le paramètre `invocation` est le dernier de toute fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="68413-146">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="68413-147">Le `invocation` paramètre donne le contexte de la cellule (par exemple, son adresse et son contenu) et vous permet d’utiliser `setResult` et des `onCanceled` méthodes.</span><span class="sxs-lookup"><span data-stu-id="68413-147">The `invocation` parameter gives context about the cell (such as its address and contents) and allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="68413-148">Ces méthodes définissent l’action d’une fonction quand elle diffuse (`setResult`) ou est annulée (`onCanceled`).</span><span class="sxs-lookup"><span data-stu-id="68413-148">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="68413-149">Si vous utilisez la machine à écrire, le gestionnaire d’appel doit être de type [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) ou [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) .</span><span class="sxs-lookup"><span data-stu-id="68413-149">If you're using TypeScript, the invocation handler needs to be of type [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) or[`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation).</span></span>

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="68413-150">Réception de données via WebSockets</span><span class="sxs-lookup"><span data-stu-id="68413-150">Receiving data via WebSockets</span></span>

<span data-ttu-id="68413-151">Dans une fonction personnalisée, vous pouvez utiliser WebSockets afin d’échanger des données avec un serveur via une connexion permanente.</span><span class="sxs-lookup"><span data-stu-id="68413-151">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="68413-152">À l’aide de WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour les données.</span><span class="sxs-lookup"><span data-stu-id="68413-152">Using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="68413-153">Exemple avec WebSockets</span><span class="sxs-lookup"><span data-stu-id="68413-153">WebSockets example</span></span>

<span data-ttu-id="68413-154">L’exemple de code suivant établit une connexion WebSocket, puis consigne chaque message entrant provenant du serveur.</span><span class="sxs-lookup"><span data-stu-id="68413-154">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="next-steps"></a><span data-ttu-id="68413-155">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="68413-155">Next steps</span></span>

- <span data-ttu-id="68413-156">En savoir plus sur les [différents types de paramètres que vos fonctions peuvent utiliser](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="68413-156">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
- <span data-ttu-id="68413-157">Découvrez comment [traiter par lots plusieurs appels d’API](custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="68413-157">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="68413-158">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="68413-158">See also</span></span>

- [<span data-ttu-id="68413-159">Valeurs volatiles dans les fonctions</span><span class="sxs-lookup"><span data-stu-id="68413-159">Volatile values in functions</span></span>](custom-functions-volatile.md)
- [<span data-ttu-id="68413-160">Créer des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="68413-160">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="68413-161">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="68413-161">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="68413-162">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="68413-162">Create custom functions in Excel</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="68413-163">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="68413-163">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
