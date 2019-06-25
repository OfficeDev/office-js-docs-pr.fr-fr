---
ms.date: 06/21/2019
description: Demander, flux de données et annuler la diffusion en continu de données externes à votre classeur avec des fonctions personnalisées dans Excel
title: Recevoir et gérer des données à l’aide de fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: 39be2f0913e2eee4b1e5e7d5f704a47dee279cf5
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128254"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="2d9a0-103">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="2d9a0-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="2d9a0-104">L’une des façon dont les fonctions personnalisées améliorent la puissance d’Excel est qu’elles reçoivent des données en provenance d’emplacements autres que le classeur, par exemple, le web ou un serveur (via WebSockets).</span><span class="sxs-lookup"><span data-stu-id="2d9a0-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="2d9a0-105">Les fonctions personnalisées peuvent demander des données par le biais XHR et extraire (`fetch`) des demandes ainsi que des flux de ces données en temps réel.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-105">Custom functions can request data through XHR and `fetch` requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="2d9a0-106">La documentation ci-dessous illustre certains exemples de requêtes web, mais pour créer une fonction diffusion en continu pour vous-même, essayez la [didacticiel relative aux fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="2d9a0-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="2d9a0-107">Fonctions qui retournent des données provenant de sources externes</span><span class="sxs-lookup"><span data-stu-id="2d9a0-107">Functions that return data from external sources</span></span>

<span data-ttu-id="2d9a0-108">Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :</span><span class="sxs-lookup"><span data-stu-id="2d9a0-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="2d9a0-109">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="2d9a0-110">Résoudre la promesse avec la valeur finale à l’aide de la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="2d9a0-111">Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme[`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)Récupérer ou à l’aide de`XmlHttpRequest` [ (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="2d9a0-112">Dans le runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant la [politique de même origine (same-origin policy)](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et le partage [CORS (partage des ressources cross-origin)](https://www.w3.org/TR/cors/) simple.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="2d9a0-113">Notez qu’une implémentation CORS simples ne peut pas utiliser les cookies et prend uniquement en charge les méthodes simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="2d9a0-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="2d9a0-114">Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="2d9a0-115">Vous pouvez également utiliser un en-tête de Type de contenu dans CORS simple, autant que le type de contenu est `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="2d9a0-116">Exemple avec XHR</span><span class="sxs-lookup"><span data-stu-id="2d9a0-116">XHR example</span></span>

<span data-ttu-id="2d9a0-117">Dans l’exemple de code suivant, la fonction**obtenirTemperature**appelle la fonction pour obtenir la température d’une zone spécifique en fonction de l’ID de thermomètre.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="2d9a0-118">La fonction sendWebRequest utilise XHR pour émettre une demande GET à un point de terminaison qui peut fournir des données.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

```js
/**
 * Receives a temperature from an online source.
 * @customfunction
 * @param {number} thermometerID Identification number of the thermometer.
 */
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions.  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
        };

        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

CustomFunctions.associate("GETTEMPERATURE", getTemperature);
```

<span data-ttu-id="2d9a0-119">Pour un autre exemple d’une demande XHR avec davantage de contexte, voir la `getFile` fonction au sein de[ce fichier](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) dans le référentiel Github[Office-ajouter-dans-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="2d9a0-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="2d9a0-120">Exemple de récupération</span><span class="sxs-lookup"><span data-stu-id="2d9a0-120">Fetch example</span></span>

<span data-ttu-id="2d9a0-121">Dans l’exemple de code suivant, la fonction `stockPriceStream` utilise un symbole boursier pour obtenir le prix d’une action toutes les 1 000 millisecondes.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-121">In the following code sample, the `stockPriceStream` function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="2d9a0-122">Pour plus d’informations sur cet exemple voir le[didacticiel relatif aux fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="2d9a0-122">For more details about this sample, see the [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function).</span></span>

> [!NOTE]
> <span data-ttu-id="2d9a0-123">Le code suivant demande une cotation boursière à l’aide de l’API IEX Trading.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-123">The following code requests a stock quote using the IEX Trading API.</span></span> <span data-ttu-id="2d9a0-124">Avant d’exécuter le code, vous devez [créer un compte gratuit avec IEX Cloud](https://iexcloud.io/) de sorte que vous puissiez obtenir le jeton d’API requis dans la demande d’API.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-124">Before you can run the code, you'll need to [create a free account with IEX Cloud](https://iexcloud.io/) so that you can get the API token that's required in the API request.</span></span>

```js
/**
 * Streams a stock price.
 * @customfunction 
 * @param {string} ticker Stock ticker.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function stockPriceStream(ticker, invocation) {
    var updateFrequency = 1000 /* milliseconds*/;
    var isPending = false;

    var timer = setInterval(function() {
        // If there is already a pending request, skip this iteration:
        if (isPending) {
            return;
        }

        //Note: In the following line, replace <YOUR_TOKEN_HERE> with the API token that you've obtained through your IEX Cloud account.
        var url = "https://cloud.iexapis.com/stable/stock/" + ticker + "/quote/latestPrice?token=<YOUR_TOKEN_HERE>"
        isPending = true;

        fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                invocation.setResult(parseFloat(text));
            })
            .catch(function(error) {
                invocation.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    invocation.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receive-data-via-websockets"></a><span data-ttu-id="2d9a0-125">Recevoir des données via WebSockets</span><span class="sxs-lookup"><span data-stu-id="2d9a0-125">Receive data via WebSockets</span></span>

<span data-ttu-id="2d9a0-126">Dans une fonction personnalisée, vous pouvez utiliser WebSockets afin d’échanger des données avec un serveur via une connexion permanente.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-126">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="2d9a0-127">Grâce à WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir les données.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-127">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="2d9a0-128">Exemple avec WebSockets</span><span class="sxs-lookup"><span data-stu-id="2d9a0-128">WebSockets example</span></span>

<span data-ttu-id="2d9a0-129">L’exemple de code suivant établit une connexion WebSocket, puis consigne chaque message entrant provenant du serveur.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-129">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="2d9a0-130">Créer une fonction de diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="2d9a0-130">Make a streaming function</span></span>

<span data-ttu-id="2d9a0-131">Les fonctions personnalisées de diffusion vous aident à copier des données vers des cellules à plusieurs reprises, sans exiger qu’un utilisateur actualise explicitement quoi que ce soit.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-131">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="2d9a0-132">Cela peut s’avérer utile pour vérifier des données actives d’un service en ligne, comme la fonction dans le [didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="2d9a0-132">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="2d9a0-133">Pour déclarer une fonction de diffusion en continu, utilisez la balise de commentaire JSDoc `@stream`.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-133">To declare a streaming function, use the JSDoc comment tag `@stream`.</span></span> <span data-ttu-id="2d9a0-134">Pour attirer l’attention des utilisateurs sur le fait que votre fonction peut réévaluer sur la base de nouvelles informations, songez à placer un flux ou une formulation pour indiquer cela dans le nom ou la description de votre fonction.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-134">To alert users to the fact that your function may re-evaluate based on new information, consider putting stream or other wording to indicate this in the name or description of your function.</span></span>

<span data-ttu-id="2d9a0-135">L’exemple suivant illustre une fonction de diffusion en continu qui augmente un nombre donné à chaque seconde d’une valeur que vous spécifiez.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-135">The following example shows a streaming function which increases a given number every second by an amount you specify.</span></span>

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
CustomFunctions.associate("INC", increment);
```

>[!NOTE]
> <span data-ttu-id="2d9a0-136">Notez qu’il existe également une catégorie de fonctions appelée fonctions annulables, qui ne *sont* pas liées à des fonctions de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-136">Note that there are also a category of functions called cancelable functions, which are *not* related to streaming functions.</span></span> <span data-ttu-id="2d9a0-137">Les versions précédentes des fonctions personnalisées nécessitaient la déclaration manuelle de `"cancelable": true` et `"streaming": true` dans JSON.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-137">Previous versions of custom functions required you to declare `"cancelable": true` and `"streaming": true` in JSON written by hand.</span></span> <span data-ttu-id="2d9a0-138">Depuis l’introduction des métadonnées générées automatiquement, seules les fonctions personnalisées asynchrones renvoyant une valeur sont annulables.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-138">Since the introduction of autogenerated metadata, only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="2d9a0-139">Les fonctions annulables permettent de mettre fin à une requête web au milieu d’une demande, en utilisant une commande [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) pour décider de l’action à effectuer lors de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-139">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="2d9a0-140">Déclarez une fonction annulable à l’aide de la balise `@cancelable`.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-140">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="2d9a0-141">Utilisation d’un paramètre d’appel</span><span class="sxs-lookup"><span data-stu-id="2d9a0-141">Using an invocation parameter</span></span>

<span data-ttu-id="2d9a0-142">Par défaut, le paramètre `invocation` est le dernier de toute fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-142">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="2d9a0-143">Le paramètre `invocation` fournit du contexte sur la cellule (par exemple, son adresse), et vous permet d’utiliser les méthodes `setResult` et `onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-143">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="2d9a0-144">Ces méthodes définissent l’action d’une fonction quand elle diffuse (`setResult`) ou est annulée (`onCanceled`).</span><span class="sxs-lookup"><span data-stu-id="2d9a0-144">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="2d9a0-145">Si vous utilisez TypeScript, le gestionnaire d’appel doit être de type `CustomFunctions.StreamingInvocation` ou `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-145">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="2d9a0-146">Exemple de fonction de diffusion et annulable</span><span class="sxs-lookup"><span data-stu-id="2d9a0-146">Streaming and cancelable function example</span></span>
<span data-ttu-id="2d9a0-147">L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-147">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="2d9a0-148">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="2d9a0-148">Note the following about this code:</span></span>

- <span data-ttu-id="2d9a0-149">Excel affiche chaque nouvelle valeur automatiquement à l’aide de la méthode `setResult`.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-149">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="2d9a0-150">Le deuxième paramètre d’entrée, l’invocation, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-150">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="2d9a0-151">Le rappel `onCanceled` définit la fonction qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-151">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function(){
    clearInterval(timer);
    }
}
CustomFunctions.associate("INCREMENT", increment);
```

>[!NOTE]
> <span data-ttu-id="2d9a0-152">Excel annule l’exécution d’une fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="2d9a0-152">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="2d9a0-153">L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-153">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="2d9a0-154">Un des arguments (entrées) de la fonction est modifié.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-154">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="2d9a0-155">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-155">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="2d9a0-156">L’utilisateur déclenche manuellement le recalcul.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-156">When the user triggers recalculation manually.</span></span> <span data-ttu-id="2d9a0-157">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="2d9a0-157">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="2d9a0-158">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="2d9a0-158">Next steps</span></span>

* <span data-ttu-id="2d9a0-159">En savoir plus sur les [différents types de paramètres que vos fonctions peuvent utiliser](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="2d9a0-159">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="2d9a0-160">Découvrez comment [traiter par lots plusieurs appels d’API](custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="2d9a0-160">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="2d9a0-161">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2d9a0-161">See also</span></span>

* [<span data-ttu-id="2d9a0-162">Valeurs volatiles dans les fonctions</span><span class="sxs-lookup"><span data-stu-id="2d9a0-162">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="2d9a0-163">Créer des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="2d9a0-163">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="2d9a0-164">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="2d9a0-164">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="2d9a0-165">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="2d9a0-165">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="2d9a0-166">Meilleures pratiques pour l’utilisation des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="2d9a0-166">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="2d9a0-167">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="2d9a0-167">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="2d9a0-168">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="2d9a0-168">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
