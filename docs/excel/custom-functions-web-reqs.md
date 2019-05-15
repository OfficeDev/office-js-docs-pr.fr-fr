---
ms.date: 05/07/2019
description: Demander, flux de données et annuler la diffusion en continu de données externes à votre classeur avec des fonctions personnalisées dans Excel
title: Recevoir et gérer des données à l’aide de fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: 61f4d0fdaea4277faedddbe075a587fb23842c08
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659634"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="99da1-103">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="99da1-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="99da1-104">L’une des façon dont les fonctions personnalisées améliorent la puissance d’Excel est qu’elles reçoivent des données en provenance d’emplacements autres que le classeur, par exemple, le web ou un serveur (via WebSockets).</span><span class="sxs-lookup"><span data-stu-id="99da1-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="99da1-105">Les fonctions personnalisées peuvent demander des données par le biais XHR et extraire (`fetch`) des demandes ainsi que des flux de ces données en temps réel.</span><span class="sxs-lookup"><span data-stu-id="99da1-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="99da1-106">La documentation ci-dessous illustre certains exemples de requêtes web, mais pour créer une fonction diffusion en continu pour vous-même, essayez la [didacticiel relative aux fonctions personnalisées](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span><span class="sxs-lookup"><span data-stu-id="99da1-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="99da1-107">Fonctions qui retournent des données provenant de sources externes</span><span class="sxs-lookup"><span data-stu-id="99da1-107">Functions that return data from external sources</span></span>

<span data-ttu-id="99da1-108">Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :</span><span class="sxs-lookup"><span data-stu-id="99da1-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="99da1-109">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="99da1-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="99da1-110">Résoudre la promesse avec la valeur finale à l’aide de la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="99da1-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="99da1-111">Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme[`Fetch`](https://developer.mozilla.org/fr-FR/docs/Web/API/Fetch_API)Récupérer ou à l’aide de`XmlHttpRequest` [ (XHR)](https://developer.mozilla.org/fr-FR/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.</span><span class="sxs-lookup"><span data-stu-id="99da1-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/fr-FR/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/fr-FR/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="99da1-112">Dans le runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant la [politique de même origine (same-origin policy)](https://developer.mozilla.org/fr-FR/docs/Web/Security/Same-origin_policy) et le partage [CORS (partage des ressources cross-origin)](https://www.w3.org/TR/cors/) simple.</span><span class="sxs-lookup"><span data-stu-id="99da1-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/fr-FR/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="99da1-113">Notez qu’une implémentation CORS simples ne peut pas utiliser les cookies et prend uniquement en charge les méthodes simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="99da1-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="99da1-114">Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="99da1-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="99da1-115">Vous pouvez également utiliser un en-tête de Type de contenu dans CORS simple, autant que le type de contenu est `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="99da1-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="99da1-116">Exemple avec XHR</span><span class="sxs-lookup"><span data-stu-id="99da1-116">XHR example</span></span>

<span data-ttu-id="99da1-117">Dans l’exemple de code suivant, la fonction**obtenirTemperature**appelle la fonction pour obtenir la température d’une zone spécifique en fonction de l’ID de thermomètre.</span><span class="sxs-lookup"><span data-stu-id="99da1-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="99da1-118">La fonction sendWebRequest utilise XHR pour émettre une demande GET à un point de terminaison qui peut fournir des données.</span><span class="sxs-lookup"><span data-stu-id="99da1-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

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

<span data-ttu-id="99da1-119">Pour un autre exemple d’une demande XHR avec davantage de contexte, voir la `getFile` fonction au sein de[ce fichier](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) dans le référentiel Github[Office-ajouter-dans-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="99da1-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="99da1-120">Exemple de récupération</span><span class="sxs-lookup"><span data-stu-id="99da1-120">Fetch example</span></span>

<span data-ttu-id="99da1-121">Dans l’exemple de code suivant, la fonction `stockPriceStream` utilise un symbole boursier pour obtenir le prix d’une action toutes les 1 000 millisecondes.</span><span class="sxs-lookup"><span data-stu-id="99da1-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="99da1-122">Pour plus d’informations sur cet exemple voir le[didacticiel relatif aux fonctions personnalisées](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="99da1-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span>

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

        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
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

## <a name="receive-data-via-websockets"></a><span data-ttu-id="99da1-123">Recevoir des données via WebSockets</span><span class="sxs-lookup"><span data-stu-id="99da1-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="99da1-124">Dans une fonction personnalisée, vous pouvez utiliser WebSockets afin d’échanger des données avec un serveur via une connexion permanente.</span><span class="sxs-lookup"><span data-stu-id="99da1-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="99da1-125">Grâce à WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir les données.</span><span class="sxs-lookup"><span data-stu-id="99da1-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="99da1-126">Exemple avec WebSockets</span><span class="sxs-lookup"><span data-stu-id="99da1-126">WebSockets example</span></span>

<span data-ttu-id="99da1-127">L’exemple de code suivant établit une connexion WebSocket, puis consigne chaque message entrant provenant du serveur.</span><span class="sxs-lookup"><span data-stu-id="99da1-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="stream-and-cancel-functions"></a><span data-ttu-id="99da1-128">Fonctions de diffusion et annulables</span><span class="sxs-lookup"><span data-stu-id="99da1-128">Stream and cancel functions</span></span>

<span data-ttu-id="99da1-129">Les fonctions personnalisées de diffusion vous aident à copier des données vers des cellules à plusieurs reprises, sans exiger qu’un utilisateur actualise explicitement quoi que ce soit.</span><span class="sxs-lookup"><span data-stu-id="99da1-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span>

<span data-ttu-id="99da1-130">Les fonctions personnalisées annulables vous permettent d’annuler l’exécution d’une fonction personnalisée de diffusion pour réduire ses consommation de bande passante, de mémoire de travail et de temps processeur.</span><span class="sxs-lookup"><span data-stu-id="99da1-130">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span>

<span data-ttu-id="99da1-131">Pour déclarer une fonction comme étant de diffusion ou annulable, les indicateurs de commentaire JSDOC `@stream` ou `@cancelable`.</span><span class="sxs-lookup"><span data-stu-id="99da1-131">To declare a function as streaming or cancelable, use the JSDOC comment tags `@stream` or `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="99da1-132">Utilisation d’un paramètre d’appel</span><span class="sxs-lookup"><span data-stu-id="99da1-132">Using an invocation parameter</span></span>

<span data-ttu-id="99da1-133">Par défaut, le paramètre `invocation` est le dernier de toute fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="99da1-133">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="99da1-134">Le paramètre `invocation` fournit du contexte sur la cellule (par exemple, son adresse), et vous permet d’utiliser les méthodes `setResult` et `onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="99da1-134">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="99da1-135">Ces méthodes définissent l’action d’une fonction quand elle diffuse (`setResult`) ou est annulée (`onCanceled`).</span><span class="sxs-lookup"><span data-stu-id="99da1-135">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="99da1-136">Si vous utilisez TypeScript, le gestionnaire d’appel doit être de type `CustomFunctions.StreamingInvocation` ou `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="99da1-136">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="99da1-137">Exemple de fonction de diffusion et annulable</span><span class="sxs-lookup"><span data-stu-id="99da1-137">Streaming and cancelable function example</span></span>
<span data-ttu-id="99da1-138">L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde.</span><span class="sxs-lookup"><span data-stu-id="99da1-138">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="99da1-139">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="99da1-139">Note the following about this code:</span></span>

- <span data-ttu-id="99da1-140">Excel affiche chaque nouvelle valeur automatiquement à l’aide de la méthode `setResult`.</span><span class="sxs-lookup"><span data-stu-id="99da1-140">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="99da1-141">Le deuxième paramètre d’entrée, l’invocation, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.</span><span class="sxs-lookup"><span data-stu-id="99da1-141">The second input parameter, , is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="99da1-142">Le rappel `onCanceled` définit la fonction qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="99da1-142">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

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
> <span data-ttu-id="99da1-143">Excel annule l’exécution d’une fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="99da1-143">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="99da1-144">L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="99da1-144">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="99da1-145">Un des arguments (entrées) de la fonction est modifié.</span><span class="sxs-lookup"><span data-stu-id="99da1-145">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="99da1-146">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="99da1-146">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="99da1-147">L’utilisateur déclenche manuellement le recalcul.</span><span class="sxs-lookup"><span data-stu-id="99da1-147">When the user triggers recalculation manually.</span></span> <span data-ttu-id="99da1-148">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="99da1-148">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="99da1-149">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="99da1-149">Next steps</span></span>

* <span data-ttu-id="99da1-150">En savoir plus sur les [différents types de paramètres que vos fonctions peuvent utiliser](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="99da1-150">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="99da1-151">Découvrez comment [traiter par lots plusieurs appels d’API](custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="99da1-151">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="99da1-152">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="99da1-152">See also</span></span>

* [<span data-ttu-id="99da1-153">Valeurs volatiles dans les fonctions</span><span class="sxs-lookup"><span data-stu-id="99da1-153">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="99da1-154">Créer des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="99da1-154">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="99da1-155">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="99da1-155">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="99da1-156">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="99da1-156">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="99da1-157">Meilleures pratiques pour l’utilisation des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="99da1-157">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="99da1-158">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="99da1-158">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="99da1-159">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="99da1-159">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
