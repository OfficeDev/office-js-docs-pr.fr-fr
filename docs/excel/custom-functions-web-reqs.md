---
ms.date: 03/21/2019
description: Demander, flux de données et annuler la diffusion en continu de données externes à votre classeur avec des fonctions personnalisées dans Excel
title: Requêtes Web et autres données gestion avec les fonctions personnalisées (aperçu)
localization_priority: Priority
ms.openlocfilehash: 9256e2aa87ec6d7b314314a1e4bc2b3793f1df5c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449707"
---
# <a name="receiving-and-handling-data-with-custom-functions"></a><span data-ttu-id="da751-103">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="da751-103">Receiving and handling data with custom functions</span></span>

<span data-ttu-id="da751-104">L’une des méthodes que les fonctions personnalisées améliorent la puissance d’Excel est en recevant des données à partir d’emplacements autre que le classeur, par exemple, le web ou un serveur (via WebSockets).</span><span class="sxs-lookup"><span data-stu-id="da751-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="da751-105">Les fonctions personnalisées peuvent demander des données par le biais XHR et récupérer des demandes ainsi que des flux de ces données en temps réel.</span><span class="sxs-lookup"><span data-stu-id="da751-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

<span data-ttu-id="da751-106">La documentation ci-dessous illustre certains exemples de requêtes web, mais pour créer une fonction diffusion en continu pour vous-même, essayez la [didacticiel relative aux fonctions personnalisées](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span><span class="sxs-lookup"><span data-stu-id="da751-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="da751-107">Fonctions qui retournent des données provenant de sources externes</span><span class="sxs-lookup"><span data-stu-id="da751-107">Functions that return data from external sources</span></span>

<span data-ttu-id="da751-108">Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :</span><span class="sxs-lookup"><span data-stu-id="da751-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="da751-109">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="da751-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="da751-110">Résoudre la promesse avec la valeur finale à l’aide de la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="da751-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="da751-111">Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme[`Fetch`](https://developer.mozilla.org/fr-FR/docs/Web/API/Fetch_API)Récupérer ou à l’aide de`XmlHttpRequest` [ (XHR)](https://developer.mozilla.org/fr-FR/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.</span><span class="sxs-lookup"><span data-stu-id="da751-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/fr-FR/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/fr-FR/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="da751-112">Dans le runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant la [politique de même origine (same-origin policy)](https://developer.mozilla.org/fr-FR/docs/Web/Security/Same-origin_policy) et le partage [CORS (partage des ressources cross-origin)](https://www.w3.org/TR/cors/) simple.</span><span class="sxs-lookup"><span data-stu-id="da751-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/fr-FR/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="da751-113">Notez qu’une implémentation CORS simples ne peut pas utiliser les cookies et prend uniquement en charge les méthodes simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="da751-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="da751-114">Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="da751-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="da751-115">Vous pouvez également utiliser un en-tête de Type de contenu dans CORS simple, autant que le type de contenu est `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="da751-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="da751-116">Exemple avec XHR</span><span class="sxs-lookup"><span data-stu-id="da751-116">XHR example</span></span>

<span data-ttu-id="da751-117">Dans l’exemple de code suivant, la fonction**obtenirTemperature**appelle la fonction pour obtenir la température d’une zone spécifique en fonction de l’ID de thermomètre.</span><span class="sxs-lookup"><span data-stu-id="da751-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="da751-118">La fonction sendWebRequest utilise XHR pour émettre une demande GET à un point de terminaison qui peut fournir des données.</span><span class="sxs-lookup"><span data-stu-id="da751-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

```JavaScript
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ 
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions  
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

<span data-ttu-id="da751-119">Pour un autre exemple d’une demande XHR avec davantage de contexte, voir la `getFile` fonction au sein de[ce fichier](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) dans le référentiel Github[Office-ajouter-dans-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="da751-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="da751-120">Exemple de récupération</span><span class="sxs-lookup"><span data-stu-id="da751-120">Fetch example</span></span>

<span data-ttu-id="da751-121">Dans l’exemple de code suivant, la fonction stockPriceStream utilise un symbole boursier pour obtenir le prix d’une action chaque 1000 millisecondes.</span><span class="sxs-lookup"><span data-stu-id="da751-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="da751-122">Pour plus d’informations sur cet exemple et obtenir le JSON correspondant, voir le[didacticiel relatif aux fonctions personnalisées](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="da751-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span> 

```JavaScript
function stockPriceStream(ticker, handler) {
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
                handler.setResult(parseFloat(text));
            })
            .catch(function(error) {
                handler.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    handler.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="da751-123">Réception de données via WebSockets</span><span class="sxs-lookup"><span data-stu-id="da751-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="da751-124">Dans une fonction personnalisée, vous pouvez utiliser WebSockets afin d’échanger des données avec un serveur via une connexion permanente.</span><span class="sxs-lookup"><span data-stu-id="da751-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="da751-125">Grâce à WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir les données.</span><span class="sxs-lookup"><span data-stu-id="da751-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="da751-126">Exemple avec WebSockets</span><span class="sxs-lookup"><span data-stu-id="da751-126">WebSockets example</span></span>

<span data-ttu-id="da751-127">L’exemple de code suivant établit une connexion WebSocket, puis consigne chaque message entrant provenant du serveur.</span><span class="sxs-lookup"><span data-stu-id="da751-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a><span data-ttu-id="da751-128">Fonctions de diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="da751-128">Streaming functions</span></span>

<span data-ttu-id="da751-129">Les fonctions personnalisées de diffusion en continu vous aident à copier des données à des cellules à plusieurs reprises au fil du temps, sans exiger qu’un utilisateur demande explicitement l’actualisation des données.</span><span class="sxs-lookup"><span data-stu-id="da751-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="da751-130">L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde.</span><span class="sxs-lookup"><span data-stu-id="da751-130">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="da751-131">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="da751-131">Note the following about this code:</span></span>

- <span data-ttu-id="da751-132">Excel affiche chaque nouvelle valeur automatiquement à l’aide du rappel setResult.</span><span class="sxs-lookup"><span data-stu-id="da751-132">Excel displays each new value automatically using the setResult callback.</span></span>
- <span data-ttu-id="da751-133">Le deuxième paramètre d’entrée, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.</span><span class="sxs-lookup"><span data-stu-id="da751-133">The second input parameter, handler, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="da751-134">Le rappel onCanceled définit la fonction qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="da751-134">The onCanceled callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="da751-135">Vous devez implémenter un gestionnaire d’annulation comme suit pour n’importe quelle fonction de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="da751-135">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="da751-136">Pour plus d’informations, voir [Annuler une fonction](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="da751-136">For more information, see [Canceling a function](#canceling-a-function).</span></span>

```JavaScript
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}

CustomFunctions.associate("INCREMENTVALUE", incrementValue);
```

<span data-ttu-id="da751-137">Lorsque vous spécifiez des métadonnées pour une fonction de diffusion en continu dans le fichier de métadonnées JSON, vous devez définir les propriétés «annulable»: vrai et «flux»: vrai au sein de l’objet, comme illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="da751-137">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties "cancelable": true and "stream": true within the options object, as shown in the following example.</span></span>

```JSON
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="da751-138">Annulation d’une fonction</span><span class="sxs-lookup"><span data-stu-id="da751-138">Canceling a function</span></span>

<span data-ttu-id="da751-139">Dans certains cas, vous devrez annuler l’exécution d’une fonction personnalisée de diffusion en continu pour réduire la consommation de bande passante, de la mémoire de travail et la charge du CPU.</span><span class="sxs-lookup"><span data-stu-id="da751-139">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="da751-140">Excel annule l’exécution d’une fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="da751-140">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="da751-141">L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="da751-141">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="da751-142">Un des arguments (entrées) de la fonction est modifié.</span><span class="sxs-lookup"><span data-stu-id="da751-142">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="da751-143">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="da751-143">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="da751-144">L’utilisateur déclenche manuellement le recalcul.</span><span class="sxs-lookup"><span data-stu-id="da751-144">When the user triggers recalculation manually.</span></span> <span data-ttu-id="da751-145">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="da751-145">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="da751-146">Pour rendre une fonction annulable, implémentez un gestionnaire de code de fonction pour savoir comment procéder lorsque celui-ci est annulé.</span><span class="sxs-lookup"><span data-stu-id="da751-146">To make a function cancelable, implement a handler in your function's code to tell it what to do when it is canceled.</span></span> <span data-ttu-id="da751-147">En outre, spécifiez spécifier la propriété `"cancelable": true` au sein de l’objet options dans les métadonnées JSON décrivant la fonction.</span><span class="sxs-lookup"><span data-stu-id="da751-147">Additionally, specify specify the property `"cancelable": true` within the options object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="da751-148">Les exemples de code dans la section précédente de cet article fournissent un exemple de ces techniques.</span><span class="sxs-lookup"><span data-stu-id="da751-148">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="see-also"></a><span data-ttu-id="da751-149">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="da751-149">See also</span></span>

* [<span data-ttu-id="da751-150">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="da751-150">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="da751-151">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="da751-151">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="da751-152">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="da751-152">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="da751-153">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="da751-153">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="da751-154">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="da751-154">Custom functions changelog</span></span>](custom-functions-changelog.md)
