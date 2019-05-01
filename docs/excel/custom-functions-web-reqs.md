---
ms.date: 04/20/2019
description: Demander, flux de données et annuler la diffusion en continu de données externes à votre classeur avec des fonctions personnalisées dans Excel
title: Requêtes Web et autres données gestion avec les fonctions personnalisées (aperçu)
localization_priority: Priority
ms.openlocfilehash: 2942ec56e46d6eb586b516eedab17c1eeb98d9c8
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353264"
---
# <a name="receiving-and-handling-data-with-custom-functions"></a><span data-ttu-id="4fbb9-103">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="4fbb9-103">Receiving and handling data with custom functions</span></span>

<span data-ttu-id="4fbb9-104">L’une des méthodes que les fonctions personnalisées améliorent la puissance d’Excel est en recevant des données à partir d’emplacements autre que le classeur, par exemple, le web ou un serveur (via WebSockets).</span><span class="sxs-lookup"><span data-stu-id="4fbb9-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="4fbb9-105">Les fonctions personnalisées peuvent demander des données par le biais XHR et récupérer des demandes ainsi que des flux de ces données en temps réel.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

<span data-ttu-id="4fbb9-106">La documentation ci-dessous illustre certains exemples de requêtes web, mais pour créer une fonction diffusion en continu pour vous-même, essayez la [didacticiel relative aux fonctions personnalisées](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span><span class="sxs-lookup"><span data-stu-id="4fbb9-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="4fbb9-107">Fonctions qui retournent des données provenant de sources externes</span><span class="sxs-lookup"><span data-stu-id="4fbb9-107">Functions that return data from external sources</span></span>

<span data-ttu-id="4fbb9-108">Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :</span><span class="sxs-lookup"><span data-stu-id="4fbb9-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="4fbb9-109">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="4fbb9-110">Résoudre la promesse avec la valeur finale à l’aide de la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="4fbb9-111">Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme[`Fetch`](https://developer.mozilla.org/fr-FR/docs/Web/API/Fetch_API)Récupérer ou à l’aide de`XmlHttpRequest` [ (XHR)](https://developer.mozilla.org/fr-FR/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/fr-FR/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/fr-FR/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="4fbb9-112">Dans le runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant la [politique de même origine (same-origin policy)](https://developer.mozilla.org/fr-FR/docs/Web/Security/Same-origin_policy) et le partage [CORS (partage des ressources cross-origin)](https://www.w3.org/TR/cors/) simple.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/fr-FR/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="4fbb9-113">Notez qu’une implémentation CORS simples ne peut pas utiliser les cookies et prend uniquement en charge les méthodes simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="4fbb9-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="4fbb9-114">Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="4fbb9-115">Vous pouvez également utiliser un en-tête de Type de contenu dans CORS simple, autant que le type de contenu est `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="4fbb9-116">Exemple avec XHR</span><span class="sxs-lookup"><span data-stu-id="4fbb9-116">XHR example</span></span>

<span data-ttu-id="4fbb9-117">Dans l’exemple de code suivant, la fonction**obtenirTemperature**appelle la fonction pour obtenir la température d’une zone spécifique en fonction de l’ID de thermomètre.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="4fbb9-118">La fonction sendWebRequest utilise XHR pour émettre une demande GET à un point de terminaison qui peut fournir des données.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

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

<span data-ttu-id="4fbb9-119">Pour un autre exemple d’une demande XHR avec davantage de contexte, voir la `getFile` fonction au sein de[ce fichier](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) dans le référentiel Github[Office-ajouter-dans-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="4fbb9-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="4fbb9-120">Exemple de récupération</span><span class="sxs-lookup"><span data-stu-id="4fbb9-120">Fetch example</span></span>

<span data-ttu-id="4fbb9-121">Dans l’exemple de code suivant, la fonction stockPriceStream utilise un symbole boursier pour obtenir le prix d’une action chaque 1000 millisecondes.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="4fbb9-122">Pour plus d’informations sur cet exemple et obtenir le JSON correspondant, voir le[didacticiel relatif aux fonctions personnalisées](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="4fbb9-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span> 

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

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="4fbb9-123">Réception de données via WebSockets</span><span class="sxs-lookup"><span data-stu-id="4fbb9-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="4fbb9-124">Dans une fonction personnalisée, vous pouvez utiliser WebSockets afin d’échanger des données avec un serveur via une connexion permanente.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="4fbb9-125">Grâce à WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir les données.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="4fbb9-126">Exemple avec WebSockets</span><span class="sxs-lookup"><span data-stu-id="4fbb9-126">WebSockets example</span></span>

<span data-ttu-id="4fbb9-127">L’exemple de code suivant établit une connexion WebSocket, puis consigne chaque message entrant provenant du serveur.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a><span data-ttu-id="4fbb9-128">Fonctions de diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="4fbb9-128">Streaming functions</span></span>

<span data-ttu-id="4fbb9-129">Les fonctions personnalisées de diffusion en continu vous aident à copier des données à des cellules à plusieurs reprises au fil du temps, sans exiger qu’un utilisateur demande explicitement l’actualisation des données.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="4fbb9-130">L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-130">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="4fbb9-131">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="4fbb9-131">Note the following about this code:</span></span>

- <span data-ttu-id="4fbb9-132">Excel affiche chaque nouvelle valeur automatiquement à l’aide du rappel setResult.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-132">Excel displays each new value automatically using the setResult callback.</span></span>
- <span data-ttu-id="4fbb9-133">Le deuxième paramètre d’entrée, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-133">The second input parameter, handler, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="4fbb9-134">Le rappel onCanceled définit la fonction qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-134">The onCanceled callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="4fbb9-135">Vous devez implémenter un gestionnaire d’annulation comme suit pour n’importe quelle fonction de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-135">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="4fbb9-136">Pour plus d’informations, voir [Annuler une fonction](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="4fbb9-136">For more information, see [Canceling a function](#canceling-a-function).</span></span>

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

<span data-ttu-id="4fbb9-137">Lorsque vous spécifiez des métadonnées pour une fonction de diffusion en continu dans le fichier de métadonnées JSON, vous pouvez générer automatiquement celui-ci à l’aide d’une balise de commentaire JSDOC `@streaming` dans le fichier script de votre fonction.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-137">When you specify metadata for a streaming function in the JSON metadata file, you can autogenerate this by using a `@streaming` JSDOC comment tag in your function's script file.</span></span> <span data-ttu-id="4fbb9-138">Pour plus d’informations, voir [Créer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="4fbb9-138">For more details, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="canceling-a-function"></a><span data-ttu-id="4fbb9-139">Annulation d’une fonction</span><span class="sxs-lookup"><span data-stu-id="4fbb9-139">Canceling a function</span></span>

<span data-ttu-id="4fbb9-140">Dans certains cas, vous devrez annuler l’exécution d’une fonction personnalisée de diffusion en continu pour réduire la consommation de bande passante, de la mémoire de travail et la charge du CPU.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-140">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="4fbb9-141">Excel annule l’exécution d’une fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="4fbb9-141">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="4fbb9-142">L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-142">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="4fbb9-143">Un des arguments (entrées) de la fonction est modifié.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-143">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="4fbb9-144">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-144">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="4fbb9-145">L’utilisateur déclenche manuellement le recalcul.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-145">When the user triggers recalculation manually.</span></span> <span data-ttu-id="4fbb9-146">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-146">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="4fbb9-147">Pour rendre une fonction annulable, implémentez un gestionnaire de code de fonction pour savoir comment procéder lorsque celui-ci est annulé.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-147">To make a function cancelable, implement a handler in your function's code to tell it what to do when it is canceled.</span></span> <span data-ttu-id="4fbb9-148">Utilisez également la balise de commentaire JSDOC `@cancelable` dans le fichier script de votre fonction.</span><span class="sxs-lookup"><span data-stu-id="4fbb9-148">Additionally, use the `@cancelable` JSDOC comment tag in your function's script file.</span></span> <span data-ttu-id="4fbb9-149">Pour plus d’informations, voir [Créer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="4fbb9-149">For more details, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="4fbb9-150">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4fbb9-150">See also</span></span>

* [<span data-ttu-id="4fbb9-151">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="4fbb9-151">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="4fbb9-152">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="4fbb9-152">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="4fbb9-153">Créer des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="4fbb9-153">Create JSON metadata for custom functions (preview)</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="4fbb9-154">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="4fbb9-154">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="4fbb9-155">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="4fbb9-155">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="4fbb9-156">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="4fbb9-156">Custom functions changelog</span></span>](custom-functions-changelog.md)
