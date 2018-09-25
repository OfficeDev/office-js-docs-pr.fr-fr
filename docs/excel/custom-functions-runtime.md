---
ms.date: 09/20/2018
description: Les fonctions personnalisées Excel utilisent un nouveau runtime JavaScript, qui diffère du runtime de contrôle WebView de compléments standard.
title: Runtime pour les fonctions personnalisées Excel
ms.openlocfilehash: fa2b2030259e05f64b8b4660ded8b80c6af1eb5a
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985794"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="b7adc-103">Runtime pour les fonctions personnalisées Excel (Aperçu)</span><span class="sxs-lookup"><span data-stu-id="b7adc-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="b7adc-104">Les fonctions personnalisées étendent les fonctionnalités d’Excel à l’aide d’un nouveau runtime JavaScript qui utilise un moteur de JavaScript en bac à sable plutôt que dans un navigateur web.</span><span class="sxs-lookup"><span data-stu-id="b7adc-104">Custom functions extend Excel’s capabilities by using a new JavaScript runtime that uses a sandboxed JavaScript engine rather than a web browser.</span></span> <span data-ttu-id="b7adc-105">Puisque les fonctions personnalisées n’ont pas besoin d’afficher des éléments d’interface utilisateur, le nouveau runtime JavaScript est optimisé pour l’exécution de calculs, ce qui vous permet d’exécuter simultanément des milliers de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b7adc-105">Because custom functions do not need to render UI elements, the new JavaScript runtime is optimized for performing calculations, enabling you to run thousands of custom functions simultaneously.</span></span>

## <a name="key-facts-about-the-new-javascript-runtime"></a><span data-ttu-id="b7adc-106">Points clés concernant le nouveau runtime JavaScript</span><span class="sxs-lookup"><span data-stu-id="b7adc-106">Key facts about the new JavaScript runtime</span></span> 

<span data-ttu-id="b7adc-107">Seules les fonctions personnalisées au sein d’un complément utiliseront le nouveau runtime JavaScript décrit dans cet article.</span><span class="sxs-lookup"><span data-stu-id="b7adc-107">Only custom functions within an add-in will use the new JavaScript runtime that's described in this article.</span></span> <span data-ttu-id="b7adc-108">Si un complément comprend d’autres composants tels que des volets Office et d’autres éléments d’interface utilisateur, en plus des fonctions personnalisées, ces autres composants du complément continueront à s’exécuter dans le module d’exécution WebView.</span><span class="sxs-lookup"><span data-stu-id="b7adc-108">If an add-in includes other components such as task panes and other UI elements, in addition to custom functions, these other components of the add-in will continue to run in the browser-like WebView runtime.</span></span>  <span data-ttu-id="b7adc-109">En outre :</span><span class="sxs-lookup"><span data-stu-id="b7adc-109">Additionally:</span></span> 

- <span data-ttu-id="b7adc-110">Le runtime JavaScript ne fournit pas d’accès au Document Object Model (DOM) ni ne prend en charge des bibliothèques telles que jQuery qui s’appuient sur le modèle DOM.</span><span class="sxs-lookup"><span data-stu-id="b7adc-110">The JavaScript runtime does not provide access to the Document Object Model (DOM) or support libraries like jQuery that rely on the DOM.</span></span>

- <span data-ttu-id="b7adc-111">Une fonction personnalisée définie dans le fichier JavaScript d’un complément peut renvoyer une `Promise` JavaScript normale au lieu de renvoyer `OfficeExtension.Promise`.</span><span class="sxs-lookup"><span data-stu-id="b7adc-111">A custom function that's defined in an add-in's JavaScript file can return a regular JavaScript `Promise` instead of returning `OfficeExtension.Promise`.</span></span>  

- <span data-ttu-id="b7adc-112">Le fichier JSON qui spécifie les métadonnées de la fonction personnalisée n’a pas besoin de spécifier **sync** ou **async** dans **options**.</span><span class="sxs-lookup"><span data-stu-id="b7adc-112">The JSON file that specifies custom function metatdata does not need to specify **sync** or **async** within **options**.</span></span>

## <a name="new-apis"></a><span data-ttu-id="b7adc-113">Nouvelles API</span><span class="sxs-lookup"><span data-stu-id="b7adc-113">New Excel JavaScript APIs</span></span> 

<span data-ttu-id="b7adc-114">Le runtime JavaScript utilisé par les fonctions personnalisées comprend les APIs suivantes :</span><span class="sxs-lookup"><span data-stu-id="b7adc-114">The JavaScript runtime that's used by custom functions has the following APIs:</span></span>

- [<span data-ttu-id="b7adc-115">XHR</span><span class="sxs-lookup"><span data-stu-id="b7adc-115">XHR</span></span>](#xhr)
- [<span data-ttu-id="b7adc-116">WebSockets</span><span class="sxs-lookup"><span data-stu-id="b7adc-116">WebSockets</span></span>](#websockets)
- [<span data-ttu-id="b7adc-117">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="b7adc-117">AsyncStorage</span></span>](#asyncstorage)
- [<span data-ttu-id="b7adc-118">API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="b7adc-118">Dialog API requirement sets</span></span>](#dialog-api)

### <a name="xhr"></a><span data-ttu-id="b7adc-119">XHR</span><span class="sxs-lookup"><span data-stu-id="b7adc-119">XHR</span></span>

<span data-ttu-id="b7adc-120">XHR est l’acronyme de [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.</span><span class="sxs-lookup"><span data-stu-id="b7adc-120">XHR stands for [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="b7adc-121">Dans le nouveau runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant une [Stratégie d’Origine Identique](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et un [CORS](https://www.w3.org/TR/cors/) simple.</span><span class="sxs-lookup"><span data-stu-id="b7adc-121">In the new JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

<span data-ttu-id="b7adc-122">Dans l’exemple de code suivant, la fonction `getTemperature()` envoie une demande web pour obtenir la température d’une zone particulière basée sur l’ID du thermomètre.</span><span class="sxs-lookup"><span data-stu-id="b7adc-122">In the following code sample, the `getTemperature()` function sends a web request to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="b7adc-123">La fonction `sendWebRequest()` utilise XHR pour émettre une demande `GET` à un point de terminaison qui peut fournir les données.</span><span class="sxs-lookup"><span data-stu-id="b7adc-123">The `sendWebRequest()` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>  

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ //sendWebRequest is defined later in this code sample
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

//Helper method that uses Office's implementation of XMLHttpRequest in the new JavaScript runtime for custom functions  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
          };
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

```

### <a name="websockets"></a><span data-ttu-id="b7adc-124">WebSockets</span><span class="sxs-lookup"><span data-stu-id="b7adc-124">WebSockets</span></span>

<span data-ttu-id="b7adc-125">[WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) est un protocole réseau qui crée une communication en temps réel entre un serveur et un ou plusieurs clients.</span><span class="sxs-lookup"><span data-stu-id="b7adc-125">[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) is a networking protocol that creates real-time communication between a server and one or more clients.</span></span> <span data-ttu-id="b7adc-126">Il est souvent utilisé pour les applications de conversation, car il permet de lire et d’écrire du texte simultanément.</span><span class="sxs-lookup"><span data-stu-id="b7adc-126">It is often used for chat applications because it allows you to read and write text simultaneously.</span></span>  

<span data-ttu-id="b7adc-127">Comme indiqué dans l’exemple de code suivant, les fonctions personnalisées peuvent utiliser WebSocket.</span><span class="sxs-lookup"><span data-stu-id="b7adc-127">As shown in the following code sample, custom functions can use WebSockets.</span></span> <span data-ttu-id="b7adc-128">Dans cet exemple, le WebSocket enregistre chaque message qu’il reçoit.</span><span class="sxs-lookup"><span data-stu-id="b7adc-128">In this example, the WebSocket logs each message that it receives.</span></span>

```ts
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a><span data-ttu-id="b7adc-129">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="b7adc-129">AsyncStorage</span></span>

<span data-ttu-id="b7adc-130">AsyncStorage est un système de stockage clé-valeur qui peut être utilisé pour stocker les jetons d’authentification.</span><span class="sxs-lookup"><span data-stu-id="b7adc-130">AsyncStorage is a key-value storage system that can be used to store authentication tokens.</span></span> <span data-ttu-id="b7adc-131">Il est :</span><span class="sxs-lookup"><span data-stu-id="b7adc-131">It is Microsoft-supported.</span></span>

- <span data-ttu-id="b7adc-132">Persistant</span><span class="sxs-lookup"><span data-stu-id="b7adc-132">persistent</span></span>
- <span data-ttu-id="b7adc-133">Non chiffré</span><span class="sxs-lookup"><span data-stu-id="b7adc-133">Unencrypted</span></span>
- <span data-ttu-id="b7adc-134">Asynchrone</span><span class="sxs-lookup"><span data-stu-id="b7adc-134">Asynchronous calls</span></span>

<span data-ttu-id="b7adc-135">AsyncStorage est globalement disponible pour tous les composants de votre complément.</span><span class="sxs-lookup"><span data-stu-id="b7adc-135">AsyncStorage is globally available to all parts of your add-in.</span></span> <span data-ttu-id="b7adc-136">Pour les fonctions personnalisées, `AsyncStorage` est exposé comme un objet global.</span><span class="sxs-lookup"><span data-stu-id="b7adc-136">For custom functions, `AsyncStorage` is exposed as a global object.</span></span> <span data-ttu-id="b7adc-137">(Pour les autres composants de votre complément, tels que les volets Office et d’autres éléments qui utilisent le runtime WebView, AsyncStorage est exposé par le biais de `OfficeRuntime`.) Chaque complément a sa propre partition de stockage, avec une taille par défaut de 5 Mo.</span><span class="sxs-lookup"><span data-stu-id="b7adc-137">(For other parts of your add-in, such as task panes and other elements that use the WebView runtime, AsyncStorage is exposed through `OfficeRuntime`.) Each add-in has its own storage partition, with a default size of 5MB.</span></span> 

<span data-ttu-id="b7adc-138">Les méthodes suivantes sont disponibles sur l’objet `AsyncStorage` :</span><span class="sxs-lookup"><span data-stu-id="b7adc-138">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`
 
<span data-ttu-id="b7adc-139">À ce stade, les méthodes `mergeItem` et `multiMerge` ne sont pas prises en charge.</span><span class="sxs-lookup"><span data-stu-id="b7adc-139">At this time, the `mergeItem` and `multiMerge` methods are not supported.</span></span>

<span data-ttu-id="b7adc-140">L’exemple de code suivant appelle la fonction `AsyncStorage.getItem` pour récupérer une valeur stockée.</span><span class="sxs-lookup"><span data-stu-id="b7adc-140">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

```js
_goGetData = async () => {
    try {
        const value = await AsyncStorage.getItem('toDoItem');
        if (value !== null) {
            //data exists and you can do something with it here
            }
        } catch (error) {
            //handle errors here
        }
    }
}
```

### <a name="dialog-api"></a><span data-ttu-id="b7adc-141">API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="b7adc-141">Dialog API scenarios</span></span>

<span data-ttu-id="b7adc-142">L’API de boîte de dialogue vous permet d’ouvrir une boîte de dialogue qui invite l’utilisateur à se connecter.</span><span class="sxs-lookup"><span data-stu-id="b7adc-142">The Dialog API enables you to open a dialog box that prompts user sign-in.</span></span> <span data-ttu-id="b7adc-143">Vous pouvez utiliser l’API de boîte de dialogue pour demander une authentification utilisateur via une ressource externe, telle que Google ou Facebook, avant que l’utilisateur ne puisse utiliser la fonction.</span><span class="sxs-lookup"><span data-stu-id="b7adc-143">You can use the Dialog API to require user authentication through an outside resource, such as Google or Facebook, before the user can use your function.</span></span>   

<span data-ttu-id="b7adc-144">Dans l’exemple de code suivant, la méthode `getTokenViaDialog()` utilise la méthode `displayWebDialog()` de l’API boîte de dialogue pour ouvrir une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="b7adc-144">In the following code sample, the `getTokenViaDialog()` method uses the Dialog API’s `displayWebDialog()` method to open a dialog box.</span></span>

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"
 
function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://myauthurl")
    .then(function (token) {
      
      // Use token to get stock price
      fetch("https://myservice.com/?token=token&ticker= + ticker")
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
            dialog.closeDialog();
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

> [!NOTE]
> <span data-ttu-id="b7adc-145">L’API de boîte de dialogue décrite dans cette section fait partie du nouveau runtime JavaScript pour les fonctions personnalisées et peut être utilisée uniquement dans les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b7adc-145">The Dialog API described in this section is part of the new JavaScript runtime for custom functions and can be used only within custom functions.</span></span> <span data-ttu-id="b7adc-146">Cette API est différente de l’[API de boîte de dialogue](../develop/dialog-api-in-office-add-ins.md) qui peut être utilisé dans les volets Office et les commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="b7adc-146">This API is different from the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands.</span></span>

## <a name="see-also"></a><span data-ttu-id="b7adc-147">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b7adc-147">See also</span></span>

* [<span data-ttu-id="b7adc-148">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="b7adc-148">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="b7adc-149">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b7adc-149">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="b7adc-150">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b7adc-150">Custom functions best practices</span></span>](custom-functions-best-practices.md)
