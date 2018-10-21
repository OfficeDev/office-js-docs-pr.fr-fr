---
ms.date: 10/17/2018
description: Comprendre les scénarios clés du développement de fonctions Excel personnalisées utilisant le nouveau runtime JavaScript.
title: Exécution de fonctions personnalisées Excel
ms.openlocfilehash: 333816c3916af1490d14b8344c4bb49094f9a7f9
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640014"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="ba791-103">Runtime pour les fonctions personnalisées Excel (préversion)</span><span class="sxs-lookup"><span data-stu-id="ba791-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="ba791-p101">Les fonctions personnalisées utilisent un nouveau runtime JavaScript qui diffère du runtime utilisé par les autres composants d’un complément, tels que le volet Office ou d’autres éléments d’interface utilisateur. Ce runtime JavaScript est conçu pour optimiser les performances des calculs dans les fonctions personnalisées et expose les nouvelles API que vous pouvez utiliser pour effectuer des actions courantes sur le Web dans des fonctions personnalisées comme une demande de données externes ou un échange de données sur une connexion permanente avec un serveur. Le runtime JavaScript fournit également l’accès aux nouvelles API dans l’espace de noms `OfficeRuntime` utilisables dans des fonctions personnalisées ou par d’autres composants d’un complément pour stocker des données ou afficher une boîte de dialogue. Cet article décrit comment utiliser ces API dans les fonctions personnalisées et souligne également les considérations supplémentaires à prendre en compte lorsque vous développez des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="ba791-p101">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements. This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server. The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box. This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="ba791-108">Demander des données externes</span><span class="sxs-lookup"><span data-stu-id="ba791-108">Requesting external data</span></span>

<span data-ttu-id="ba791-p102">Au sein d’une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API telle que [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou à l’aide de [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API Web standard qui émet des demandes HTTP pour interagir avec les serveurs. Dans le runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant la [Politique de même origine](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et le simple [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="ba791-p102">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers. Within the JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

### <a name="xhr-example"></a><span data-ttu-id="ba791-111">Exemple XHR</span><span class="sxs-lookup"><span data-stu-id="ba791-111">XHR example</span></span>

<span data-ttu-id="ba791-p103">Dans l’exemple de code suivant, la fonction `getTemperature` appelle la fonction `sendWebRequest` pour obtenir la température d’une zone particulière basée sur l’ID de thermomètre. La fonction `sendWebRequest` utilise XHR pour émettre une demande `GET` à un point de terminaison qui peut fournir les données.</span><span class="sxs-lookup"><span data-stu-id="ba791-p103">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID. The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span> 

> [!NOTE] 
> <span data-ttu-id="ba791-p104">Lorsque vous utilisez fetch ou XHR, une nouvelle `Promise` JavaScript est renvoyée. Avant septembre 2018, vous deviez spécifier `OfficeExtension.Promise` pour utiliser les promesses dans l’API JavaScript pour Office, mais vous pouvez désormais simplement utiliser une `Promise` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ba791-p104">When using fetch or XHR, a new JavaScript `Promise` is returned. Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

```js
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
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="ba791-116">Réception de données via WebSockets</span><span class="sxs-lookup"><span data-stu-id="ba791-116">Receiving data via WebSockets</span></span>

<span data-ttu-id="ba791-p105">Au sein d’une fonction personnalisée, vous pouvez utiliser [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) pour échanger des données via une connexion permanente avec un serveur. En utilisant WebSocket, votre fonction personnalisée peut ouvrir une connexion avec un serveur et ensuite automatiquement recevoir des messages à partir du serveur lorsque certains événements se produisent, sans avoir à appeler explicitement le serveur pour des données.</span><span class="sxs-lookup"><span data-stu-id="ba791-p105">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server. By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="ba791-119">Exemple WebSockets</span><span class="sxs-lookup"><span data-stu-id="ba791-119">WebSockets example</span></span>

<span data-ttu-id="ba791-120">L’exemple de code suivant établit une  connexion `WebSocket`, puis enregistre chaque message entrant provenant du serveur.</span><span class="sxs-lookup"><span data-stu-id="ba791-120">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="ba791-121">Stockage des données et accès aux données</span><span class="sxs-lookup"><span data-stu-id="ba791-121">Storing and accessing data</span></span>

<span data-ttu-id="ba791-p106">Au sein d’une fonction personnalisée (ou au sein d’un composant du complément), vous pouvez stocker et accéder aux données à l’aide de l’objet `OfficeRuntime.AsyncStorage`. `AsyncStorage` est un système de stockage persistant, non chiffré, clé-valeur qui offre une alternative à [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé dans des fonctions personnalisées. Un complément peut stocker jusqu’à 10 Mo de données à l’aide de `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="ba791-p106">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object. `AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions. An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="ba791-125">Les méthodes suivantes sont disponibles sur l’objet `AsyncStorage` :</span><span class="sxs-lookup"><span data-stu-id="ba791-125">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a><span data-ttu-id="ba791-126">Exemple AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="ba791-126">AsyncStorage example</span></span> 

<span data-ttu-id="ba791-127">L’exemple de code suivant appelle la fonction `AsyncStorage.getItem` pour récupérer une valeur stockée.</span><span class="sxs-lookup"><span data-stu-id="ba791-127">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

```typescript
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
```

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="ba791-128">Affichage d’une boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="ba791-128">Open a dialog box</span></span>

<span data-ttu-id="ba791-p107">Au sein d’une fonction personnalisée (ou au sein d’un composant du complément), vous pouvez utiliser l’API `OfficeRuntime.displayWebDialogOptions` pour afficher une boîte de dialogue. Cette API de boîte de dialogue offre une alternative à l’[API Boîte de dialogue](../develop/dialog-api-in-office-add-ins.md) qui peut être utilisé dans les volets Office et les commandes de complément, mais pas dans les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="ba791-p107">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialogOptions` API to display a dialog box. This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="ba791-131">Exemple de l’API Boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="ba791-131">Dialog API example</span></span> 

<span data-ttu-id="ba791-132">Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la fonction `displayWebDialogOptions` de l'API Dialog pour afficher une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="ba791-132">In the following code sample, the `getTokenViaDialog` method uses the Dialog API’s `displayWebDialogOptions` method to open a dialog box.</span></span>

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
        OfficeRuntime.displayWebDialogOptions(url, {
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

## <a name="additional-considerations"></a><span data-ttu-id="ba791-133">Considérations supplémentaires</span><span class="sxs-lookup"><span data-stu-id="ba791-133">Additional considerations</span></span>

<span data-ttu-id="ba791-p108">Afin de créer un complément qui sera exécuté sur plusieurs plateformes (parmi les locataires clés des compléments Office), vous ne devez pas accéder au Document Object Model (DOM) dans des fonctions personnalisées ou utiliser des bibliothèques telles que jQuery qui s’appuient sur le modèle DOM. Dans Excel pour Windows, où les fonctions personnalisées utilisent le runtime JavaScript, les fonctions personnalisées ne peuvent pas accéder au DOM.</span><span class="sxs-lookup"><span data-stu-id="ba791-p108">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM. On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="ba791-136">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ba791-136">See also</span></span>

* [<span data-ttu-id="ba791-137">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="ba791-137">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="ba791-138">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ba791-138">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="ba791-139">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ba791-139">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="ba791-140">Didacticiel sur les fonctions personnalisées d’Excel</span><span class="sxs-lookup"><span data-stu-id="ba791-140">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
