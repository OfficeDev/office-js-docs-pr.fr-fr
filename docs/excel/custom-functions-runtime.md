---
ms.date: 10/17/2018
description: Comprendre les scénarios clés dans le développement de fonctions personnalisées Excel qui utilisent le nouveau runtime JavaScript.
title: Runtime pour les fonctions personnalisées Excel
ms.openlocfilehash: 7261eb214e97a2a05e08571a31ac9b449df14083
ms.sourcegitcommit: 161a0625646a8c2ebaf1773c6369ee7cc96aa07b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/01/2018
ms.locfileid: "25891710"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="084db-103">Runtime pour les fonctions personnalisées Excel (aperçu)</span><span class="sxs-lookup"><span data-stu-id="084db-103">Runtime for Excel custom functions (preview)</span></span>

<span data-ttu-id="084db-104">Les fonctions personnalisées utilisent un nouveau runtime JavaScript différent de celui utilisé par d’autres parties d’un complément, par exemple, le volet des tâches ou d’autres éléments d’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="084db-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="084db-105">Ce runtime JavaScript est conçu pour optimiser les performances des calculs dans les fonctions personnalisées. Il comporte également de nouvelles API que vous pouvez utiliser pour effectuer des actions courantes sur le web au sein des fonctions personnalisées telles que la demande des données externes ou l’échange de données avec un serveur par le biais d’une connexion permanente.</span><span class="sxs-lookup"><span data-stu-id="084db-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span> <span data-ttu-id="084db-106">Le runtime JavaScript offre également l’accès aux nouvelles API dans l’espace de noms `OfficeRuntime` qui peut être utilisé au sein des fonctions personnalisées ou par d’autres parties d’un complément afin de stocker des données ou d’afficher une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="084db-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="084db-107">Cet article décrit comment utiliser ces API au sein des fonctions personnalisées et présente des facteurs supplémentaires à prendre en compte dans le cadre du développement de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="084db-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="084db-108">Demande de données externes</span><span class="sxs-lookup"><span data-stu-id="084db-108">Requesting external data</span></span>

<span data-ttu-id="084db-109">Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme [Récupérer](https://developer.mozilla.org/fr-FR/docs/Web/API/Fetch_API) ou de [XmlHttpRequest (XHR)](https://developer.mozilla.org/fr-FR/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.</span><span class="sxs-lookup"><span data-stu-id="084db-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/fr-FR/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/fr-FR/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="084db-110">Dans le runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant la [politique de même origine (same-origin policy)](https://developer.mozilla.org/fr-FR/docs/Web/Security/Same-origin_policy) et le partage [CORS (partage des ressources cross-origin)](https://www.w3.org/TR/cors/) simple.</span><span class="sxs-lookup"><span data-stu-id="084db-110">Within the JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/fr-FR/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

### <a name="xhr-example"></a><span data-ttu-id="084db-111">Exemple avec XHR</span><span class="sxs-lookup"><span data-stu-id="084db-111">XHR example</span></span>

<span data-ttu-id="084db-112">Dans l’exemple de code suivant, la fonction `getTemperature` appelle la fonction `sendWebRequest` pour obtenir la température d’une zone spécifique en fonction de l’ID de thermomètre.</span><span class="sxs-lookup"><span data-stu-id="084db-112">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="084db-113">La fonction `sendWebRequest` utilise XHR pour émettre une demande `GET` à un point de terminaison qui peut fournir des données.</span><span class="sxs-lookup"><span data-stu-id="084db-113">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span> 

> [!NOTE] 
> <span data-ttu-id="084db-114">Lorsque vous utilisez l’API de récupération ou XHR, un nouvel élément `Promise` est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="084db-114">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="084db-115">Avant septembre 2018, vous deviez spécifier `OfficeExtension.Promise` pour utiliser des promesses au sein de l’API JavaScript Office, mais vous pouvez désormais simplement utiliser un élément `Promise` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="084db-115">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="084db-116">Réception de données via WebSockets</span><span class="sxs-lookup"><span data-stu-id="084db-116">Receiving data via WebSockets</span></span>

<span data-ttu-id="084db-117">Dans une fonction personnalisée, vous pouvez utiliser [WebSockets](https://developer.mozilla.org/fr-FR/docs/Web/API/WebSockets_API) afin d’échanger des données avec un serveur via une connexion permanente.</span><span class="sxs-lookup"><span data-stu-id="084db-117">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/fr-FR/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="084db-118">Grâce à WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir les données.</span><span class="sxs-lookup"><span data-stu-id="084db-118">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="084db-119">Exemple avec WebSockets</span><span class="sxs-lookup"><span data-stu-id="084db-119">WebSockets example</span></span>

<span data-ttu-id="084db-120">L’exemple de code suivant établit une connexion `WebSocket`, puis consigne chaque message entrant provenant du serveur.</span><span class="sxs-lookup"><span data-stu-id="084db-120">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="084db-121">Accès aux données et stockage</span><span class="sxs-lookup"><span data-stu-id="084db-121">Storing and accessing data</span></span>

<span data-ttu-id="084db-122">Dans une fonction personnalisée (ou tout autre partie d’un complément), vous pouvez accéder aux données et les stocker à l’aide de l’objet `OfficeRuntime.AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="084db-122">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object.</span></span> <span data-ttu-id="084db-123">`AsyncStorage` est un système de stockage clé-valeur permanent et non chiffré qui permet de remplacer [localStorage](https://developer.mozilla.org/fr-FR/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé au sein de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="084db-123">`AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/fr-FR/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="084db-124">Un complément peut stocker jusqu’à 10 Mo de données à l’aide de l’objet `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="084db-124">An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="084db-125">`AsyncStorage` est conçu comme une solution de stockage partagé, ce qui signifie que plusieurs parties d’un complément ont accès aux mêmes données.</span><span class="sxs-lookup"><span data-stu-id="084db-125">`AsyncStorage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="084db-126">Par exemple, les jetons destinés à l’authentification utilisateur peuvent être stockés dans `AsyncStorage`, car ce système de stockage est accessible à la fois par le biais d’une fonction personnalisée et via des éléments d’interface utilisateur de complément, par exemple, un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="084db-126">For example, tokens for user authentication may be stored in `AsyncStorage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="084db-127">De même, si deux compléments partagent le même domaine (par exemple, www.contoso.com/addin1, www.contoso.com/addin2), ils sont également autorisés à partager des informations entre eux via `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="084db-127">Similarly, if two add-ins share the same domain (e.g. www.contoso.com/addin1, www.contoso.com/addin2), they are also permitted to share information back and forth through `AsyncStorage`.</span></span> <span data-ttu-id="084db-128">Notez que les compléments ayant différents sous-domaines possèdent différentes instances de l’objet `AsyncStorage` (par exemple, subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span><span class="sxs-lookup"><span data-stu-id="084db-128">Note that add-ins which have different subdomains will have different instances of `AsyncStorage` (e.g. subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span></span> 

<span data-ttu-id="084db-129">Comme `AsyncStorage` peut être un emplacement partagé, il est important de savoir qu’il est possible de remplacer des paires clé-valeur.</span><span class="sxs-lookup"><span data-stu-id="084db-129">Because `AsyncStorage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="084db-130">Les méthodes suivantes sont disponibles avec l’objet `AsyncStorage` :</span><span class="sxs-lookup"><span data-stu-id="084db-130">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - <span data-ttu-id="084db-131">`multiRemove` : notez qu’il n’existe aucune implémentation d’une méthode pour effacer toutes les informations (par exemple, `clear`).</span><span class="sxs-lookup"><span data-stu-id="084db-131">`multiRemove`: You will note that there is no implementation of a method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="084db-132">À la place, vous devez utiliser l’objet `multiRemove` pour supprimer plusieurs entrées à la fois.</span><span class="sxs-lookup"><span data-stu-id="084db-132">Instead, you should instead use `multiRemove` to remove multiple entries at a time.</span></span>

### <a name="asyncstorage-example"></a><span data-ttu-id="084db-133">Exemple avec AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="084db-133">AsyncStorage example</span></span> 

<span data-ttu-id="084db-134">L’exemple de code suivant appelle la fonction `AsyncStorage.getItem` pour récupérer une valeur du stockage.</span><span class="sxs-lookup"><span data-stu-id="084db-134">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

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

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="084db-135">Affichage d’une boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="084db-135">Displaying a dialog box</span></span>

<span data-ttu-id="084db-136">Dans une fonction personnalisée (ou tout autre partie d’un complément), vous pouvez utiliser l’API `OfficeRuntime.displayWebDialogOptions` pour afficher une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="084db-136">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialogOptions` API to display a dialog box.</span></span> <span data-ttu-id="084db-137">Cette API de boîte de dialogue permet de remplacer l’[API Boîte de dialogue](../develop/dialog-api-in-office-add-ins.md), qui peut être utilisée dans des volets des tâches et des commandes de complément, mais pas au sein de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="084db-137">This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="084db-138">Exemple d’API Boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="084db-138">Dialog API example</span></span> 

<span data-ttu-id="084db-139">Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la fonction `displayWebDialogOptions` de l’API Boîte de dialogue pour afficher une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="084db-139">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialogOptions` function to display a dialog box.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="084db-140">Considérations supplémentaires</span><span class="sxs-lookup"><span data-stu-id="084db-140">Additional considerations</span></span>

<span data-ttu-id="084db-141">Pour créer un complément qui s’exécute sur plusieurs plateformes (l’un des clients clés des compléments Office), vous ne devez pas accéder au Document DOM (Object Model) dans les fonctions personnalisées ou utiliser de bibliothèques comme jQuery qui dépendent du DOM.</span><span class="sxs-lookup"><span data-stu-id="084db-141">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="084db-142">Sur Excel pour Windows, où les fonctions personnalisées utilisent le runtime JavaScript, les fonctions personnalisées ne peuvent pas accéder au DOM.</span><span class="sxs-lookup"><span data-stu-id="084db-142">On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="084db-143">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="084db-143">See also</span></span>

* [<span data-ttu-id="084db-144">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="084db-144">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="084db-145">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="084db-145">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="084db-146">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="084db-146">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="084db-147">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="084db-147">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
