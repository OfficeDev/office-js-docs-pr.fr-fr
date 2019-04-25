---
ms.date: 02/06/2019
description: Comprendre les scénarios clés dans le développement de fonctions personnalisées Excel qui utilisent le nouveau runtime JavaScript.
title: Runtime pour les fonctions personnalisées Excel (aperçu)
localization_priority: Normal
ms.openlocfilehash: 85024b6c3559e2a5f32bae9297787f8052bba38d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448216"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="d22a9-103">Runtime pour les fonctions personnalisées Excel (aperçu)</span><span class="sxs-lookup"><span data-stu-id="d22a9-103">Runtime for Excel custom functions (preview)</span></span>

<span data-ttu-id="d22a9-104">Les fonctions personnalisées utilisent un nouveau runtime JavaScript différent de celui utilisé par d’autres parties d’un complément, par exemple, le volet des tâches ou d’autres éléments d’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d22a9-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="d22a9-105">Ce runtime JavaScript est conçu pour optimiser les performances des calculs dans les fonctions personnalisées. Il comporte également de nouvelles API que vous pouvez utiliser pour effectuer des actions courantes sur le web au sein des fonctions personnalisées telles que la demande des données externes ou l’échange de données avec un serveur par le biais d’une connexion permanente.</span><span class="sxs-lookup"><span data-stu-id="d22a9-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span> <span data-ttu-id="d22a9-106">Le runtime JavaScript offre également l’accès aux nouvelles API dans l’espace de noms `OfficeRuntime` qui peut être utilisé au sein des fonctions personnalisées ou par d’autres parties d’un complément afin de stocker des données ou d’afficher une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="d22a9-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="d22a9-107">Cet article décrit comment utiliser ces API au sein des fonctions personnalisées et présente des facteurs supplémentaires à prendre en compte dans le cadre du développement de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="d22a9-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="d22a9-108">Demande de données externes</span><span class="sxs-lookup"><span data-stu-id="d22a9-108">Requesting external data</span></span>

<span data-ttu-id="d22a9-109">Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme [Récupérer](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou de [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.</span><span class="sxs-lookup"><span data-stu-id="d22a9-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="d22a9-110">Dans le runtime JavaScript utilisé par les fonctions personnalisées, XHR implémente des mesures de sécurité supplémentaires en imposant une [stratégie de même origine](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et un simple [cors](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="d22a9-110">Within the JavaScript runtime used by custom functions, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="d22a9-111">Notez qu’une implémentation CORS simples ne peut pas utiliser les cookies et prend uniquement en charge les méthodes simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="d22a9-111">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="d22a9-112">Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="d22a9-112">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="d22a9-113">Vous pouvez également utiliser un `Content-Type` en-tête dans un simple cors, à condition que `application/x-www-form-urlencoded`le `text/plain`type de `multipart/form-data`contenu soit,, ou.</span><span class="sxs-lookup"><span data-stu-id="d22a9-113">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="d22a9-114">Exemple avec XHR</span><span class="sxs-lookup"><span data-stu-id="d22a9-114">XHR example</span></span>

<span data-ttu-id="d22a9-115">Dans l’exemple de code suivant, la fonction `getTemperature` appelle la fonction `sendWebRequest` pour obtenir la température d’une zone spécifique en fonction de l’ID de thermomètre.</span><span class="sxs-lookup"><span data-stu-id="d22a9-115">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="d22a9-116">La fonction `sendWebRequest` utilise XHR pour émettre une demande `GET` à un point de terminaison qui peut fournir des données.</span><span class="sxs-lookup"><span data-stu-id="d22a9-116">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>

> [!NOTE] 
> <span data-ttu-id="d22a9-117">Lorsque vous utilisez l’API de récupération ou XHR, un nouvel élément `Promise` est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="d22a9-117">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="d22a9-118">Avant septembre 2018, vous deviez spécifier `OfficeExtension.Promise` pour utiliser des promesses au sein de l’API JavaScript Office, mais vous pouvez désormais simplement utiliser un élément `Promise` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d22a9-118">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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
        
        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="d22a9-119">Réception de données via WebSockets</span><span class="sxs-lookup"><span data-stu-id="d22a9-119">Receiving data via WebSockets</span></span>

<span data-ttu-id="d22a9-120">Dans une fonction personnalisée, vous pouvez utiliser [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) afin d’échanger des données avec un serveur via une connexion permanente.</span><span class="sxs-lookup"><span data-stu-id="d22a9-120">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="d22a9-121">Grâce à WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir les données.</span><span class="sxs-lookup"><span data-stu-id="d22a9-121">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="d22a9-122">Exemple avec WebSockets</span><span class="sxs-lookup"><span data-stu-id="d22a9-122">WebSockets example</span></span>

<span data-ttu-id="d22a9-123">L’exemple de code suivant établit une connexion `WebSocket`, puis consigne chaque message entrant provenant du serveur.</span><span class="sxs-lookup"><span data-stu-id="d22a9-123">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="d22a9-124">Accès aux données et stockage</span><span class="sxs-lookup"><span data-stu-id="d22a9-124">Storing and accessing data</span></span>

<span data-ttu-id="d22a9-125">Dans une fonction personnalisée (ou tout autre partie d’un complément), vous pouvez accéder aux données et les stocker à l’aide de l’objet `OfficeRuntime.AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="d22a9-125">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object.</span></span> <span data-ttu-id="d22a9-126">`AsyncStorage` est un système de stockage clé-valeur permanent et non chiffré qui permet de remplacer [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé au sein de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="d22a9-126">`AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="d22a9-127">Un complément peut stocker jusqu’à 10 Mo de données à l’aide de l’objet `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="d22a9-127">An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="d22a9-128">`AsyncStorage` est conçu comme une solution de stockage partagé, ce qui signifie que plusieurs parties d’un complément ont accès aux mêmes données.</span><span class="sxs-lookup"><span data-stu-id="d22a9-128">`AsyncStorage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="d22a9-129">Par exemple, les jetons destinés à l’authentification utilisateur peuvent être stockés dans `AsyncStorage`, car ce système de stockage est accessible à la fois par le biais d’une fonction personnalisée et via des éléments d’interface utilisateur de complément, par exemple, un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="d22a9-129">For example, tokens for user authentication may be stored in `AsyncStorage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="d22a9-130">De même, si deux compléments partagent le même domaine (par exemple, www.contoso.com/addin1, www.contoso.com/addin2), ils sont également autorisés à partager des informations entre eux via `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="d22a9-130">Similarly, if two add-ins share the same domain (e.g. www.contoso.com/addin1, www.contoso.com/addin2), they are also permitted to share information back and forth through `AsyncStorage`.</span></span> <span data-ttu-id="d22a9-131">Notez que les compléments ayant différents sous-domaines possèdent différentes instances de l’objet `AsyncStorage` (par exemple, subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span><span class="sxs-lookup"><span data-stu-id="d22a9-131">Note that add-ins which have different subdomains will have different instances of `AsyncStorage` (e.g. subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span></span> 

<span data-ttu-id="d22a9-132">Comme `AsyncStorage` peut être un emplacement partagé, il est important de savoir qu’il est possible de remplacer des paires clé-valeur.</span><span class="sxs-lookup"><span data-stu-id="d22a9-132">Because `AsyncStorage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="d22a9-133">Les méthodes suivantes sont disponibles avec l’objet `AsyncStorage` :</span><span class="sxs-lookup"><span data-stu-id="d22a9-133">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - <span data-ttu-id="d22a9-134">`multiRemove` : notez qu’il n’existe aucune implémentation d’une méthode pour effacer toutes les informations (par exemple, `clear`).</span><span class="sxs-lookup"><span data-stu-id="d22a9-134">`multiRemove`: You will note that there is no implementation of a method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="d22a9-135">À la place, vous devez utiliser l’objet `multiRemove` pour supprimer plusieurs entrées à la fois.</span><span class="sxs-lookup"><span data-stu-id="d22a9-135">Instead, you should instead use `multiRemove` to remove multiple entries at a time.</span></span>

### <a name="asyncstorage-example"></a><span data-ttu-id="d22a9-136">Exemple avec AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="d22a9-136">AsyncStorage example</span></span> 

<span data-ttu-id="d22a9-137">L'exemple de code suivant appelle `AsyncStorage.setItem` la fonction pour définir une clé et une `AsyncStorage`valeur.</span><span class="sxs-lookup"><span data-stu-id="d22a9-137">The following code sample calls the `AsyncStorage.setItem` function to set a key and value into `AsyncStorage`.</span></span>

```JavaScript
function StoreValue(key, value) {

  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="d22a9-138">Considérations supplémentaires</span><span class="sxs-lookup"><span data-stu-id="d22a9-138">Additional considerations</span></span>

<span data-ttu-id="d22a9-139">Pour créer un complément qui s’exécute sur plusieurs plateformes (l’un des clients clés des compléments Office), vous ne devez pas accéder au Document DOM (Object Model) dans les fonctions personnalisées ou utiliser de bibliothèques comme jQuery qui dépendent du DOM.</span><span class="sxs-lookup"><span data-stu-id="d22a9-139">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="d22a9-140">Sur Excel pour Windows, où les fonctions personnalisées utilisent le runtime JavaScript, les fonctions personnalisées ne peuvent pas accéder au DOM.</span><span class="sxs-lookup"><span data-stu-id="d22a9-140">On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="d22a9-141">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d22a9-141">See also</span></span>

* [<span data-ttu-id="d22a9-142">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="d22a9-142">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="d22a9-143">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d22a9-143">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d22a9-144">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d22a9-144">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="d22a9-145">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="d22a9-145">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="d22a9-146">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="d22a9-146">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
