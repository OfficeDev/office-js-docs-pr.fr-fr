---
ms.date: 05/08/2019
description: Comprendre les scénarios clés dans le développement de fonctions personnalisées Excel qui utilisent le nouveau runtime JavaScript.
title: Runtime pour les fonctions personnalisées Excel
localization_priority: Normal
ms.openlocfilehash: bc8635e370a7b48af07bc169c2d2334ef0fba8ef
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33951976"
---
# <a name="runtime-for-excel-custom-functions"></a>Runtime pour les fonctions personnalisées Excel

Les fonctions personnalisées utilisent un nouveau runtime JavaScript différent de celui utilisé par d’autres parties d’un complément, par exemple, le volet des tâches ou d’autres éléments d’interface utilisateur. Ce runtime JavaScript est conçu pour optimiser les performances des calculs dans les fonctions personnalisées. Il comporte également de nouvelles API que vous pouvez utiliser pour effectuer des actions courantes sur le web au sein des fonctions personnalisées telles que la demande des données externes ou l’échange de données avec un serveur par le biais d’une connexion permanente.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Le runtime JavaScript offre également l’accès aux nouvelles API dans l’espace de noms `OfficeRuntime` qui peut être utilisé au sein des fonctions personnalisées ou par d’autres parties d’un complément afin de stocker des données ou d’afficher une boîte de dialogue. Cet article décrit comment utiliser ces API au sein des fonctions personnalisées et présente des facteurs supplémentaires à prendre en compte dans le cadre du développement de fonctions personnalisées.

## <a name="requesting-external-data"></a>Demande de données externes

Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme [Récupérer](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou de [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.

Dans le runtime JavaScript utilisé par les fonctions personnalisées, XHR implémente des mesures de sécurité supplémentaires en imposant une [stratégie de même origine](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et un simple [cors](https://www.w3.org/TR/cors/).

Notez qu’une implémentation CORS simples ne peut pas utiliser les cookies et prend uniquement en charge les méthodes simples (GET, HEAD, POST). Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`. Vous pouvez également utiliser un `Content-Type` en-tête dans un simple cors, à condition que `application/x-www-form-urlencoded`le `text/plain`type de `multipart/form-data`contenu soit,, ou.

### <a name="xhr-example"></a>Exemple avec XHR

Dans l’exemple de code suivant, la fonction `getTemperature` appelle la fonction `sendWebRequest` pour obtenir la température d’une zone spécifique en fonction de l’ID de thermomètre. La fonction `sendWebRequest` utilise XHR pour émettre une demande `GET` à un point de terminaison qui peut fournir des données.

> [!NOTE] 
> Lorsque vous utilisez l’API de récupération ou XHR, un nouvel élément `Promise` est renvoyé. Avant septembre 2018, vous deviez spécifier `OfficeExtension.Promise` pour utiliser des promesses au sein de l’API JavaScript Office, mais vous pouvez désormais simplement utiliser un élément `Promise` JavaScript.

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

## <a name="receiving-data-via-websockets"></a>Réception de données via WebSockets

Dans une fonction personnalisée, vous pouvez utiliser [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) afin d’échanger des données avec un serveur via une connexion permanente. Grâce à WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir les données.

### <a name="websockets-example"></a>Exemple avec WebSockets

L’exemple de code suivant établit une connexion `WebSocket`, puis consigne chaque message entrant provenant du serveur.

```JavaScript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>Accès aux données et stockage

Dans une fonction personnalisée (ou tout autre partie d’un complément), vous pouvez accéder aux données et les stocker à l’aide de l’objet `OfficeRuntime.storage`. `Storage` est un système de stockage clé-valeur permanent et non chiffré qui permet de remplacer [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé au sein de fonctions personnalisées. `Storage`offre 10 Mo de données par domaine. Les domaines peuvent être partagés par plusieurs compléments.

`Storage` est conçu comme une solution de stockage partagé, ce qui signifie que plusieurs parties d’un complément ont accès aux mêmes données. Par exemple, les jetons destinés à l’authentification utilisateur peuvent être stockés dans `storage`, car ce système de stockage est accessible à la fois par le biais d’une fonction personnalisée et via des éléments d’interface utilisateur de complément, par exemple, un volet des tâches. De même, si deux compléments partagent le même domaine (par exemple, www.contoso.com/addin1, www.contoso.com/addin2), ils sont également autorisés à partager des informations entre eux via `storage`. Notez que les compléments ayant différents sous-domaines possèdent différentes instances de l’objet `storage` (par exemple, subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).

Comme `storage` peut être un emplacement partagé, il est important de savoir qu’il est possible de remplacer des paires clé-valeur.

Les méthodes suivantes sont disponibles avec l’objet `storage` :

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

.[!NOTE]
> Il n’existe pas de méthode pour effacer toutes les informations `clear`(par exemple,). À la place, vous devez utiliser l’objet `removeItems` pour supprimer plusieurs entrées à la fois.

### <a name="officeruntimestorage-example"></a>Exemple de OfficeRuntime. Storage

L’exemple de code suivant appelle `OfficeRuntime.storage.setItem` la fonction pour définir une clé et une `storage`valeur.

```JavaScript
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a>Considérations supplémentaires

Pour créer un complément qui s’exécute sur plusieurs plateformes (l’un des clients clés des compléments Office), vous ne devez pas accéder au Document DOM (Object Model) dans les fonctions personnalisées ou utiliser de bibliothèques comme jQuery qui dépendent du DOM. Dans Excel sur Windows, où les fonctions personnalisées utilisent le runtime JavaScript, les fonctions personnalisées ne peuvent pas accéder au DOM.

## <a name="next-steps"></a>Étapes suivantes
Découvrez [les meilleures pratiques essentielles pour les fonctions personnalisées](custom-functions-best-practices.md).

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Architecture de fonctions](custom-functions-architecture.md)
* [Afficher une boîte de dialogue dans les fonctions personnalisées](custom-functions-dialog.md)
* [Didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md)
