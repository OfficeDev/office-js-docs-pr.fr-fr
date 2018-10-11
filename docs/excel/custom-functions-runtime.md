---
ms.date: 10/03/2018
description: Comprendre les scénarios clés du développement de fonctions Excel personnalisées utilisant le nouveau runtime JavaScript.
title: Exécution de fonctions personnalisées Excel
ms.openlocfilehash: a48b02a8ca404b51740d9052d199da934eb9312e
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459104"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Runtime pour les fonctions personnalisées Excel (aperçu)

Les fonctions personnalisées utilisent un nouveau runtime JavaScript différent du runtime utilisé par les autres parties d’un complément, comme le volet de tâches ou autres éléments d’interface utilisateur. Ce runtime JavaScript est conçu pour optimiser la performance des calculs dans les fonctions personnalisées et expose de nouvelles API que vous pouvez utiliser pour effectuer des actions web ordinaires comme des requêtes de données externes ou des échanges de données sur une connexion permanente avec un serveur. Le runtime JavaScript offre également un accès à de nouvelles API dans le namespace `OfficeRuntime` qui peut être utilisé dans le cadre de fonctions personnalisées ou par d'autres parties d'un complément pour stocker des données ou pour afficher une boîte de dialogue. Cet article décrit comment utiliser ces API dans le cadre de fonctions personnalisées et souligne aussi des considérations supplémentaires à avoir lors du développement de fonctions personnalisées.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>Requête de données externes

Au sein d’une fonction personnalisée, vous pouvez faire des requêtes données externes à l’aide d’une API comme [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou avec [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs. Dans le nouveau runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant une [Stratégie d’Origine Identique](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et un [CORS](https://www.w3.org/TR/cors/) simple.  

### <a name="xhr-example"></a>Exemple XHR

Dans l’exemple de code suivant, la fonction `getTemperature` appelle la  fonction `sendWebRequest` pour obtenir la température d’une zone particulière basée sur l’ID de thermomètre. La fonction `sendWebRequest` utilise XHR pour émettre une demande `GET` à un point de terminaison qui peut fournir les données. 

> [!NOTE] 
> Lorsque vous utilisez l’extraction ou XHR, un nouveau JavaScript `Promise` est renvoyé. Avant septembre 2018, il fallait spécifier  `OfficeExtension.Promise`  pour utiliser les promesses dans l'API JavaScript Office, mais il est maintenant possible d'utiliser juste un JavaScript `Promise`.

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

## <a name="receiving-data-via-websockets"></a>Réception de données via WebSockets

Au sein d’une fonction personnalisée, vous pouvez utiliser [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) pour échanger des données via une connexion permanente avec un serveur. En utilisant WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur puis recevoir automatiquement des messages du serveur lorsque certains évènements se produisent, sans devoir interroger le serveur pour les données.

### <a name="websockets-example"></a>Exemple WebSockets

L’exemple de code suivant établit une  connexion `WebSocket`, puis enregistre chaque message entrant provenant du serveur. 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>Stockage et accès aux données

Au sein d’une fonction personnalisée (ou au sein d’une autre partie d'un complément), vous pouvez stocker et accéder aux données à l’aide de l'objet `OfficeRuntime.AsyncStorage`. `AsyncStorage` est un système de stockage permanent, non crypté, clé/valeur qui offre une alternative à  [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé dans les fonctions personnalisées. Un complément peut stocker jusqu'à 10 Mo de données avec `AsyncStorage`.

Les méthodes suivantes sont disponibles sur l’objet `AsyncStorage` :
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a>Exemple AsyncStorage 

L’exemple de code suivant appelle la fonction `AsyncStorage.getItem` pour récupérer une valeur stockée.

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

## <a name="displaying-a-dialog-box"></a>Affichage d'une boîte de dialogue

Au sein d’une fonction personnalisée (ou au sein d’une autre partie d'un complément), vous pouvez utiliser l'API `OfficeRuntime.displayWebDialogOptions` pour afficher une boîte de dialogue. Cette API de dialogue offre une alternative à l' [API Dialog](../develop/dialog-api-in-office-add-ins.md) qui peut être utilisé dans les volets Office et les commandes de complément, mais pas dans les fonctions personnalisées.

### <a name="dialog-api-example"></a>Exemple API Dialog 

Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la fonction `displayWebDialogOptions` de l'API Dialog pour afficher une boîte de dialogue.

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

## <a name="additional-considerations"></a>Considérations supplémentaires

Pour créer un complément qui s’exécute sur plusieurs plates-formes (l’un des principaux clients des compléments Office), vous ne devez pas accéder au DOM (Document Object Model) dans les fonctions personnalisées ni utiliser des bibliothèques comme jQuery qui s’appuient sur le modèle DOM. Dans Excel pour Windows, où les fonctions personnalisées utilisent le runtime JavaScript, les fonctions personnalisées ne peuvent pas accéder au DOM.

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Métadonnées des fonctions personnalisées](custom-functions-json.md)
* [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel sur les fonctions personnalisées d’Excel](excel-tutorial-custom-functions.md)
