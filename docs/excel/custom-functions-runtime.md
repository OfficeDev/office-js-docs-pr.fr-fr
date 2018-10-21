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
# <a name="runtime-for-excel-custom-functions-preview"></a>Runtime pour les fonctions personnalisées Excel (préversion)

Les fonctions personnalisées utilisent un nouveau runtime JavaScript qui diffère du runtime utilisé par les autres composants d’un complément, tels que le volet Office ou d’autres éléments d’interface utilisateur. Ce runtime JavaScript est conçu pour optimiser les performances des calculs dans les fonctions personnalisées et expose les nouvelles API que vous pouvez utiliser pour effectuer des actions courantes sur le Web dans des fonctions personnalisées comme une demande de données externes ou un échange de données sur une connexion permanente avec un serveur. Le runtime JavaScript fournit également l’accès aux nouvelles API dans l’espace de noms `OfficeRuntime` utilisables dans des fonctions personnalisées ou par d’autres composants d’un complément pour stocker des données ou afficher une boîte de dialogue. Cet article décrit comment utiliser ces API dans les fonctions personnalisées et souligne également les considérations supplémentaires à prendre en compte lorsque vous développez des fonctions personnalisées.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>Demander des données externes

Au sein d’une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API telle que [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou à l’aide de [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API Web standard qui émet des demandes HTTP pour interagir avec les serveurs. Dans le runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant la [Politique de même origine](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et le simple [CORS](https://www.w3.org/TR/cors/).  

### <a name="xhr-example"></a>Exemple XHR

Dans l’exemple de code suivant, la fonction `getTemperature` appelle la fonction `sendWebRequest` pour obtenir la température d’une zone particulière basée sur l’ID de thermomètre. La fonction `sendWebRequest` utilise XHR pour émettre une demande `GET` à un point de terminaison qui peut fournir les données. 

> [!NOTE] 
> Lorsque vous utilisez fetch ou XHR, une nouvelle `Promise` JavaScript est renvoyée. Avant septembre 2018, vous deviez spécifier `OfficeExtension.Promise` pour utiliser les promesses dans l’API JavaScript pour Office, mais vous pouvez désormais simplement utiliser une `Promise` JavaScript.

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

Au sein d’une fonction personnalisée, vous pouvez utiliser [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) pour échanger des données via une connexion permanente avec un serveur. En utilisant WebSocket, votre fonction personnalisée peut ouvrir une connexion avec un serveur et ensuite automatiquement recevoir des messages à partir du serveur lorsque certains événements se produisent, sans avoir à appeler explicitement le serveur pour des données.

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

## <a name="storing-and-accessing-data"></a>Stockage des données et accès aux données

Au sein d’une fonction personnalisée (ou au sein d’un composant du complément), vous pouvez stocker et accéder aux données à l’aide de l’objet `OfficeRuntime.AsyncStorage`. `AsyncStorage` est un système de stockage persistant, non chiffré, clé-valeur qui offre une alternative à [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé dans des fonctions personnalisées. Un complément peut stocker jusqu’à 10 Mo de données à l’aide de `AsyncStorage`.

Les méthodes suivantes sont disponibles sur l’objet `AsyncStorage` :
 
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

## <a name="displaying-a-dialog-box"></a>Affichage d’une boîte de dialogue

Au sein d’une fonction personnalisée (ou au sein d’un composant du complément), vous pouvez utiliser l’API `OfficeRuntime.displayWebDialogOptions` pour afficher une boîte de dialogue. Cette API de boîte de dialogue offre une alternative à l’[API Boîte de dialogue](../develop/dialog-api-in-office-add-ins.md) qui peut être utilisé dans les volets Office et les commandes de complément, mais pas dans les fonctions personnalisées.

### <a name="dialog-api-example"></a>Exemple de l’API Boîte de dialogue 

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

Afin de créer un complément qui sera exécuté sur plusieurs plateformes (parmi les locataires clés des compléments Office), vous ne devez pas accéder au Document Object Model (DOM) dans des fonctions personnalisées ou utiliser des bibliothèques telles que jQuery qui s’appuient sur le modèle DOM. Dans Excel pour Windows, où les fonctions personnalisées utilisent le runtime JavaScript, les fonctions personnalisées ne peuvent pas accéder au DOM.

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Métadonnées des fonctions personnalisées](custom-functions-json.md)
* [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel sur les fonctions personnalisées d’Excel](excel-tutorial-custom-functions.md)
