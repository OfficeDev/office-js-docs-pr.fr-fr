---
ms.date: 09/27/2018
description: Les fonctions personnalisées Excel utilisent un nouveau runtime JavaScript, qui diffère du runtime de contrôle WebView de compléments standard.
title: Runtime de fonctions personnalisées Excel
ms.openlocfilehash: 7489cd66851d1e0c24ef573ffa920b794cf749c2
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348758"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Runtime pour les fonctions personnalisées Excel (aperçu)

Les fonctions personnalisées étendent les fonctionnalités d’Excel à l’aide d’un nouveau runtime JavaScript qui utilise un moteur de JavaScript en bac à sable plutôt que dans un navigateur web. Puisque les fonctions personnalisées n’ont pas besoin d’afficher des éléments d’interface utilisateur, le nouveau runtime JavaScript est optimisé pour l’exécution de calculs, ce qui vous permet d’exécuter simultanément des milliers de fonctions personnalisées.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="key-facts-about-the-new-javascript-runtime"></a>Points clés concernant le nouveau runtime JavaScript 

Seules les fonctions personnalisées au sein d’un complément utiliseront le nouveau runtime JavaScript décrit dans cet article. Si un complément comprend d’autres composants tels que des volets Office et d’autres éléments d’interface utilisateur, en plus des fonctions personnalisées, ces autres composants du complément continueront à s’exécuter dans le module d’exécution WebView.  En outre : 

- Le runtime JavaScript ne fournit pas d’accès au Document Object Model (DOM) ni ne prend en charge des bibliothèques telles que jQuery qui s’appuient sur le modèle DOM.

- Une fonction personnalisée définie dans le fichier JavaScript d’un complément peut renvoyer une `Promise` JavaScript normale au lieu de renvoyer `OfficeExtension.Promise`.  

- Le fichier JSON qui spécifie les métadonnées de la fonction personnalisée n’a pas besoin de spécifier **sync** ou **async** dans **options**.

## <a name="new-apis"></a>Nouvelles API 

Le runtime JavaScript utilisé par les fonctions personnalisées comprend les APIs suivantes :

- [XHR](#xhr)
- [WebSockets](#websockets)
- [AsyncStorage](#asyncstorage)
- [API de boîte de dialogue](#dialog-api)

### <a name="xhr"></a>XHR

XHR est l’acronyme de [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs. Dans le nouveau runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant une [Stratégie d’Origine Identique](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et un [CORS](https://www.w3.org/TR/cors/) simple.  

Dans l’exemple de code suivant, la fonction `getTemperature()` envoie une demande web pour obtenir la température d’une zone particulière basée sur l’ID du thermomètre. La fonction `sendWebRequest()` utilise XHR pour émettre une demande `GET` à un point de terminaison qui peut fournir les données.  

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

### <a name="websockets"></a>WebSockets

[WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) est un protocole réseau qui crée une communication en temps réel entre un serveur et un ou plusieurs clients. Il est souvent utilisé pour les applications de conversation, car il permet de lire et d’écrire du texte simultanément.  

Comme indiqué dans l’exemple de code suivant, les fonctions personnalisées peuvent utiliser WebSocket. Dans cet exemple, le WebSocket enregistre chaque message qu’il reçoit.

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a>AsyncStorage

AsyncStorage est un système de stockage clé-valeur qui peut être utilisé pour stocker les jetons d’authentification. Il est :

- Persistant
- Non chiffré
- Asynchrone

AsyncStorage est globalement disponible pour tous les composants de votre complément. Pour les fonctions personnalisées, `AsyncStorage` est exposé comme un objet global. (Pour les autres composants de votre complément, tels que les volets Office et d’autres éléments qui utilisent le runtime WebView, AsyncStorage est exposé par le biais de `OfficeRuntime`.) Chaque complément a sa propre partition de stockage, avec une taille par défaut de 5 Mo. 

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
 
À ce stade, les méthodes `mergeItem` et `multiMerge` ne sont pas prises en charge.

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
}
```

### <a name="dialog-api"></a>API de boîte de dialogue

L’API de boîte de dialogue vous permet d’ouvrir une boîte de dialogue qui invite l’utilisateur à se connecter. Vous pouvez utiliser l’API de boîte de dialogue pour demander une authentification utilisateur via une ressource externe, telle que Google ou Facebook, avant que l’utilisateur ne puisse utiliser la fonction.   

Dans l’exemple de code suivant, la méthode `getTokenViaDialog()` utilise la méthode `displayWebDialog()` de l’API boîte de dialogue pour ouvrir une boîte de dialogue.

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
> L’API de boîte de dialogue décrite dans cette section fait partie du nouveau runtime JavaScript pour les fonctions personnalisées et peut être utilisée uniquement dans les fonctions personnalisées. Cette API est différente de l’[API de boîte de dialogue](../develop/dialog-api-in-office-add-ins.md) qui peut être utilisé dans les volets Office et les commandes de complément.

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Métadonnées des fonctions personnalisées](custom-functions-json.md)
* [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel sur les fonctions personnalisées d’Excel](excel-tutorial-custom-functions.md)
