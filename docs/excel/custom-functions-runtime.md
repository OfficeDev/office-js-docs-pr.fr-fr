---
ms.date: 01/08/2019
description: Comprendre les scénarios clés dans le développement de fonctions personnalisées Excel qui utilisent le nouveau runtime JavaScript.
title: Runtime pour les fonctions personnalisées Excel (aperçu)
localization_priority: Normal
ms.openlocfilehash: dd8158da4ebcccac61b8ab6958a101489bf5a668
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742316"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Runtime pour les fonctions personnalisées Excel (aperçu)

Les fonctions personnalisées utilisent un nouveau runtime JavaScript différent de celui utilisé par d’autres parties d’un complément, par exemple, le volet des tâches ou d’autres éléments d’interface utilisateur. Ce runtime JavaScript est conçu pour optimiser les performances des calculs dans les fonctions personnalisées. Il comporte également de nouvelles API que vous pouvez utiliser pour effectuer des actions courantes sur le web au sein des fonctions personnalisées telles que la demande des données externes ou l’échange de données avec un serveur par le biais d’une connexion permanente. Le runtime JavaScript offre également l’accès aux nouvelles API dans l’espace de noms `OfficeRuntime` qui peut être utilisé au sein des fonctions personnalisées ou par d’autres parties d’un complément afin de stocker des données ou d’afficher une boîte de dialogue. Cet article décrit comment utiliser ces API au sein des fonctions personnalisées et présente des facteurs supplémentaires à prendre en compte dans le cadre du développement de fonctions personnalisées.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>Demande de données externes

Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme [Récupérer](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou de [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs. Dans le runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant la [politique de même origine (same-origin policy)](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et le partage [CORS (partage des ressources cross-origin)](https://www.w3.org/TR/cors/) simple.  

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
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a>Réception de données via WebSockets

Dans une fonction personnalisée, vous pouvez utiliser [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) afin d’échanger des données avec un serveur via une connexion permanente. Grâce à WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir les données.

### <a name="websockets-example"></a>Exemple avec WebSockets

L’exemple de code suivant établit une connexion `WebSocket`, puis consigne chaque message entrant provenant du serveur. 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>Accès aux données et stockage

Dans une fonction personnalisée (ou tout autre partie d’un complément), vous pouvez accéder aux données et les stocker à l’aide de l’objet `OfficeRuntime.AsyncStorage`. `AsyncStorage` est un système de stockage clé-valeur permanent et non chiffré qui permet de remplacer [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé au sein de fonctions personnalisées. Un complément peut stocker jusqu’à 10 Mo de données à l’aide de l’objet `AsyncStorage`.

`AsyncStorage` est conçu comme une solution de stockage partagé, ce qui signifie que plusieurs parties d’un complément ont accès aux mêmes données. Par exemple, les jetons destinés à l’authentification utilisateur peuvent être stockés dans `AsyncStorage`, car ce système de stockage est accessible à la fois par le biais d’une fonction personnalisée et via des éléments d’interface utilisateur de complément, par exemple, un volet des tâches. De même, si deux compléments partagent le même domaine (par exemple, www.contoso.com/addin1, www.contoso.com/addin2), ils sont également autorisés à partager des informations entre eux via `AsyncStorage`. Notez que les compléments ayant différents sous-domaines possèdent différentes instances de l’objet `AsyncStorage` (par exemple, subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2). 

Comme `AsyncStorage` peut être un emplacement partagé, il est important de savoir qu’il est possible de remplacer des paires clé-valeur.

Les méthodes suivantes sont disponibles avec l’objet `AsyncStorage` :
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove` : notez qu’il n’existe aucune implémentation d’une méthode pour effacer toutes les informations (par exemple, `clear`). À la place, vous devez utiliser l’objet `multiRemove` pour supprimer plusieurs entrées à la fois.

### <a name="asyncstorage-example"></a>Exemple avec AsyncStorage 

L’exemple de code suivant appelle la fonction `AsyncStorage.getItem` pour récupérer une valeur du stockage.

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

Dans une fonction personnalisée (ou tout autre partie d’un complément), vous pouvez utiliser l’API `OfficeRuntime.displayWebDialog` pour afficher une boîte de dialogue. Cette API de boîte de dialogue permet de remplacer l’[API Boîte de dialogue](../develop/dialog-api-in-office-add-ins.md), qui peut être utilisée dans des volets des tâches et des commandes de complément, mais pas au sein de fonctions personnalisées.

### <a name="dialog-api-example"></a>Exemple d’API Boîte de dialogue

Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la fonction `displayWebDialog` de l’API Boîte de dialogue pour afficher une boîte de dialogue.

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
            dialog.close();
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

Pour créer un complément qui s’exécute sur plusieurs plateformes (l’un des clients clés des compléments Office), vous ne devez pas accéder au Document DOM (Object Model) dans les fonctions personnalisées ou utiliser de bibliothèques comme jQuery qui dépendent du DOM. Sur Excel pour Windows, où les fonctions personnalisées utilisent le runtime JavaScript, les fonctions personnalisées ne peuvent pas accéder au DOM.

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Fonctions personnalisées changelog](custom-functions-changelog.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
