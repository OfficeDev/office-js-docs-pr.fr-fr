---
ms.date: 05/07/2019
description: Demander, flux de données et annuler la diffusion en continu de données externes à votre classeur avec des fonctions personnalisées dans Excel
title: Recevoir et gérer des données à l’aide de fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: 61f4d0fdaea4277faedddbe075a587fb23842c08
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659634"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>Recevoir et gérer des données à l’aide de fonctions personnalisées

L’une des façon dont les fonctions personnalisées améliorent la puissance d’Excel est qu’elles reçoivent des données en provenance d’emplacements autres que le classeur, par exemple, le web ou un serveur (via WebSockets). Les fonctions personnalisées peuvent demander des données par le biais XHR et extraire (`fetch`) des demandes ainsi que des flux de ces données en temps réel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

La documentation ci-dessous illustre certains exemples de requêtes web, mais pour créer une fonction diffusion en continu pour vous-même, essayez la [didacticiel relative aux fonctions personnalisées](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).

## <a name="functions-that-return-data-from-external-sources"></a>Fonctions qui retournent des données provenant de sources externes

Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :

1. Renvoyer une promesse JavaScript à Excel.
2. Résoudre la promesse avec la valeur finale à l’aide de la fonction de rappel.

Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme[`Fetch`](https://developer.mozilla.org/fr-FR/docs/Web/API/Fetch_API)Récupérer ou à l’aide de`XmlHttpRequest` [ (XHR)](https://developer.mozilla.org/fr-FR/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.

Dans le runtime JavaScript, XHR implémente des mesures de sécurité supplémentaires en exigeant la [politique de même origine (same-origin policy)](https://developer.mozilla.org/fr-FR/docs/Web/Security/Same-origin_policy) et le partage [CORS (partage des ressources cross-origin)](https://www.w3.org/TR/cors/) simple.

Notez qu’une implémentation CORS simples ne peut pas utiliser les cookies et prend uniquement en charge les méthodes simples (GET, HEAD, POST). Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`. Vous pouvez également utiliser un en-tête de Type de contenu dans CORS simple, autant que le type de contenu est `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.

### <a name="xhr-example"></a>Exemple avec XHR

Dans l’exemple de code suivant, la fonction**obtenirTemperature**appelle la fonction pour obtenir la température d’une zone spécifique en fonction de l’ID de thermomètre. La fonction sendWebRequest utilise XHR pour émettre une demande GET à un point de terminaison qui peut fournir des données.

```js
/**
 * Receives a temperature from an online source.
 * @customfunction
 * @param {number} thermometerID Identification number of the thermometer.
 */
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions.  
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

Pour un autre exemple d’une demande XHR avec davantage de contexte, voir la `getFile` fonction au sein de[ce fichier](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) dans le référentiel Github[Office-ajouter-dans-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).

### <a name="fetch-example"></a>Exemple de récupération

Dans l’exemple de code suivant, la fonction `stockPriceStream` utilise un symbole boursier pour obtenir le prix d’une action toutes les 1 000 millisecondes. Pour plus d’informations sur cet exemple voir le[didacticiel relatif aux fonctions personnalisées](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).

```js
/**
 * Streams a stock price.
 * @customfunction 
 * @param {string} ticker Stock ticker.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function stockPriceStream(ticker, invocation) {
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
                invocation.setResult(parseFloat(text));
            })
            .catch(function(error) {
                invocation.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    invocation.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receive-data-via-websockets"></a>Recevoir des données via WebSockets

Dans une fonction personnalisée, vous pouvez utiliser WebSockets afin d’échanger des données avec un serveur via une connexion permanente. Grâce à WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir les données.

### <a name="websockets-example"></a>Exemple avec WebSockets

L’exemple de code suivant établit une connexion WebSocket, puis consigne chaque message entrant provenant du serveur.

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="stream-and-cancel-functions"></a>Fonctions de diffusion et annulables

Les fonctions personnalisées de diffusion vous aident à copier des données vers des cellules à plusieurs reprises, sans exiger qu’un utilisateur actualise explicitement quoi que ce soit.

Les fonctions personnalisées annulables vous permettent d’annuler l’exécution d’une fonction personnalisée de diffusion pour réduire ses consommation de bande passante, de mémoire de travail et de temps processeur.

Pour déclarer une fonction comme étant de diffusion ou annulable, les indicateurs de commentaire JSDOC `@stream` ou `@cancelable`.

### <a name="using-an-invocation-parameter"></a>Utilisation d’un paramètre d’appel

Par défaut, le paramètre `invocation` est le dernier de toute fonction personnalisée. Le paramètre `invocation` fournit du contexte sur la cellule (par exemple, son adresse), et vous permet d’utiliser les méthodes `setResult` et `onCanceled`. Ces méthodes définissent l’action d’une fonction quand elle diffuse (`setResult`) ou est annulée (`onCanceled`).

Si vous utilisez TypeScript, le gestionnaire d’appel doit être de type `CustomFunctions.StreamingInvocation` ou `CustomFunctions.CancelableInvocation`.

### <a name="streaming-and-cancelable-function-example"></a>Exemple de fonction de diffusion et annulable
L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde. Tenez compte des informations suivantes à propos de ce code :

- Excel affiche chaque nouvelle valeur automatiquement à l’aide de la méthode `setResult`.
- Le deuxième paramètre d’entrée, l’invocation, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.
- Le rappel `onCanceled` définit la fonction qui s’exécute lorsque la fonction est annulée.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function(){
    clearInterval(timer);
    }
}
CustomFunctions.associate("INCREMENT", increment);
```

>[!NOTE]
> Excel annule l’exécution d’une fonction dans les situations suivantes :
>
> - L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.
> - Un des arguments (entrées) de la fonction est modifié. Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.
> - L’utilisateur déclenche manuellement le recalcul. Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.

## <a name="next-steps"></a>Étapes suivantes

* En savoir plus sur les [différents types de paramètres que vos fonctions peuvent utiliser](custom-functions-parameter-options.md).
* Découvrez comment [traiter par lots plusieurs appels d’API](custom-functions-batching.md).

## <a name="see-also"></a>Voir aussi

* [Valeurs volatiles dans les fonctions](custom-functions-volatile.md)
* [Créer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques pour l’utilisation des fonctions personnalisées](custom-functions-best-practices.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
