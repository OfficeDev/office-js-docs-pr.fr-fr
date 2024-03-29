---
ms.date: 05/02/2022
description: Demandez, diffusez et annulez la diffusion en continu de données externes vers votre classeur avec des fonctions personnalisées dans Excel.
title: Recevoir et gérer des données à l’aide de fonctions personnalisées
ms.localizationpriority: medium
ms.openlocfilehash: fbe319e79d4cded5fe4b37ce5a654e633996f22a
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958544"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>Recevoir et gérer des données à l’aide de fonctions personnalisées

L’une des façons dont les fonctions personnalisées améliorent la puissance d’Excel consiste à recevoir des données à partir d’emplacements autres que le classeur, tels que le web ou un serveur (via [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API)). Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme[`Fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API)Récupérer ou à l’aide de`XmlHttpRequest` [ (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![GIF d’une fonction personnalisée qui diffuse l’heure à partir d’une API.](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a>Fonctions qui retournent des données provenant de sources externes

Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :

1. Renvoyer un [JavaScript `Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) à Excel.
2. Résolvez la `Promise` valeur finale à l’aide de la fonction de rappel.

### <a name="fetch-example"></a>Exemple de récupération

Dans l’exemple de code suivant, la `webRequest` fonction accède à une API externe hypothétique qui suit le nombre de personnes actuellement sur la Station spatiale internationale. La fonction retourne un Code JavaScript `Promise` et l’utilise pour demander des `fetch` informations à partir de l’API hypothétique. Les données obtenues sont transformées en JSON et la `names` propriété est convertie en chaîne, qui est utilisée pour résoudre la promesse.

Lorsque vous développez vos propres fonctions, vous souhaitez peut-être effectuer une action si la requête Web ne se termine pas en temps voulu ou envisager de [regrouper plusieurs demandes API](custom-functions-batching.md).

```JS
/**
 * Requests the names of the people currently on the International Space Station.
 * Note: This function requests data from a hypothetical URL. In practice, replace the URL with a data source for your scenario.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace"; // This is a hypothetical URL.
  return new Promise(function (resolve, reject) {
    fetch(url)
      .then(function (response){
        return response.json();
        }
      )
      .then(function (json) {
        resolve(JSON.stringify(json.names));
      })
  })
}
```

> [!NOTE]
> L’utilisation de `fetch` permet d’éviter les rappels imbriqués et peut être préférable à XHR dans certains cas.

### <a name="xhr-example"></a>Exemple avec XHR

Dans l’exemple de code suivant, la `getStarCount` fonction appelle l’API Github pour découvrir la quantité d’étoiles donnée au référentiel d’un utilisateur particulier. Il s’agit d’une fonction asynchrone qui retourne un Code JavaScript `Promise`. Lorsque des données sont obtenues à partir de l’appel web, la promesse est résolue, ce qui retourne les données à la cellule.

```TS
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param userName string name of organization or user.
 * @param repoName string name of the repository.
 * @return number of stars.
 */

async function getStarCount(userName: string, repoName: string) {

  const url = "https://api.github.com/repos/" + userName + "/" + repoName;

  let xhttp = new XMLHttpRequest();

  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;

      if (xhttp.status == 200) {
        resolve(JSON.parse(xhttp.responseText).watchers_count);
      } else {
        reject({
          status: xhttp.status,

          statusText: xhttp.statusText
        });
      }
    };

    xhttp.open("GET", url, true);

    xhttp.send();
  });
}
```

## <a name="make-a-streaming-function"></a>Créer une fonction de diffusion en continu

Les fonctions personnalisées de diffusion vous aident à copier des données vers des cellules à plusieurs reprises, sans exiger qu’un utilisateur actualise explicitement quoi que ce soit. Cela peut s’avérer utile pour vérifier les données actives d’un service en ligne, comme la fonction dans le [didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).

Pour déclarer une fonction de diffusion en continu, vous pouvez utiliser l’une des deux options suivantes.

- Balise `@streaming` .
- Paramètre d’appel `CustomFunctions.StreamingInvocation` .

L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde. Notez ce qui suit à propos de ce code.

- Excel affiche chaque nouvelle valeur automatiquement à l’aide de la méthode `setResult`.
- Le deuxième paramètre d’entrée `invocation`, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.
- Le `onCanceled` rappel définit la fonction qui s’exécute lorsque la fonction est annulée.
- La diffusion en continu n’est pas nécessairement liée à la création d’une requête web. Dans ce cas, la fonction n’effectue pas de requête web, mais reçoit toujours des données à intervalles définis. Elle nécessite donc l’utilisation du paramètre de streaming `invocation` .

```JS
/**
 * Increments a value once a second.
 * @customfunction INC increment
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

## <a name="cancel-a-function"></a>Annuler une fonction

Excel annule l’exécution d’une fonction dans les situations suivantes.

- L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.
- Un des arguments (entrées) de la fonction est modifié. Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.
- L’utilisateur déclenche manuellement le recalcul. Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.

Vous pouvez également définir une valeur de diffusion en continu par défaut pour gérer les cas lorsqu’une demande est effectuée, mais que vous êtes en mode hors connexion.

> [!NOTE]
> Il existe également une catégorie de fonctions appelées fonctions annulables, qui ne sont _pas_ liées aux fonctions de diffusion en continu. Seules les fonctions personnalisées asynchrones qui retournent une valeur sont annulables. Les fonctions annulables permettent de mettre fin à une requête web au milieu d’une demande, en utilisant une commande [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) pour décider de l’action à effectuer lors de l’annulation. Déclarez une fonction annulable à l’aide de la balise `@cancelable`.

### <a name="use-an-invocation-parameter"></a>Utiliser un paramètre d’appel

Par défaut, le paramètre `invocation` est le dernier de toute fonction personnalisée. Le `invocation` paramètre donne un contexte sur la cellule (par exemple, son adresse et son contenu) et vous permet d’utiliser la méthode et `onCanceled` l’événement `setResult` pour définir ce qu’une fonction fait lorsqu’elle diffuse (`setResult`) ou est annulée (`onCanceled`).

Si vous utilisez TypeScript, le gestionnaire d’appel doit être de type [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) ou [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation).

## <a name="receiving-data-via-websockets"></a>Réception de données via WebSockets

Dans une fonction personnalisée, vous pouvez utiliser [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API) afin d’échanger des données avec un serveur via une connexion permanente. À l’aide de WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages du serveur lorsque certains événements se produisent, sans avoir à interroger explicitement le serveur pour obtenir des données.

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

## <a name="next-steps"></a>Étapes suivantes

- En savoir plus sur les [différents types de paramètres que vos fonctions peuvent utiliser](custom-functions-parameter-options.md).
- Découvrez comment [traiter par lots plusieurs appels d’API](custom-functions-batching.md).

## <a name="see-also"></a>Voir aussi

- [Valeurs volatiles dans les fonctions](custom-functions-volatile.md)
- [Créer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
- [Créer manuellement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
