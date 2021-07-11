---
ms.date: 03/15/2021
description: Demander, flux de données et annuler la diffusion en continu de données externes à votre classeur avec des fonctions personnalisées dans Excel
title: Recevoir et gérer des données à l’aide de fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: 60f09b791b13d34a4a7f307bb9677c9fcc72ee97
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349597"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>Recevoir et gérer des données à l’aide de fonctions personnalisées

L’une des façon dont les fonctions personnalisées améliorent la puissance d’Excel est qu’elles reçoivent des données en provenance d’emplacements autres que le classeur, par exemple, le web ou un serveur (via WebSockets). Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme[`Fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API)Récupérer ou à l’aide de`XmlHttpRequest` [ (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![GIF d’une fonction personnalisée qui diffuse l’heure à partir d’une API.](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a>Fonctions qui retournent des données provenant de sources externes

Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :

1. Renvoyer une promesse JavaScript à Excel.
2. Résoudre la promesse avec la valeur finale à l’aide de la fonction de rappel.

### <a name="fetch-example"></a>Exemple de récupération

Dans l’exemple de code suivant, la fonction atteint l’API hypothétique Contoso « Nombre de personnes dans l’espace », qui suit le nombre de personnes actuellement sur la station spatiale `webRequest` internationale. La fonction renvoie une promesse JavaScript et utilise la récupération pour demander des informations à l’API. Les données obtenues sont transformées en JSON et la `names` propriété est convertie en chaîne, ce qui permet de résoudre la promesse.

Lorsque vous développez vos propres fonctions, vous souhaitez peut-être effectuer une action si la requête Web ne se termine pas en temps voulu ou envisager de [regrouper plusieurs demandes API](custom-functions-batching.md).

```JS
/**
 * Requests the names of the people currently on the International Space Station from a hypothetical API.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace";
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

>[!NOTE]
>L’utilisation de `Fetch` permet d’éviter les rappels imbriqués et peut être préférable à XHR dans certains cas.

### <a name="xhr-example"></a>Exemple avec XHR

Dans l’exemple de code suivant, la fonction appelle l’API Github pour découvrir la quantité d’étoiles données au référentiel `getStarCount` d’un utilisateur particulier. Il s’agit d’une fonction asynchrone qui renvoie une promesse JavaScript. Lorsque des données sont obtenues à partir de l’appel Web, la promesse est résolue et renvoie les données à la cellule.

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

Pour déclarer une fonction de diffusion en continu, vous pouvez utiliser :

- `@streaming`Balise.
- Paramètre `CustomFunctions.StreamingInvocation` d’appel.

L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde. Notez les points suivants concernant ce code.

- Excel affiche chaque nouvelle valeur automatiquement à l’aide de la méthode `setResult`.
- Le deuxième paramètre d’entrée, l’invocation, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.
- Le rappel `onCanceled` définit la fonction qui s’exécute lorsque la fonction est annulée.
- La diffusion en continu n’est pas nécessairement liée à la création d’une requête Web : dans ce cas, la fonction ne crée pas de requête Web, mais continue d’obtenir des données à intervalles définis, de sorte qu’elle nécessite l’utilisation du paramètre `invocation` de diffusion en continu.

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

## <a name="canceling-a-function"></a>Annulation d’une fonction

Excel l’exécution d’une fonction dans les situations suivantes.

- L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.
- Un des arguments (entrées) de la fonction est modifié. Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.
- L’utilisateur déclenche manuellement le recalcul. Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.

Vous pouvez également définir une valeur de diffusion en continu par défaut pour gérer les cas lorsqu’une demande est effectuée, mais que vous êtes en mode hors connexion.

Notez qu’il existe également une catégorie de fonctions appelée fonctions annulables, qui ne sont _pas_ liées à des fonctions de diffusion en continu. Seules les fonctions personnalisées asynchrones qui retournent une valeur sont annulables. Les fonctions annulables permettent de mettre fin à une requête web au milieu d’une demande, en utilisant une commande [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) pour décider de l’action à effectuer lors de l’annulation. Déclarez une fonction annulable à l’aide de la balise `@cancelable`.

### <a name="using-an-invocation-parameter"></a>Utilisation d’un paramètre d’appel

Par défaut, le paramètre `invocation` est le dernier de toute fonction personnalisée. Le paramètre donne un contexte sur la cellule (par exemple, son adresse et son contenu) et vous permet `invocation` d’utiliser et `setResult` des `onCanceled` méthodes. Ces méthodes définissent l’action d’une fonction quand elle diffuse (`setResult`) ou est annulée (`onCanceled`).

Si vous utilisez TypeScript, le handler d’appel doit être de type [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) ou [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) .

## <a name="receiving-data-via-websockets"></a>Réception de données via WebSockets

Dans une fonction personnalisée, vous pouvez utiliser WebSockets afin d’échanger des données avec un serveur via une connexion permanente. À l’aide des WebSockets, votre fonction personnalisée peut ouvrir une connexion avec un serveur, puis recevoir automatiquement des messages à partir du serveur lorsque certains événements se produisent, sans avoir à interrogé explicitement le serveur pour obtenir des données.

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
- [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
