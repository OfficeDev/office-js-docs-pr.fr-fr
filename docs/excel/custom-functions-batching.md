---
ms.date: 09/09/2022
description: Traitez ensemble les fonctions personnalisées pour réduire les appels réseau à un service à distance.
title: Le traitement par lots de fonctions personnalisées nécessite un service à distance
ms.localizationpriority: medium
ms.openlocfilehash: f779351789350bbc591b1b5d7a975ff9f70cda26
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234920"
---
# <a name="batch-custom-function-calls-for-a-remote-service"></a>Appels de fonction personnalisée Batch pour un service distant

Si vos fonctions personnalisées appellent un service à distance, vous pouvez utiliser un modèle le traitement par lots pour réduire le nombre d’appels réseau au service à distance. Pour réduire les boucles réseau, traitez par lots tous les appels en un seul appel du service web. Cette procédure est idéale lorsque la feuille de calcul est recalculée.

Par exemple, si une personne a utilisé votre fonction personnalisée dans 100 cellules d’une feuille de calcul et a ensuite recalculé la feuille de calcul, votre fonction personnalisée s’exécute 100 fois et effectue 100 appels réseau. Si vous utilisez un modèle de traitement par lots, les appels peuvent être combinés pour rassembler l’ensemble des 100 calculs en un seul appel réseau.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a>Afficher l’exemple terminé

Pour afficher l’exemple terminé, suivez cet article et collez les exemples de code dans votre propre projet. Par exemple, pour créer un projet de fonction personnalisée pour TypeScript, utilisez le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md), puis ajoutez tout le code de cet article au projet. Exécutez le code et essayez-le.

Vous pouvez également télécharger ou afficher l’exemple de projet complet au [modèle de traitement par lot de fonctions personnalisées](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching). Si vous voulez afficher l’ensemble du code avant de poursuivre la lecture, examinez le [fichier de script](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Excel-custom-functions/Batching/src/functions/functions.js).

## <a name="create-the-batching-pattern-in-this-article"></a>Créer le modèle le traitement par lots dans cet article

Pour configurer le traitement par lots pour vos fonctions personnalisées, vous devez écrire trois sections principales de code.

1. Opération [push](#add-the-_pushoperation-function) pour ajouter une nouvelle opération au lot d’appels chaque fois qu’Excel appelle votre fonction personnalisée.
2. [Fonction permettant d’effectuer la requête à distance](#make-the-remote-request) lorsque le lot est prêt.
3. [Code du serveur pour répondre à la demande de lot](#process-the-batch-call-on-the-remote-service), calculer tous les résultats de l’opération et retourner les valeurs.

Dans les sections suivantes, vous allez apprendre à construire le code un exemple à la fois. Il est recommandé de créer un tout nouveau projet de fonctions personnalisées à l’aide du [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) . Pour créer un projet, consultez [Prise en main du développement de fonctions personnalisées Excel](../quickstarts/excel-custom-functions-quickstart.md). Vous pouvez utiliser TypeScript ou JavaScript.

## <a name="batch-each-call-to-your-custom-function"></a>Traiter par lots chaque appel de votre fonction personnalisée

Vos fonctions personnalisées sont basées sur l’appel d’un service à distance pour effectuer l’opération et calculer le résultat dont elles ont besoin. Cette méthode leur offre un moyen de stocker chaque opération demandée dans un traitement par lots. Plus tard, vous apprendrez à créer une fonction `_pushOperation` pour traitement des opérations par lots. Tout d’abord, consultez l’exemple de code suivant pour découvrir la procédure d’appel de `_pushOperation` à partir de votre fonction personnalisée.

Dans le code suivant, la fonction personnalisée effectue une division, mais s’appuie sur un service à distance pour effectuer le calcul réel. Elle appelle `_pushOperation` pour traiter l’opération par lots, ainsi que d’autres opérations sur le service à distance. Elle nomme l’opération **div2**. Vous pouvez utiliser un schéma d’affectation de noms de votre choix pour les opérations tant que le service à distance utilise également le même schéma (plus d’informations sur le service à distance disponibles plus tard). En outre, les arguments dont le service à distance a besoin pour exécuter l’opération sont transmis.

### <a name="add-the-div2-custom-function"></a>Ajouter la fonction personnalisée div2

Ajoutez le code suivant à votre **fichierfunctions.js** ou **functions.ts** (selon si vous avez utilisé JavaScript ou TypeScript).

```javascript
/**
 * Divides two numbers using batching
 * @CustomFunction
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend, divisor) {
  return _pushOperation("div2", [dividend, divisor]);
}
```

### <a name="add-global-variables-for-tracking-batch-requests"></a>Ajouter des variables globales pour le suivi des requêtes par lots

Ensuite, ajoutez deux variables globales à votre **fichierfunctions.js** ou **functions.ts** . `_isBatchedRequestScheduled` est important ultérieurement pour le minutage des appels par lots au service distant.

```javascript
let _batch = [];
let _isBatchedRequestScheduled = false;
```

### <a name="add-the-_pushoperation-function"></a>Ajouter la `_pushOperation` fonction

Quand Excel appelle votre fonction personnalisée, vous devez envoyer l’opération dans le tableau de commandes. Le code **de fonction _pushOperation** suivant montre comment ajouter une nouvelle opération à partir d’une fonction personnalisée. Il crée une nouvelle entrée de traitement par lots, crée une nouvelle promesse de résolution ou de rejet de l’opération, et transmet l’entrée dans le tableau de traitement par lots.

Ce code vérifie également si un traitement par lots est planifié. Dans cet exemple, l’exécution de chaque traitement par lots est prévue toutes les 100 millisecondes. Vous pouvez ajuster cette valeur si nécessaire. Des valeurs supérieures entraînent l’envoi de traitements par lots plus grands au service à distance et l’augmentation du temps d’attente pour que l’utilisateur puisse afficher les résultats. Des valeurs inférieures ont tendance à envoyer davantage de traitements par lots au service à distance, mais avec un temps de réponse rapide pour les utilisateurs.

La fonction crée un objet **invocationEntry** qui contient le nom de chaîne de l’opération à exécuter. Par exemple, si vous aviez deux fonctions personnalisées nommées `multiply` et `divide`, vous pouvez les réutiliser comme noms d’opération dans vos entrées de traitement par lots. `args` contient les arguments passés à votre fonction personnalisée à partir d’Excel. Enfin, `resolve` ou `reject` les méthodes stockent une promesse contenant les informations retournées par le service distant.

Ajoutez le code suivant à votre **fichierfunctions.js** ou **functions.ts** .

```javascript
// This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.
function _pushOperation(op, args) {
  // Create an entry for your custom function.
  console.log("pushOperation");
  const invocationEntry = {
    operation: op, // e.g., sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g., 100 ms.
  if (!_isBatchedRequestScheduled) {
    console.log("schedule remote request");
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a>Créer la demande à distance

L’objectif de la fonction `_makeRemoteRequest` consiste à transmettre le traitement par lots d’opérations au service à distance, puis de renvoyer les résultats à chaque fonction personnalisée. Elle crée tout d’abord une copie du tableau de traitement par lots. Cela permet aux appels simultanés de fonctions personnalisées à partir d’Excel de commencer immédiatement le traitement par lots dans un nouveau tableau. La copie est ensuite transformée en un tableau plus simple qui ne contient pas les informations sur la promesse. Transmettre les promesses à un service à distance n’aurait aucun sens, car elles ne fonctionneraient pas. `_makeRemoteRequest` rejette ou résout chaque promesse en fonction de ce que le service à distance renvoie.

Ajoutez le code suivant à votre **fichierfunctions.js** ou **functions.ts** .

```javascript
// This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  try{
  console.log("makeRemoteRequest");
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });
  console.log("makeRemoteRequest2");
  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      console.log("responseBatch in fetchFromRemoteService");
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
          console.log("rejecting promise");
        } else {
          console.log("fulfilling promise");
          console.log(response);

          batchCopy[index].resolve(response.result);
        }
      });
    });
    console.log("makeRemoteRequest3");
  } catch (error) {
    console.log("error name:" + error.name);
    console.log("error message:" + error.message);
    console.log(error);
  }
}
```

### <a name="modify-_makeremoterequest-for-your-own-solution"></a>Modifier `_makeRemoteRequest` pour votre propre solution

La fonction `_makeRemoteRequest` appelle `_fetchFromRemoteService` qui, comme vous le verrez plus tard, est simplement une imitation représentant le service à distance. Cela facilite l’étude et l’exécution du code dans cet article. Toutefois, lorsque vous souhaitez utiliser ce code pour un service distant réel, vous devez apporter les modifications suivantes.

- Déterminez la manière dont vous souhaitez sérialiser les opérations de traitement par lots sur le réseau. Par exemple, vous souhaiterez peut-être placer le tableau dans un corps JSON.
- Au lieu d’appeler `_fetchFromRemoteService`, vous devez passer le véritable appel réseau au service à distance en transmettant le traitement par lots des opérations.

## <a name="process-the-batch-call-on-the-remote-service"></a>Traiter l’appel de traitement par lots sur le service à distance

La dernière étape consiste à gérer l’appel de traitement par lots dans le service à distance. L’exemple de code suivant affiche la fonction `_fetchFromRemoteService`. Cette fonction décompresse chaque opération, effectue l’opération spécifiée et renvoie les résultats. À des fins d’apprentissage dans cet article, la fonction `_fetchFromRemoteService` est conçue de manière à s’exécuter dans votre complément web et à imiter un service à distance. Vous pouvez ajouter ce code à votre **fichierfunctions.js** ou **functions.ts** afin de pouvoir étudier et exécuter tout le code de cet article sans avoir à configurer un service distant réel.

Ajoutez le code suivant à votre **fichierfunctions.js** ou **functions.ts** .

```javascript
// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a promise that may contain a batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.
async function _fetchFromRemoteService(requestBatch) {
  // Simulate a slow network request to the server.
  console.log("_fetchFromRemoteService");
  await pause(1000);
  console.log("postpause");
  return requestBatch.map((request) => {
    console.log("requestBatch server side");
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myResult = args[0] * args[1];
        console.log(myResult);
        return {
          result: myResult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms) {
  console.log("pause");
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a>Modifier `_fetchFromRemoteService` pour votre service à distance en direct

Pour modifier la `_fetchFromRemoteService` fonction à exécuter dans votre service distant en direct, apportez les modifications suivantes.

- Selon votre plateforme serveur (Node.js ou autres), mappez l’appel du réseau client à cette fonction.
- Supprimez la fonction `pause`, qui reproduit la latence du réseau dans le cadre de l’imitation.
- Modifiez la déclaration de fonction de manière à ce qu’elle fonctionne avec le paramètre transmis si le paramètre est modifié à des fins de réseau. Par exemple, au lieu d’un tableau, il peut s’agir d’un corps JSON d’opérations traitées par lots à traiter.
- Modifiez la fonction de manière à effectuer les opérations (ou appelez les fonctions qui effectuent les opérations).
- Appliquez un mécanisme d’authentification approprié. Veillez à ce que seuls les appelants corrects puissent accéder à la fonction.
- Placez le code dans le service à distance.

## <a name="next-steps"></a>Étapes suivantes

Découvrez [les différents paramètres](custom-functions-parameter-options.md) que vous pouvez utiliser dans vos fonctions personnalisées. Ou parcourez les concepts de base d’un [appel web via une fonction personnalisée](custom-functions-web-reqs.md).

## <a name="see-also"></a>Voir aussi

- [Valeurs volatiles dans les fonctions](custom-functions-volatile.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
