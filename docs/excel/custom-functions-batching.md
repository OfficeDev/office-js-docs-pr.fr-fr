---
ms.date: 07/08/2021
description: Traitez ensemble les fonctions personnalisées pour réduire les appels réseau à un service à distance.
title: Le traitement par lots de fonctions personnalisées nécessite un service à distance
ms.localizationpriority: medium
ms.openlocfilehash: 0cf1a1df922a08f63af80498da2e357d285775e9
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074230"
---
# <a name="batch-custom-function-calls-for-a-remote-service"></a>Traitement par lots d’appels de fonction personnalisée pour un service distant

Si vos fonctions personnalisées appellent un service à distance, vous pouvez utiliser un modèle le traitement par lots pour réduire le nombre d’appels réseau au service à distance. Pour réduire les boucles réseau, traitez par lots tous les appels en un seul appel du service web. Cette procédure est idéale lorsque la feuille de calcul est recalculée.

Par exemple, si une personne a utilisé votre fonction personnalisée dans 100 cellules d’une feuille de calcul et a ensuite recalculé la feuille de calcul, votre fonction personnalisée s’exécute 100 fois et effectue 100 appels réseau. Si vous utilisez un modèle de traitement par lots, les appels peuvent être combinés pour rassembler l’ensemble des 100 calculs en un seul appel réseau.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a>Afficher l’exemple terminé

Vous pouvez suivre cet article et coller les exemples de code dans votre propre projet. Par exemple, vous pouvez utiliser le [générateur Yo Office](https://github.com/OfficeDev/generator-office)pour créer un projet de fonction personnalisée pour TypeScript, puis ajouter l’ensemble du code de cet article au projet. Vous pouvez alors exécuter le code, puis le tester.

Vous pouvez également télécharger ou afficher l’exemple de projet complet dans [Modèle de traitement par lots de fonction personnalisée](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching). Si vous voulez afficher l’ensemble du code avant de poursuivre la lecture, examinez le [fichier de script](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Excel-custom-functions/Batching/src/functions/functions.js).

## <a name="create-the-batching-pattern-in-this-article"></a>Créer le modèle le traitement par lots dans cet article

Pour configurer le traitement par lots pour vos fonctions personnalisées, vous devez écrire trois sections principales de code.

1. Une opération push pour ajouter une nouvelle opération au traitement par lots des appels chaque fois qu’Excel appelle votre fonction personnalisée.
2. Une fonction pour créer la demande à distance lorsque le traitement par lots est prêt.
3. Du code serveur pour répondre à la demande de traitement par lots, calculer tous les résultats de l’opération et retourner les valeurs.

Les sections suivantes vous montrent comment construire le premier exemple de code pas à pas. Vous ajoutez chaque exemple de code à votre fichier **functions.ts**. Il est recommandé de créer un projet de fonctions personnalisées à l’aide du générateur Yo Office. Pour créer un projet, consultez [Prise en main du développement de fonctions personnalisées Excel](../quickstarts/excel-custom-functions-quickstart.md) et utilisez TypeScript au lieu de JavaScript.

## <a name="batch-each-call-to-your-custom-function"></a>Traiter par lots chaque appel de votre fonction personnalisée

Vos fonctions personnalisées sont basées sur l’appel d’un service à distance pour effectuer l’opération et calculer le résultat dont elles ont besoin. Cette méthode leur offre un moyen de stocker chaque opération demandée dans un traitement par lots. Plus tard, vous apprendrez à créer une fonction `_pushOperation` pour traitement des opérations par lots. Tout d’abord, consultez l’exemple de code suivant pour découvrir la procédure d’appel de `_pushOperation` à partir de votre fonction personnalisée.

Dans le code suivant, la fonction personnalisée effectue une division, mais s’appuie sur un service à distance pour effectuer le calcul réel. Elle appelle `_pushOperation` pour traiter l’opération par lots, ainsi que d’autres opérations sur le service à distance. Elle nomme l’opération **div2**. Vous pouvez utiliser un schéma d’affectation de noms de votre choix pour les opérations tant que le service à distance utilise également le même schéma (plus d’informations sur le service à distance disponibles plus tard). En outre, les arguments dont le service à distance a besoin pour exécuter l’opération sont transmis.

### <a name="add-the-div2-custom-function-to-functionsts"></a>Ajouter la fonction personnalisée div2 à functions.ts

```typescript
/**
 * @CustomFunction
 * Divides two numbers using batching
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend: number, divisor: number) {
  return _pushOperation(
    "div2",
    [dividend, divisor]
  );
}
```

Ensuite, vous allez définir le tableau de traitement par lots qui va stocker toutes les opérations à transmettre en un seul appel réseau. Le code suivant montre comment définir une interface en décrivant chaque entrée de traitement par lots dans le tableau. L’interface définit une opération, qui est un nom de chaîne de l’opération à exécuter. Par exemple, si vous aviez deux fonctions personnalisées nommées `multiply` et `divide`, vous pouvez les réutiliser comme noms d’opération dans vos entrées de traitement par lots. `args` contient les arguments transmis à votre fonction personnalisée à partir d’Excel. Et enfin, `resolve` ou `reject` stocke une promesse en conservant les informations que le service à distance renvoie.

```typescript
interface IBatchEntry {
  operation: string;
  args: any[];
  resolve: (data: any) => void;
  reject: (error: Error) => void;
}
```

Ensuite, créez le tableau de traitement par lots qui utilise l’interface précédente. Pour savoir si un traitement par lots est prévu ou non, créez une variable `_isBatchedRequestSchedule`. Cette opération s’avère importante pour plus tard pour minuter les appels au service à distance.

```typescript
const _batch: IBatchEntry[] = [];
let _isBatchedRequestScheduled = false;
```

Enfin, lorsqu’Excel appelle votre fonction personnalisée, vous devez transmettre l’opération au tableau de traitement par lots. Le code suivant montre comment ajouter une nouvelle opération à partir d’une fonction personnalisée. Il crée une nouvelle entrée de traitement par lots, crée une nouvelle promesse de résolution ou de rejet de l’opération, et transmet l’entrée dans le tableau de traitement par lots.

Ce code vérifie également si un traitement par lots est planifié. Dans cet exemple, l’exécution de chaque traitement par lots est prévue toutes les 100 millisecondes. Vous pouvez ajuster cette valeur si nécessaire. Des valeurs supérieures entraînent l’envoi de traitements par lots plus grands au service à distance et l’augmentation du temps d’attente pour que l’utilisateur puisse afficher les résultats. Des valeurs inférieures ont tendance à envoyer davantage de traitements par lots au service à distance, mais avec un temps de réponse rapide pour les utilisateurs.

### <a name="add-the-_pushoperation-function-to-functionsts"></a>Ajouter la fonction `_pushOperation` à functions.ts

```typescript
function _pushOperation(op: string, args: any[]) {
  // Create an entry for your custom function.
  const invocationEntry: IBatchEntry = {
    operation: op, // e.g. sum
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
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a>Créer la demande à distance

L’objectif de la fonction `_makeRemoteRequest` consiste à transmettre le traitement par lots d’opérations au service à distance, puis de renvoyer les résultats à chaque fonction personnalisée. Elle crée tout d’abord une copie du tableau de traitement par lots. Cela permet aux appels simultanés de fonctions personnalisées à partir d’Excel de commencer immédiatement le traitement par lots dans un nouveau tableau. La copie est ensuite transformée en un tableau plus simple qui ne contient pas les informations sur la promesse. Transmettre les promesses à un service à distance n’aurait aucun sens, car elles ne fonctionneraient pas. `_makeRemoteRequest` rejette ou résout chaque promesse en fonction de ce que le service à distance renvoie.

### <a name="add-the-following-_makeremoterequest-method-to-functionsts"></a>Ajouter la méthode `_makeRemoteRequest` suivante à functions.ts

```typescript
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });

  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
        } else {
          console.log(response);
          batchCopy[index].resolve(response.result);
        }
      });
    });
}
```

### <a name="modify-_makeremoterequest-for-your-own-solution"></a>Modifier `_makeRemoteRequest` pour votre propre solution

La fonction `_makeRemoteRequest` appelle `_fetchFromRemoteService` qui, comme vous le verrez plus tard, est simplement une imitation représentant le service à distance. Cela facilite l’étude et l’exécution du code dans cet article. Toutefois, lorsque vous souhaitez utiliser ce code pour un service distant réel, vous devez apporter les modifications suivantes.

- Déterminez la manière dont vous souhaitez sérialiser les opérations de traitement par lots sur le réseau. Par exemple, vous souhaiterez peut-être placer le tableau dans un corps JSON.
- Au lieu d’appeler `_fetchFromRemoteService`, vous devez passer le véritable appel réseau au service à distance en transmettant le traitement par lots des opérations.

## <a name="process-the-batch-call-on-the-remote-service"></a>Traiter l’appel de traitement par lots sur le service à distance

La dernière étape consiste à gérer l’appel de traitement par lots dans le service à distance. L’exemple de code suivant affiche la fonction `_fetchFromRemoteService`. Cette fonction décompresse chaque opération, effectue l’opération spécifiée et renvoie les résultats. À des fins d’apprentissage dans cet article, la fonction `_fetchFromRemoteService` est conçue de manière à s’exécuter dans votre complément web et à imiter un service à distance. Vous pouvez ajouter ce code à votre fichier **functions.ts** afin d’examiner et d’exécuter l’ensemble du code de cet article sans devoir configurer de service à distance réel.

### <a name="add-the-following-_fetchfromremoteservice-function-to-functionsts"></a>Ajouter la fonction `_fetchFromRemoteService` suivante à functions.ts

```typescript
async function _fetchFromRemoteService(
  requestBatch: Array<{ operation: string, args: any[] }>
): Promise<IServerResponse[]> {
  // Simulate a slow network request to the server;
  await pause(1000);

  return requestBatch.map((request): IServerResponse => {
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myresult = args[0] * args[1];
        console.log(myresult);
        return {
          result: myresult
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

function pause(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a>Modifier `_fetchFromRemoteService` pour votre service à distance en direct

Pour modifier la fonction à exécuter dans votre service distant en `_fetchFromRemoteService` direct, a effectuer les modifications suivantes.

- Selon votre plateforme serveur (Node.js ou autres), mappez l’appel du réseau client à cette fonction.
- Supprimez la fonction `pause`, qui reproduit la latence du réseau dans le cadre de l’imitation.
- Modifiez la déclaration de fonction de manière à ce qu’elle fonctionne avec le paramètre transmis si le paramètre est modifié à des fins de réseau. Par exemple, au lieu d’un tableau, il peut s’agir d’un corps JSON d’opérations traitées par lots à traiter.
- Modifiez la fonction de manière à effectuer les opérations (ou appelez les fonctions qui effectuent les opérations).
- Appliquez un mécanisme d’authentification approprié. Veillez à ce que seuls les appelants corrects puissent accéder à la fonction.
- Placez le code dans le service à distance.

## <a name="next-steps"></a>Étapes suivantes

Découvrez [les différents paramètres](custom-functions-parameter-options.md) que vous pouvez utiliser dans vos fonctions personnalisées. Ou parcourez les concepts de base d’un [appel web via une fonction personnalisée](custom-functions-web-reqs.md).

## <a name="see-also"></a>Voir aussi

* [Valeurs volatiles dans les fonctions](custom-functions-volatile.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
