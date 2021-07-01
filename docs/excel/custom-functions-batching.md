---
ms.date: 07/10/2019
description: Traitez ensemble les fonctions personnalisées pour réduire les appels réseau à un service à distance.
title: Le traitement par lots de fonctions personnalisées nécessite un service à distance
localization_priority: Normal
ms.openlocfilehash: 2ad9532fab26ff3ec8289a8892d518ab2570c6d6
ms.sourcegitcommit: d372de1a25dbad983fa9872c6af19a916f63f317
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/30/2021
ms.locfileid: "53204996"
---
# <a name="batching-custom-function-calls-for-a-remote-service"></a><span data-ttu-id="8578b-103">Le traitement par lots de fonctions personnalisées nécessite un service à distance</span><span class="sxs-lookup"><span data-stu-id="8578b-103">Batching custom function calls for a remote service</span></span>

<span data-ttu-id="8578b-104">Si vos fonctions personnalisées appellent un service à distance, vous pouvez utiliser un modèle le traitement par lots pour réduire le nombre d’appels réseau au service à distance.</span><span class="sxs-lookup"><span data-stu-id="8578b-104">If your custom functions call a remote service you can use a batching pattern to reduce the number of network calls to the remote service.</span></span> <span data-ttu-id="8578b-105">Pour réduire les boucles réseau, traitez par lots tous les appels en un seul appel du service web.</span><span class="sxs-lookup"><span data-stu-id="8578b-105">To reduce network round trips you batch all the calls into a single call to the web service.</span></span> <span data-ttu-id="8578b-106">Cette procédure est idéale lorsque la feuille de calcul est recalculée.</span><span class="sxs-lookup"><span data-stu-id="8578b-106">This is ideal when the spreadsheet is recalculated.</span></span>

<span data-ttu-id="8578b-107">Par exemple, si une personne a utilisé votre fonction personnalisée dans 100 cellules d’une feuille de calcul et a ensuite recalculé la feuille de calcul, votre fonction personnalisée s’exécute 100 fois et effectue 100 appels réseau.</span><span class="sxs-lookup"><span data-stu-id="8578b-107">For example, if someone used your custom function in 100 cells in a spreadsheet, and then recalculated the spreadsheet, your custom function would run 100 times and make 100 network calls.</span></span> <span data-ttu-id="8578b-108">Si vous utilisez un modèle de traitement par lots, les appels peuvent être combinés pour rassembler l’ensemble des 100 calculs en un seul appel réseau.</span><span class="sxs-lookup"><span data-stu-id="8578b-108">By using a batching pattern, the calls can be combined to make all 100 calculations in a single network call.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a><span data-ttu-id="8578b-109">Afficher l’exemple terminé</span><span class="sxs-lookup"><span data-stu-id="8578b-109">View the completed sample</span></span>

<span data-ttu-id="8578b-110">Vous pouvez suivre cet article et coller les exemples de code dans votre propre projet.</span><span class="sxs-lookup"><span data-stu-id="8578b-110">You can follow this article and paste the code examples into your own project.</span></span> <span data-ttu-id="8578b-111">Par exemple, vous pouvez utiliser le [générateur Yo Office](https://github.com/OfficeDev/generator-office)pour créer un projet de fonction personnalisée pour TypeScript, puis ajouter l’ensemble du code de cet article au projet.</span><span class="sxs-lookup"><span data-stu-id="8578b-111">For example, you can use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create a new custom function project for TypeScript, then add all the code from this article to the project.</span></span> <span data-ttu-id="8578b-112">Vous pouvez alors exécuter le code, puis le tester.</span><span class="sxs-lookup"><span data-stu-id="8578b-112">You can then run the code and try it out.</span></span>

<span data-ttu-id="8578b-113">Vous pouvez également télécharger ou afficher l’exemple de projet complet dans [Modèle de traitement par lots de fonction personnalisée](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching).</span><span class="sxs-lookup"><span data-stu-id="8578b-113">Also, you can download or view the complete sample project at [Custom function batching pattern](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching).</span></span> <span data-ttu-id="8578b-114">Si vous voulez afficher l’ensemble du code avant de poursuivre la lecture, examinez le [fichier de script](https://github.com/OfficeDev/PnP-OfficeAddins/blob/main/Excel-custom-functions/Batching/src/functions/functions.js).</span><span class="sxs-lookup"><span data-stu-id="8578b-114">If you want to view the code in whole before reading any further, take a look at the [script file](https://github.com/OfficeDev/PnP-OfficeAddins/blob/main/Excel-custom-functions/Batching/src/functions/functions.js).</span></span>

## <a name="create-the-batching-pattern-in-this-article"></a><span data-ttu-id="8578b-115">Créer le modèle le traitement par lots dans cet article</span><span class="sxs-lookup"><span data-stu-id="8578b-115">Create the batching pattern in this article</span></span>

<span data-ttu-id="8578b-116">Pour configurer le traitement par lots pour vos fonctions personnalisées, vous devez écrire trois sections principales de code.</span><span class="sxs-lookup"><span data-stu-id="8578b-116">To set up batching for your custom functions you'll need to write three main sections of code.</span></span>

1. <span data-ttu-id="8578b-117">Une opération push pour ajouter une nouvelle opération au traitement par lots des appels chaque fois qu’Excel appelle votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8578b-117">A push operation to add a new operation to the batch of calls each time Excel calls your custom function.</span></span>
2. <span data-ttu-id="8578b-118">Une fonction pour créer la demande à distance lorsque le traitement par lots est prêt.</span><span class="sxs-lookup"><span data-stu-id="8578b-118">A function to make the remote request when the batch is ready.</span></span>
3. <span data-ttu-id="8578b-119">Du code serveur pour répondre à la demande de traitement par lots, calculer tous les résultats de l’opération et retourner les valeurs.</span><span class="sxs-lookup"><span data-stu-id="8578b-119">Server code to respond to the batch request, calculate all of the operation results, and return the values.</span></span>

<span data-ttu-id="8578b-120">Les sections suivantes vous montrent comment construire le premier exemple de code pas à pas.</span><span class="sxs-lookup"><span data-stu-id="8578b-120">In the following sections you will be shown how to construct the code one example at a time.</span></span> <span data-ttu-id="8578b-121">Vous ajoutez chaque exemple de code à votre fichier **functions.ts**.</span><span class="sxs-lookup"><span data-stu-id="8578b-121">You'll add each code example to your **functions.ts** file.</span></span> <span data-ttu-id="8578b-122">Il est recommandé de créer un projet de fonctions personnalisées à l’aide du générateur Yo Office.</span><span class="sxs-lookup"><span data-stu-id="8578b-122">It's recommended you create a brand new custom functions project using the Yo Office generator.</span></span> <span data-ttu-id="8578b-123">Pour créer un projet, consultez [Prise en main du développement de fonctions personnalisées Excel](../quickstarts/excel-custom-functions-quickstart.md) et utilisez TypeScript au lieu de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8578b-123">To create a new project see [Get started developing Excel custom functions](../quickstarts/excel-custom-functions-quickstart.md) and use TypeScript instead of JavaScript.</span></span>

## <a name="batch-each-call-to-your-custom-function"></a><span data-ttu-id="8578b-124">Traiter par lots chaque appel de votre fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="8578b-124">Batch each call to your custom function</span></span>

<span data-ttu-id="8578b-125">Vos fonctions personnalisées sont basées sur l’appel d’un service à distance pour effectuer l’opération et calculer le résultat dont elles ont besoin.</span><span class="sxs-lookup"><span data-stu-id="8578b-125">Your custom functions work by calling a remote service to perform the operation and calculate the result they need.</span></span> <span data-ttu-id="8578b-126">Cette méthode leur offre un moyen de stocker chaque opération demandée dans un traitement par lots.</span><span class="sxs-lookup"><span data-stu-id="8578b-126">This provides a way for them to store each requested operation into a batch.</span></span> <span data-ttu-id="8578b-127">Plus tard, vous apprendrez à créer une fonction `_pushOperation` pour traitement des opérations par lots.</span><span class="sxs-lookup"><span data-stu-id="8578b-127">Later you'll see how to create a `_pushOperation` function to batch the operations.</span></span> <span data-ttu-id="8578b-128">Tout d’abord, consultez l’exemple de code suivant pour découvrir la procédure d’appel de `_pushOperation` à partir de votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8578b-128">First, take a look at the following code example to see how to call `_pushOperation` from your custom function.</span></span>

<span data-ttu-id="8578b-129">Dans le code suivant, la fonction personnalisée effectue une division, mais s’appuie sur un service à distance pour effectuer le calcul réel.</span><span class="sxs-lookup"><span data-stu-id="8578b-129">In the following code, the custom function performs division but relies on a remote service to do the actual calculation.</span></span> <span data-ttu-id="8578b-130">Elle appelle `_pushOperation` pour traiter l’opération par lots, ainsi que d’autres opérations sur le service à distance.</span><span class="sxs-lookup"><span data-stu-id="8578b-130">It calls `_pushOperation` to batch the operation along with other operations to the remote service.</span></span> <span data-ttu-id="8578b-131">Elle nomme l’opération **div2**.</span><span class="sxs-lookup"><span data-stu-id="8578b-131">It names the operation **div2**.</span></span> <span data-ttu-id="8578b-132">Vous pouvez utiliser un schéma d’affectation de noms de votre choix pour les opérations tant que le service à distance utilise également le même schéma (plus d’informations sur le service à distance disponibles plus tard).</span><span class="sxs-lookup"><span data-stu-id="8578b-132">You can use any naming scheme you want for operations as long as the remote service is also using the same scheme (more on the remote service later).</span></span> <span data-ttu-id="8578b-133">En outre, les arguments dont le service à distance a besoin pour exécuter l’opération sont transmis.</span><span class="sxs-lookup"><span data-stu-id="8578b-133">Also, the arguments the remote service will need to run the operation are passed.</span></span>

### <a name="add-the-div2-custom-function-to-functionsts"></a><span data-ttu-id="8578b-134">Ajouter la fonction personnalisée div2 à functions.ts</span><span class="sxs-lookup"><span data-stu-id="8578b-134">Add the div2 custom function to functions.ts</span></span>

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

<span data-ttu-id="8578b-135">Ensuite, vous allez définir le tableau de traitement par lots qui va stocker toutes les opérations à transmettre en un seul appel réseau.</span><span class="sxs-lookup"><span data-stu-id="8578b-135">Next, you will define the batch array which will store all operations to be passed in one network call.</span></span> <span data-ttu-id="8578b-136">Le code suivant montre comment définir une interface en décrivant chaque entrée de traitement par lots dans le tableau.</span><span class="sxs-lookup"><span data-stu-id="8578b-136">The following code shows how to define an interface describing each batch entry in the array.</span></span> <span data-ttu-id="8578b-137">L’interface définit une opération, qui est un nom de chaîne de l’opération à exécuter.</span><span class="sxs-lookup"><span data-stu-id="8578b-137">The interface defines an operation, which is a string name of which operation to run.</span></span> <span data-ttu-id="8578b-138">Par exemple, si vous aviez deux fonctions personnalisées nommées `multiply` et `divide`, vous pouvez les réutiliser comme noms d’opération dans vos entrées de traitement par lots.</span><span class="sxs-lookup"><span data-stu-id="8578b-138">For example, if you had two custom functions named `multiply` and `divide`, you could reuse those as the operation names in your batch entries.</span></span> <span data-ttu-id="8578b-139">`args` contient les arguments transmis à votre fonction personnalisée à partir d’Excel.</span><span class="sxs-lookup"><span data-stu-id="8578b-139">`args` will hold the arguments that were passed to your custom function from Excel.</span></span> <span data-ttu-id="8578b-140">Et enfin, `resolve` ou `reject` stocke une promesse en conservant les informations que le service à distance renvoie.</span><span class="sxs-lookup"><span data-stu-id="8578b-140">And finally, `resolve` or `reject` will store a promise holding the information the remote service returns.</span></span>

```typescript
interface IBatchEntry {
  operation: string;
  args: any[];
  resolve: (data: any) => void;
  reject: (error: Error) => void;
}
```

<span data-ttu-id="8578b-141">Ensuite, créez le tableau de traitement par lots qui utilise l’interface précédente.</span><span class="sxs-lookup"><span data-stu-id="8578b-141">Next, create the batch array that uses the previous interface.</span></span> <span data-ttu-id="8578b-142">Pour savoir si un traitement par lots est prévu ou non, créez une variable `_isBatchedRequestSchedule`.</span><span class="sxs-lookup"><span data-stu-id="8578b-142">To track if a batch is scheduled or not, create an `_isBatchedRequestSchedule` variable.</span></span> <span data-ttu-id="8578b-143">Cette opération s’avère importante pour plus tard pour minuter les appels au service à distance.</span><span class="sxs-lookup"><span data-stu-id="8578b-143">This will be important later for timing batch calls to the remote service.</span></span>

```typescript
const _batch: IBatchEntry[] = [];
let _isBatchedRequestScheduled = false;
```

<span data-ttu-id="8578b-144">Enfin, lorsqu’Excel appelle votre fonction personnalisée, vous devez transmettre l’opération au tableau de traitement par lots.</span><span class="sxs-lookup"><span data-stu-id="8578b-144">Finally when Excel calls your custom function, you need to push the operation into the batch array.</span></span> <span data-ttu-id="8578b-145">Le code suivant montre comment ajouter une nouvelle opération à partir d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8578b-145">The following code shows how to add a new operation from a custom function.</span></span> <span data-ttu-id="8578b-146">Il crée une nouvelle entrée de traitement par lots, crée une nouvelle promesse de résolution ou de rejet de l’opération, et transmet l’entrée dans le tableau de traitement par lots.</span><span class="sxs-lookup"><span data-stu-id="8578b-146">It creates a new batch entry, creates a new promise to resolve or reject the operation, and pushes the entry into the batch array.</span></span>

<span data-ttu-id="8578b-147">Ce code vérifie également si un traitement par lots est planifié.</span><span class="sxs-lookup"><span data-stu-id="8578b-147">This code also checks to see if a batch is scheduled.</span></span> <span data-ttu-id="8578b-148">Dans cet exemple, l’exécution de chaque traitement par lots est prévue toutes les 100 millisecondes.</span><span class="sxs-lookup"><span data-stu-id="8578b-148">In this example, each batch is scheduled to run every 100ms.</span></span> <span data-ttu-id="8578b-149">Vous pouvez ajuster cette valeur si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="8578b-149">You can adjust this value as needed.</span></span> <span data-ttu-id="8578b-150">Des valeurs supérieures entraînent l’envoi de traitements par lots plus grands au service à distance et l’augmentation du temps d’attente pour que l’utilisateur puisse afficher les résultats.</span><span class="sxs-lookup"><span data-stu-id="8578b-150">Higher values result in bigger batches being sent to the remote service, and a longer wait time for the user to see results.</span></span> <span data-ttu-id="8578b-151">Des valeurs inférieures ont tendance à envoyer davantage de traitements par lots au service à distance, mais avec un temps de réponse rapide pour les utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="8578b-151">Lower values tend to send more batches to the remote service, but with a quick response time for users.</span></span>

### <a name="add-the-_pushoperation-function-to-functionsts"></a><span data-ttu-id="8578b-152">Ajouter la fonction `_pushOperation` à functions.ts</span><span class="sxs-lookup"><span data-stu-id="8578b-152">Add the `_pushOperation` function to functions.ts</span></span>

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

## <a name="make-the-remote-request"></a><span data-ttu-id="8578b-153">Créer la demande à distance</span><span class="sxs-lookup"><span data-stu-id="8578b-153">Make the remote request</span></span>

<span data-ttu-id="8578b-154">L’objectif de la fonction `_makeRemoteRequest` consiste à transmettre le traitement par lots d’opérations au service à distance, puis de renvoyer les résultats à chaque fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8578b-154">The purpose of the `_makeRemoteRequest` function is to pass the batch of operations to the remote service, and then return the results to each custom function.</span></span> <span data-ttu-id="8578b-155">Elle crée tout d’abord une copie du tableau de traitement par lots.</span><span class="sxs-lookup"><span data-stu-id="8578b-155">It first creates a copy of the batch array.</span></span> <span data-ttu-id="8578b-156">Cela permet aux appels simultanés de fonctions personnalisées à partir d’Excel de commencer immédiatement le traitement par lots dans un nouveau tableau.</span><span class="sxs-lookup"><span data-stu-id="8578b-156">This allows concurrent custom function calls from Excel to immediately begin batching in a new array.</span></span> <span data-ttu-id="8578b-157">La copie est ensuite transformée en un tableau plus simple qui ne contient pas les informations sur la promesse.</span><span class="sxs-lookup"><span data-stu-id="8578b-157">The copy is then turned into a simpler array that does not contain the promise information.</span></span> <span data-ttu-id="8578b-158">Transmettre les promesses à un service à distance n’aurait aucun sens, car elles ne fonctionneraient pas.</span><span class="sxs-lookup"><span data-stu-id="8578b-158">It wouldn't make sense to pass the promises to a remote service since they would not work.</span></span> <span data-ttu-id="8578b-159">`_makeRemoteRequest` rejette ou résout chaque promesse en fonction de ce que le service à distance renvoie.</span><span class="sxs-lookup"><span data-stu-id="8578b-159">The `_makeRemoteRequest` will either reject or resolve each promise based on what the remote service returns.</span></span>

### <a name="add-the-following-_makeremoterequest-method-to-functionsts"></a><span data-ttu-id="8578b-160">Ajouter la méthode `_makeRemoteRequest` suivante à functions.ts</span><span class="sxs-lookup"><span data-stu-id="8578b-160">Add the following `_makeRemoteRequest` method to functions.ts</span></span>

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

### <a name="modify-_makeremoterequest-for-your-own-solution"></a><span data-ttu-id="8578b-161">Modifier `_makeRemoteRequest` pour votre propre solution</span><span class="sxs-lookup"><span data-stu-id="8578b-161">Modify `_makeRemoteRequest` for your own solution</span></span>

<span data-ttu-id="8578b-162">La fonction `_makeRemoteRequest` appelle `_fetchFromRemoteService` qui, comme vous le verrez plus tard, est simplement une imitation représentant le service à distance.</span><span class="sxs-lookup"><span data-stu-id="8578b-162">The `_makeRemoteRequest` function calls `_fetchFromRemoteService` which, as you'll see later, is just a mock representing the remote service.</span></span> <span data-ttu-id="8578b-163">Cela facilite l’étude et l’exécution du code dans cet article.</span><span class="sxs-lookup"><span data-stu-id="8578b-163">This makes it easier to study and run the code in this article.</span></span> <span data-ttu-id="8578b-164">Mais si vous voulez utiliser ce code pour un vrai service à distance, vous devez effectuer les modifications suivantes :</span><span class="sxs-lookup"><span data-stu-id="8578b-164">But when you want to use this code for an actual remote service you should make the following changes:</span></span>

- <span data-ttu-id="8578b-165">Déterminez la manière dont vous souhaitez sérialiser les opérations de traitement par lots sur le réseau.</span><span class="sxs-lookup"><span data-stu-id="8578b-165">Decide how to serialize the batch operations over the network.</span></span> <span data-ttu-id="8578b-166">Par exemple, vous souhaiterez peut-être placer le tableau dans un corps JSON.</span><span class="sxs-lookup"><span data-stu-id="8578b-166">For example, you may want to put the array into a JSON body.</span></span>
- <span data-ttu-id="8578b-167">Au lieu d’appeler `_fetchFromRemoteService`, vous devez passer le véritable appel réseau au service à distance en transmettant le traitement par lots des opérations.</span><span class="sxs-lookup"><span data-stu-id="8578b-167">Instead of calling `_fetchFromRemoteService` you need to make the actual network call to the remote service passing the batch of operations.</span></span>

## <a name="process-the-batch-call-on-the-remote-service"></a><span data-ttu-id="8578b-168">Traiter l’appel de traitement par lots sur le service à distance</span><span class="sxs-lookup"><span data-stu-id="8578b-168">Process the batch call on the remote service</span></span>

<span data-ttu-id="8578b-169">La dernière étape consiste à gérer l’appel de traitement par lots dans le service à distance.</span><span class="sxs-lookup"><span data-stu-id="8578b-169">The last step is to handle the batch call in the remote service.</span></span> <span data-ttu-id="8578b-170">L’exemple de code suivant affiche la fonction `_fetchFromRemoteService`.</span><span class="sxs-lookup"><span data-stu-id="8578b-170">The following code sample shows the `_fetchFromRemoteService` function.</span></span> <span data-ttu-id="8578b-171">Cette fonction décompresse chaque opération, effectue l’opération spécifiée et renvoie les résultats.</span><span class="sxs-lookup"><span data-stu-id="8578b-171">This function unpacks each operation, performs the specified operation, and returns the results.</span></span> <span data-ttu-id="8578b-172">À des fins d’apprentissage dans cet article, la fonction `_fetchFromRemoteService` est conçue de manière à s’exécuter dans votre complément web et à imiter un service à distance.</span><span class="sxs-lookup"><span data-stu-id="8578b-172">For learning purposes in this article, the `_fetchFromRemoteService` function is designed to run in your web add-in and mock a remote service.</span></span> <span data-ttu-id="8578b-173">Vous pouvez ajouter ce code à votre fichier **functions.ts** afin d’examiner et d’exécuter l’ensemble du code de cet article sans devoir configurer de service à distance réel.</span><span class="sxs-lookup"><span data-stu-id="8578b-173">You can add this code to your **functions.ts** file so that you can study and run all the code in this article without having to set up an actual remote service.</span></span>

### <a name="add-the-following-_fetchfromremoteservice-function-to-functionsts"></a><span data-ttu-id="8578b-174">Ajouter la fonction `_fetchFromRemoteService` suivante à functions.ts</span><span class="sxs-lookup"><span data-stu-id="8578b-174">Add the following `_fetchFromRemoteService` function to functions.ts</span></span>

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

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a><span data-ttu-id="8578b-175">Modifier `_fetchFromRemoteService` pour votre service à distance en direct</span><span class="sxs-lookup"><span data-stu-id="8578b-175">Modify `_fetchFromRemoteService` for your live remote service</span></span>

<span data-ttu-id="8578b-176">Pour modifier la fonction `_fetchFromRemoteService` de manière à l’exécuter dans votre service à distance en direct, apportez les modifications suivantes :</span><span class="sxs-lookup"><span data-stu-id="8578b-176">To modify the `_fetchFromRemoteService` function to run in your live remote service, make the following changes:</span></span>

- <span data-ttu-id="8578b-177">Selon votre plateforme serveur (Node.js ou autres), mappez l’appel du réseau client à cette fonction.</span><span class="sxs-lookup"><span data-stu-id="8578b-177">Depending on your server platform (Node.js or others) map the client network call to this function.</span></span>
- <span data-ttu-id="8578b-178">Supprimez la fonction `pause`, qui reproduit la latence du réseau dans le cadre de l’imitation.</span><span class="sxs-lookup"><span data-stu-id="8578b-178">Remove the `pause` function which simulates network latency as part of the mock.</span></span>
- <span data-ttu-id="8578b-179">Modifiez la déclaration de fonction de manière à ce qu’elle fonctionne avec le paramètre transmis si le paramètre est modifié à des fins de réseau.</span><span class="sxs-lookup"><span data-stu-id="8578b-179">Modify the function declaration to work with the parameter passed if the parameter is changed for network purposes.</span></span> <span data-ttu-id="8578b-180">Par exemple, au lieu d’un tableau, il peut s’agir d’un corps JSON d’opérations traitées par lots à traiter.</span><span class="sxs-lookup"><span data-stu-id="8578b-180">For example, instead of an array, it may be a JSON body of batched operations to process.</span></span>
- <span data-ttu-id="8578b-181">Modifiez la fonction de manière à effectuer les opérations (ou appelez les fonctions qui effectuent les opérations).</span><span class="sxs-lookup"><span data-stu-id="8578b-181">Modify the function to perform the operations (or call functions that do the operations).</span></span>
- <span data-ttu-id="8578b-182">Appliquez un mécanisme d’authentification approprié.</span><span class="sxs-lookup"><span data-stu-id="8578b-182">Apply an appropriate authentication mechanism.</span></span> <span data-ttu-id="8578b-183">Veillez à ce que seuls les appelants corrects puissent accéder à la fonction.</span><span class="sxs-lookup"><span data-stu-id="8578b-183">Ensure that only the correct callers can access the function.</span></span>
- <span data-ttu-id="8578b-184">Placez le code dans le service à distance.</span><span class="sxs-lookup"><span data-stu-id="8578b-184">Place the code in the remote service.</span></span>

## <a name="next-steps"></a><span data-ttu-id="8578b-185">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="8578b-185">Next steps</span></span>
<span data-ttu-id="8578b-186">Découvrez [les différents paramètres](custom-functions-parameter-options.md) que vous pouvez utiliser dans vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8578b-186">Learn about [the various parameters](custom-functions-parameter-options.md) you can use in your custom functions.</span></span> <span data-ttu-id="8578b-187">Ou parcourez les concepts de base d’un [appel web via une fonction personnalisée](custom-functions-web-reqs.md).</span><span class="sxs-lookup"><span data-stu-id="8578b-187">Or review the basics behind making [a web call through a custom function](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8578b-188">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8578b-188">See also</span></span>

* [<span data-ttu-id="8578b-189">Valeurs volatiles dans les fonctions</span><span class="sxs-lookup"><span data-stu-id="8578b-189">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="8578b-190">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="8578b-190">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="8578b-191">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="8578b-191">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
