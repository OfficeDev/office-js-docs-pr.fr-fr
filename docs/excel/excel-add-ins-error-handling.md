---
title: Gestion des erreurs
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b07012516cbe15374d0707c157738117a9c8fe96
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459230"
---
# <a name="error-handling"></a><span data-ttu-id="9345f-102">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="9345f-102">Error handling</span></span>

<span data-ttu-id="9345f-p101">Lorsque vous créez un complément à l'aide de l'API JavaScript Excel, veillez à inclure une logique de traitement des erreurs afin de prendre en compte les erreurs d'exécution. Cela est essentiel en raison de la nature asynchrone de l'API.</span><span class="sxs-lookup"><span data-stu-id="9345f-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="9345f-105">Pour plus d’informations sur la méthode **sync()** et la nature asynchrone de l’API JavaScript d'Excel, voir [Concepts de programmation fondamentaux avec l’API JavaScript d'Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="9345f-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="9345f-106">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="9345f-106">Best practices</span></span>

<span data-ttu-id="9345f-p102">Tout au long des exemples de code de cette documentation, vous remarquerez que chaque appel à `Excel.run` est accompagné d’une `catch` instruction permettant de détecter les erreurs éventuelles dans le fichier `Excel.run`. Nous vous recommandons d'utiliser le même modèle lorsque vous créez un complément à l'aide des API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="9345f-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

```js
Excel.run(function (context) { 
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);     
```

## <a name="api-errors"></a><span data-ttu-id="9345f-109">Erreurs API</span><span class="sxs-lookup"><span data-stu-id="9345f-109">API errors</span></span> 

<span data-ttu-id="9345f-110">Quand une demande d’API JavaScript pour Excel ne parvient pas à s’exécuter correctement, l’API renvoie un objet d’erreur qui contient les propriétés suivantes :</span><span class="sxs-lookup"><span data-stu-id="9345f-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="9345f-p103">**code**: la propriété  `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Par exemple, le code d’erreur «InvalidReference» indique que la référence n’est pas valide pour l’opération spécifiée. Les codes d’erreur ne sont pas localisés.</span><span class="sxs-lookup"><span data-stu-id="9345f-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span> 

- <span data-ttu-id="9345f-p104">**message**: la propriété  `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée. Le message d’erreur n’est pas destiné à la consommation par les utilisateurs finaux ; Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="9345f-p104">**message**: The `message` property of an error message contains a summary of the error in the localized string. The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="9345f-116">**debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause première de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="9345f-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="9345f-p105">Si vous utilisez `console.log()` pour imprimer des messages d’erreur sur la console, ces messages ne seront visibles que sur le serveur. Les utilisateurs finaux ne verront pas ces messages d’erreur dans le panneau de tâches du complément ou ailleurs dans l'application hôte.</span><span class="sxs-lookup"><span data-stu-id="9345f-p105">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server. End users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="see-also"></a><span data-ttu-id="9345f-119">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9345f-119">See also</span></span>

- [<span data-ttu-id="9345f-120">Concepts  de programmation fondamentaux avec l’API JavaScript d'Excel</span><span class="sxs-lookup"><span data-stu-id="9345f-120">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="9345f-121">Objet OfficeExtension.Error (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="9345f-121">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
