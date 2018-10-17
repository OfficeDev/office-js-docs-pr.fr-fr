---
title: Gestion des erreurs
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: caba29f7d6949cc6d9df1498ac0a3d4f5de6c4ee
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579813"
---
# <a name="error-handling"></a><span data-ttu-id="9bdfa-102">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="9bdfa-102">Error handling</span></span>

<span data-ttu-id="9bdfa-p101">Lorsque vous créez un complément à l'aide de l'API JavaScript Excel, veillez à inclure une logique de traitement des erreurs afin de prendre en compte les erreurs d'exécution. Cela est essentiel en raison de la nature asynchrone de l'API.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="9bdfa-105">Pour plus d’informations sur la méthode **sync()** et la nature asynchrone de l’API JavaScript d'Excel, voir [Concepts de programmation fondamentaux avec l’API JavaScript d'Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="9bdfa-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="9bdfa-106">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="9bdfa-106">Best practices</span></span>

<span data-ttu-id="9bdfa-p102">Tout au long des exemples de code de cette documentation, vous remarquerez que chaque appel à `Excel.run` est accompagné d’une `catch` instruction permettant de détecter les erreurs éventuelles dans le fichier `Excel.run`. Nous vous recommandons d'utiliser le même modèle lorsque vous créez un complément à l'aide des API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="9bdfa-109">Erreurs API</span><span class="sxs-lookup"><span data-stu-id="9bdfa-109">API errors</span></span> 

<span data-ttu-id="9bdfa-110">Quand une demande d’API JavaScript pour Excel ne parvient pas à s’exécuter correctement, l’API renvoie un objet d’erreur qui contient les propriétés suivantes :</span><span class="sxs-lookup"><span data-stu-id="9bdfa-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="9bdfa-p103">**code**: la propriété  `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Par exemple, le code d’erreur «InvalidReference» indique que la référence n’est pas valide pour l’opération spécifiée. Les codes d’erreur ne sont pas localisés.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span> 

- <span data-ttu-id="9bdfa-p104">**message**: la propriété  `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée. Le message d’erreur n’est pas destiné à la consommation par les utilisateurs finaux ; Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-p104">**message**: The `message` property of an error message contains a summary of the error in the localized string. The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="9bdfa-116">**debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause première de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="9bdfa-p105">Si vous utilisez `console.log()` pour imprimer des messages d’erreur sur la console, ces messages ne seront visibles que sur le serveur. Les utilisateurs finaux ne verront pas ces messages d’erreur dans le panneau de tâches du complément ou ailleurs dans l'application hôte.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-p105">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server. End users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="9bdfa-119">Messages d’erreur</span><span class="sxs-lookup"><span data-stu-id="9bdfa-119">Error Messages</span></span>

<span data-ttu-id="9bdfa-120">Le tableau suivant est une liste des erreurs que l’API peut renvoyer.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-120">The following table defines a list of errors that the API may return.</span></span>

|<span data-ttu-id="9bdfa-121">error.code</span><span class="sxs-lookup"><span data-stu-id="9bdfa-121">error.code</span></span> | <span data-ttu-id="9bdfa-122">error.message</span><span class="sxs-lookup"><span data-stu-id="9bdfa-122">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="9bdfa-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="9bdfa-123">InvalidArgument</span></span> |<span data-ttu-id="9bdfa-124">L’argument est manquant ou non valide, ou a un format incorrect.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="9bdfa-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="9bdfa-125">InvalidRequest</span></span>  |<span data-ttu-id="9bdfa-126">Impossible de traiter la demande.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-126">Cannot process the request.</span></span>|
|<span data-ttu-id="9bdfa-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="9bdfa-127">InvalidReference</span></span>|<span data-ttu-id="9bdfa-128">Cette référence n’est pas valide pour l’opération en cours.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="9bdfa-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="9bdfa-129">InvalidBinding</span></span>  |<span data-ttu-id="9bdfa-130">Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="9bdfa-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="9bdfa-131">InvalidSelection</span></span>|<span data-ttu-id="9bdfa-132">La sélection en cours est incorrecte pour cette action.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="9bdfa-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="9bdfa-133">Unauthenticated</span></span> |<span data-ttu-id="9bdfa-134">Les informations d’authentification requises sont manquantes ou incorrectes.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="9bdfa-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="9bdfa-135">AccessDenied</span></span> |<span data-ttu-id="9bdfa-136">Vous ne pouvez pas effectuer l’opération demandée.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="9bdfa-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="9bdfa-137">ItemNotFound</span></span> |<span data-ttu-id="9bdfa-138">La ressource demandée n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="9bdfa-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="9bdfa-139">ActivityLimitReached</span></span>|<span data-ttu-id="9bdfa-140">La limite d’activité a été atteinte.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="9bdfa-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="9bdfa-141">GeneralException</span></span>|<span data-ttu-id="9bdfa-142">Une erreur interne s’est produite lors du traitement de la demande.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="9bdfa-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="9bdfa-143">NotImplemented</span></span>  |<span data-ttu-id="9bdfa-144">La fonctionnalité demandée n’est pas implémentée</span><span class="sxs-lookup"><span data-stu-id="9bdfa-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="9bdfa-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="9bdfa-145">ServiceNotAvailable</span></span>|<span data-ttu-id="9bdfa-146">Le service n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-146">The service is unavailable.</span></span>|
|<span data-ttu-id="9bdfa-147">Conflict</span><span class="sxs-lookup"><span data-stu-id="9bdfa-147">Conflict</span></span>|<span data-ttu-id="9bdfa-148">La demande n’a pas pu être traitée en raison d’un conflit.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="9bdfa-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="9bdfa-149">ItemAlreadyExists</span></span>|<span data-ttu-id="9bdfa-150">La ressource en cours de création existe déjà.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="9bdfa-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="9bdfa-151">UnsupportedOperation</span></span>|<span data-ttu-id="9bdfa-152">L’opération tentée n’est pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="9bdfa-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="9bdfa-153">RequestAborted</span></span>|<span data-ttu-id="9bdfa-154">La demande a été interrompue pendant l’exécution.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="9bdfa-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="9bdfa-155">ApiNotAvailable</span></span>|<span data-ttu-id="9bdfa-156">L’API demandée n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-156">The requested API is not available.</span></span>|
|<span data-ttu-id="9bdfa-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="9bdfa-157">InsertDeleteConflict</span></span>|<span data-ttu-id="9bdfa-158">L’opération d’insertion ou de suppression tentée a créé un conflit.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="9bdfa-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="9bdfa-159">InvalidOperation</span></span>|<span data-ttu-id="9bdfa-160">L’opération tentée n’est pas valide sur l’objet.</span><span class="sxs-lookup"><span data-stu-id="9bdfa-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="9bdfa-161">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9bdfa-161">See also</span></span>

- [<span data-ttu-id="9bdfa-162">Concepts  de programmation fondamentaux avec l’API JavaScript d'Excel</span><span class="sxs-lookup"><span data-stu-id="9bdfa-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="9bdfa-163">Objet OfficeExtension.Error (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="9bdfa-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
