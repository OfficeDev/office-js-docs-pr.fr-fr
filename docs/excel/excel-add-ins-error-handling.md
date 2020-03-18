---
title: Gestion des erreurs
description: En savoir plus sur la logique de gestion des erreurs de l’API JavaScript Excel pour prendre en compte les erreurs d’exécution.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bee5824d8854a55d5ac4041be1335ce239b31a9e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717165"
---
# <a name="error-handling"></a><span data-ttu-id="1fbb7-103">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="1fbb7-103">Error handling</span></span>

<span data-ttu-id="1fbb7-p101">Lorsque vous créez un complément à l’aide de l’API JavaScript pour Excel, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. Il s’agit d’une étape essentielle en raison de la nature asynchrone de l’API.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="1fbb7-106">Pour plus d’informations sur `sync()` la méthode et la nature asynchrone de l’API JavaScript pour Excel, voir [concepts de programmation fondamentaux avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="1fbb7-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="1fbb7-107">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="1fbb7-107">Best practices</span></span>

<span data-ttu-id="1fbb7-p102">Dans l’ensemble des exemples de code de cette documentation, vous remarquerez que chaque appel à `Excel.run` est accompagné d’une instruction `catch` afin de détecter les erreurs qui se produisent au sein de `Excel.run`. Nous vous recommandons d’utiliser le même modèle lorsque vous développez un complément à l’aide des API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="1fbb7-110">Erreurs API</span><span class="sxs-lookup"><span data-stu-id="1fbb7-110">API errors</span></span>

<span data-ttu-id="1fbb7-111">Quand une demande d’API JavaScript pour Excel ne parvient pas à s’exécuter correctement, l’API renvoie un objet d’erreur qui contient les propriétés suivantes :</span><span class="sxs-lookup"><span data-stu-id="1fbb7-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="1fbb7-112">**code** :  la propriété `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-112">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="1fbb7-113">Par exemple, le code d’erreur « InvalidReference » indique que la référence n’est pas valide pour l’opération spécifiée.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-113">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="1fbb7-114">Les codes d’erreur ne sont pas traduits.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-114">Error codes are not localized.</span></span>

- <span data-ttu-id="1fbb7-115">**message** : la propriété `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="1fbb7-116">Le message d’erreur n’est pas conçu pour être utilisé par l’utilisateur final. Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finals.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="1fbb7-117">**debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause principale de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="1fbb7-118">Si vous utilisez `console.log()` pour imprimer les messages d’erreur de la console, ces messages ne seront visibles que sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="1fbb7-119">Les utilisateurs finals ne verront pas ces messages d’erreur dans le volet Office du complément ni n’importe où dans l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-119">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="1fbb7-120">Messages d’erreur</span><span class="sxs-lookup"><span data-stu-id="1fbb7-120">Error Messages</span></span>

<span data-ttu-id="1fbb7-121">Le tableau suivant contient la liste des erreurs que l’API peut renvoyer.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="1fbb7-122">error.code</span><span class="sxs-lookup"><span data-stu-id="1fbb7-122">error.code</span></span> | <span data-ttu-id="1fbb7-123">error.message</span><span class="sxs-lookup"><span data-stu-id="1fbb7-123">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="1fbb7-124">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="1fbb7-124">InvalidArgument</span></span> |<span data-ttu-id="1fbb7-125">L’argument est manquant ou non valide, ou a un format incorrect.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-125">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="1fbb7-126">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="1fbb7-126">InvalidRequest</span></span>  |<span data-ttu-id="1fbb7-127">Impossible de traiter la demande.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-127">Cannot process the request.</span></span>|
|<span data-ttu-id="1fbb7-128">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="1fbb7-128">InvalidReference</span></span>|<span data-ttu-id="1fbb7-129">Cette référence n’est pas valide pour l’opération en cours.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-129">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="1fbb7-130">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="1fbb7-130">InvalidBinding</span></span>  |<span data-ttu-id="1fbb7-131">Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-131">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="1fbb7-132">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="1fbb7-132">InvalidSelection</span></span>|<span data-ttu-id="1fbb7-133">La sélection en cours est incorrecte pour cette action.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-133">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="1fbb7-134">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="1fbb7-134">Unauthenticated</span></span> |<span data-ttu-id="1fbb7-135">Les informations d’authentification requises sont manquantes ou incorrectes.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-135">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="1fbb7-136">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="1fbb7-136">AccessDenied</span></span> |<span data-ttu-id="1fbb7-137">Vous ne pouvez pas effectuer l’opération demandée.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-137">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="1fbb7-138">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="1fbb7-138">ItemNotFound</span></span> |<span data-ttu-id="1fbb7-139">La ressource demandée n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-139">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="1fbb7-140">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="1fbb7-140">ActivityLimitReached</span></span>|<span data-ttu-id="1fbb7-141">La limite d’activité a été atteinte.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-141">Activity limit has been reached.</span></span>|
|<span data-ttu-id="1fbb7-142">GeneralException</span><span class="sxs-lookup"><span data-stu-id="1fbb7-142">GeneralException</span></span>|<span data-ttu-id="1fbb7-143">Une erreur interne s’est produite lors du traitement de la demande.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-143">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="1fbb7-144">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="1fbb7-144">NotImplemented</span></span>  |<span data-ttu-id="1fbb7-145">La fonctionnalité demandée n’est pas implémentée</span><span class="sxs-lookup"><span data-stu-id="1fbb7-145">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="1fbb7-146">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="1fbb7-146">ServiceNotAvailable</span></span>|<span data-ttu-id="1fbb7-147">Le service n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-147">The service is unavailable.</span></span>|
|<span data-ttu-id="1fbb7-148">Conflict</span><span class="sxs-lookup"><span data-stu-id="1fbb7-148">Conflict</span></span>|<span data-ttu-id="1fbb7-149">La demande n’a pas pu être traitée en raison d’un conflit.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-149">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="1fbb7-150">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="1fbb7-150">ItemAlreadyExists</span></span>|<span data-ttu-id="1fbb7-151">La ressource en cours de création existe déjà.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-151">The resource being created already exists.</span></span>|
|<span data-ttu-id="1fbb7-152">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="1fbb7-152">UnsupportedOperation</span></span>|<span data-ttu-id="1fbb7-153">L’opération tentée n’est pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-153">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="1fbb7-154">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="1fbb7-154">RequestAborted</span></span>|<span data-ttu-id="1fbb7-155">La demande a été interrompue pendant l’exécution.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-155">The request was aborted during run time.</span></span>|
|<span data-ttu-id="1fbb7-156">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="1fbb7-156">ApiNotAvailable</span></span>|<span data-ttu-id="1fbb7-157">L’API demandée n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-157">The requested API is not available.</span></span>|
|<span data-ttu-id="1fbb7-158">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="1fbb7-158">InsertDeleteConflict</span></span>|<span data-ttu-id="1fbb7-159">L’opération d’insertion ou de suppression tentée a créé un conflit.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-159">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="1fbb7-160">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="1fbb7-160">InvalidOperation</span></span>|<span data-ttu-id="1fbb7-161">L’opération tentée n’est pas valide sur l’objet.</span><span class="sxs-lookup"><span data-stu-id="1fbb7-161">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="1fbb7-162">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1fbb7-162">See also</span></span>

- [<span data-ttu-id="1fbb7-163">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="1fbb7-163">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="1fbb7-164">Objet OfficeExtension.Error (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="1fbb7-164">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error)
