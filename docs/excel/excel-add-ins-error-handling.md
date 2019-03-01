---
title: Gestion des erreurs
description: ''
ms.date: 10/16/2018
localization_priority: Normal
ms.openlocfilehash: 8c6de5d2a22fdb4614742ddfb7fbf566780c0f0f
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/28/2019
ms.locfileid: "29388961"
---
# <a name="error-handling"></a><span data-ttu-id="c79a3-102">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="c79a3-102">Error handling</span></span>

<span data-ttu-id="c79a3-103">Lorsque vous créez un complément à l’aide de l’API JavaScript pour Excel, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="c79a3-103">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="c79a3-104">Il s’agit d’une étape essentielle en raison de la nature asynchrone de l’API.</span><span class="sxs-lookup"><span data-stu-id="c79a3-104">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="c79a3-105">Pour plus d’informations sur la méthode **sync()** et la nature asynchrone de l’API JavaScript pour Excel, reportez-vous à la rubrique [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="c79a3-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="c79a3-106">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="c79a3-106">Best practices</span></span>

<span data-ttu-id="c79a3-107">Dans l’ensemble des exemples de code de cette documentation, vous remarquerez que chaque appel à `Excel.run` est accompagné d’une instruction `catch` afin de détecter les erreurs qui se produisent au sein de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="c79a3-107">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="c79a3-108">Nous vous recommandons d’utiliser le même modèle lorsque vous développez un complément à l’aide des API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="c79a3-108">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="c79a3-109">Erreurs API</span><span class="sxs-lookup"><span data-stu-id="c79a3-109">API errors</span></span>

<span data-ttu-id="c79a3-110">Quand une demande d’API JavaScript pour Excel ne parvient pas à s’exécuter correctement, l’API renvoie un objet d’erreur qui contient les propriétés suivantes :</span><span class="sxs-lookup"><span data-stu-id="c79a3-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="c79a3-111">**code** :  la propriété `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`.</span><span class="sxs-lookup"><span data-stu-id="c79a3-111">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="c79a3-112">Par exemple, le code d’erreur « InvalidReference » indique que la référence n’est pas valide pour l’opération spécifiée.</span><span class="sxs-lookup"><span data-stu-id="c79a3-112">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="c79a3-113">Les codes d’erreur ne sont pas traduits.</span><span class="sxs-lookup"><span data-stu-id="c79a3-113">Error codes are not localized.</span></span>

- <span data-ttu-id="c79a3-114">**message** : la propriété `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée.</span><span class="sxs-lookup"><span data-stu-id="c79a3-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="c79a3-115">Le message d’erreur n’est pas conçu pour être utilisé par l’utilisateur final. Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finals.</span><span class="sxs-lookup"><span data-stu-id="c79a3-115">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="c79a3-116">**debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause principale de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="c79a3-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="c79a3-117">Si vous utilisez `console.log()` pour imprimer les messages d’erreur de la console, ces messages ne seront visibles que sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="c79a3-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="c79a3-118">Les utilisateurs finals ne verront pas ces messages d’erreur dans le volet Office du complément ni n’importe où dans l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="c79a3-118">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="c79a3-119">Messages d’erreur</span><span class="sxs-lookup"><span data-stu-id="c79a3-119">Error Messages</span></span>

<span data-ttu-id="c79a3-120">Le tableau suivant contient la liste des erreurs que l’API peut renvoyer.</span><span class="sxs-lookup"><span data-stu-id="c79a3-120">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="c79a3-121">error.code</span><span class="sxs-lookup"><span data-stu-id="c79a3-121">error.code</span></span> | <span data-ttu-id="c79a3-122">error.message</span><span class="sxs-lookup"><span data-stu-id="c79a3-122">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="c79a3-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="c79a3-123">InvalidArgument</span></span> |<span data-ttu-id="c79a3-124">L’argument est manquant ou non valide, ou a un format incorrect.</span><span class="sxs-lookup"><span data-stu-id="c79a3-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="c79a3-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="c79a3-125">InvalidRequest</span></span>  |<span data-ttu-id="c79a3-126">Impossible de traiter la demande.</span><span class="sxs-lookup"><span data-stu-id="c79a3-126">Cannot process the request.</span></span>|
|<span data-ttu-id="c79a3-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="c79a3-127">InvalidReference</span></span>|<span data-ttu-id="c79a3-128">Cette référence n’est pas valide pour l’opération en cours.</span><span class="sxs-lookup"><span data-stu-id="c79a3-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="c79a3-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="c79a3-129">InvalidBinding</span></span>  |<span data-ttu-id="c79a3-130">Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.</span><span class="sxs-lookup"><span data-stu-id="c79a3-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="c79a3-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c79a3-131">InvalidSelection</span></span>|<span data-ttu-id="c79a3-132">La sélection en cours est incorrecte pour cette action.</span><span class="sxs-lookup"><span data-stu-id="c79a3-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="c79a3-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="c79a3-133">Unauthenticated</span></span> |<span data-ttu-id="c79a3-134">Les informations d’authentification requises sont manquantes ou incorrectes.</span><span class="sxs-lookup"><span data-stu-id="c79a3-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="c79a3-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="c79a3-135">AccessDenied</span></span> |<span data-ttu-id="c79a3-136">Vous ne pouvez pas effectuer l’opération demandée.</span><span class="sxs-lookup"><span data-stu-id="c79a3-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="c79a3-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="c79a3-137">ItemNotFound</span></span> |<span data-ttu-id="c79a3-138">La ressource demandée n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="c79a3-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="c79a3-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="c79a3-139">ActivityLimitReached</span></span>|<span data-ttu-id="c79a3-140">La limite d’activité a été atteinte.</span><span class="sxs-lookup"><span data-stu-id="c79a3-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="c79a3-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c79a3-141">GeneralException</span></span>|<span data-ttu-id="c79a3-142">Une erreur interne s’est produite lors du traitement de la demande.</span><span class="sxs-lookup"><span data-stu-id="c79a3-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="c79a3-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="c79a3-143">NotImplemented</span></span>  |<span data-ttu-id="c79a3-144">La fonctionnalité demandée n’est pas implémentée</span><span class="sxs-lookup"><span data-stu-id="c79a3-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="c79a3-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="c79a3-145">ServiceNotAvailable</span></span>|<span data-ttu-id="c79a3-146">Le service n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="c79a3-146">The service is unavailable.</span></span>|
|<span data-ttu-id="c79a3-147">Conflict</span><span class="sxs-lookup"><span data-stu-id="c79a3-147">Conflict</span></span>|<span data-ttu-id="c79a3-148">La demande n’a pas pu être traitée en raison d’un conflit.</span><span class="sxs-lookup"><span data-stu-id="c79a3-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="c79a3-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="c79a3-149">ItemAlreadyExists</span></span>|<span data-ttu-id="c79a3-150">La ressource en cours de création existe déjà.</span><span class="sxs-lookup"><span data-stu-id="c79a3-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="c79a3-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="c79a3-151">UnsupportedOperation</span></span>|<span data-ttu-id="c79a3-152">L’opération tentée n’est pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="c79a3-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="c79a3-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="c79a3-153">RequestAborted</span></span>|<span data-ttu-id="c79a3-154">La demande a été interrompue pendant l’exécution.</span><span class="sxs-lookup"><span data-stu-id="c79a3-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="c79a3-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="c79a3-155">ApiNotAvailable</span></span>|<span data-ttu-id="c79a3-156">L’API demandée n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="c79a3-156">The requested API is not available.</span></span>|
|<span data-ttu-id="c79a3-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="c79a3-157">InsertDeleteConflict</span></span>|<span data-ttu-id="c79a3-158">L’opération d’insertion ou de suppression tentée a créé un conflit.</span><span class="sxs-lookup"><span data-stu-id="c79a3-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="c79a3-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="c79a3-159">InvalidOperation</span></span>|<span data-ttu-id="c79a3-160">L’opération tentée n’est pas valide sur l’objet.</span><span class="sxs-lookup"><span data-stu-id="c79a3-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="c79a3-161">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c79a3-161">See also</span></span>

- [<span data-ttu-id="c79a3-162">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c79a3-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c79a3-163">Objet OfficeExtension.Error (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="c79a3-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error)
