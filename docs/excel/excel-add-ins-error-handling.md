---
title: Gestion des erreurs
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 23a70b1d66befb971c3c1394eb9162c19f2ee176
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348085"
---
# <a name="error-handling"></a><span data-ttu-id="f931e-102">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="f931e-102">Error handling</span></span>

<span data-ttu-id="f931e-103">Lorsque vous créez un complément à l’aide de l’API JavaScript pour Excel, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="f931e-103">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="f931e-104">Il s’agit d’une étape essentielle en raison de la nature asynchrone de l’API.</span><span class="sxs-lookup"><span data-stu-id="f931e-104">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="f931e-105">Pour plus d’informations sur la méthode **sync()** et la nature asynchrone de l’API JavaScript pour Excel, reportez-vous à la rubrique [Concepts de base de l’API JavaScript pour Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="f931e-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="f931e-106">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="f931e-106">Best practices</span></span>

<span data-ttu-id="f931e-107">Dans l’ensemble des exemples de code de cette documentation, vous remarquerez que chaque appel à `Excel.run` est accompagné d’une instruction `catch` afin de détecter les erreurs qui se produisent au sein de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="f931e-107">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="f931e-108">Nous vous recommandons d’utiliser le même modèle lorsque vous développez un complément à l’aide des API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="f931e-108">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="f931e-109">Erreurs API</span><span class="sxs-lookup"><span data-stu-id="f931e-109">API errors</span></span> 

<span data-ttu-id="f931e-110">Quand une demande d’API JavaScript pour Excel ne parvient pas à s’exécuter correctement, l’API renvoie un objet d’erreur qui contient les propriétés suivantes :</span><span class="sxs-lookup"><span data-stu-id="f931e-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="f931e-111">**code** :  la propriété `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`.</span><span class="sxs-lookup"><span data-stu-id="f931e-111">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="f931e-112">Par exemple, le code d’erreur « InvalidReference » indique que la référence n’est pas valide pour l’opération spécifiée.</span><span class="sxs-lookup"><span data-stu-id="f931e-112">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="f931e-113">Les codes d’erreur ne sont pas traduits.</span><span class="sxs-lookup"><span data-stu-id="f931e-113">Error codes are not localized.</span></span> 

- <span data-ttu-id="f931e-114">**message** : la propriété `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée.</span><span class="sxs-lookup"><span data-stu-id="f931e-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="f931e-115">Le message d’erreur n’est pas conçu pour être utilisé par l’utilisateur final. Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finals.</span><span class="sxs-lookup"><span data-stu-id="f931e-115">The error message is not intended for end-user consumption; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end-users.</span></span>

- <span data-ttu-id="f931e-116">**debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause principale de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="f931e-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="f931e-117">Si vous utilisez `console.log()` pour imprimer les messages d’erreur de la console, ces messages ne seront visibles que sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="f931e-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="f931e-118">Les utilisateurs finaux ne verront pas ces messages d’erreur dans le volet Office du complément ou n’importe où dans l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="f931e-118">End-users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="see-also"></a><span data-ttu-id="f931e-119">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f931e-119">See also</span></span>

- [<span data-ttu-id="f931e-120">Concepts de base de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="f931e-120">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f931e-121">Objet OfficeExtension.Error (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="f931e-121">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
