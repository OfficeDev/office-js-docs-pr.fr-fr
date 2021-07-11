---
ms.date: 09/25/2020
description: Comprendre Excel fonctions personnalisées qui n’utilisent pas de volet de tâches et leur runtime JavaScript spécifique.
title: Runtime pour les fonctions personnalisées sans interface Excel’interface utilisateur
localization_priority: Normal
ms.openlocfilehash: aa2cf2632ddf9eb1ad1eb202b031ee2ca686af01
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349622"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a><span data-ttu-id="9d343-103">Runtime pour les fonctions personnalisées sans interface Excel’interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="9d343-103">Runtime for UI-less Excel custom functions</span></span>

<span data-ttu-id="9d343-104">Les fonctions personnalisées qui n’utilisent pas de volet de tâches (fonctions personnalisées sans interface utilisateur) utilisent un runtime JavaScript conçu pour optimiser les performances des calculs.</span><span class="sxs-lookup"><span data-stu-id="9d343-104">Custom functions that don't use a task pane (UI-less custom functions) use a JavaScript runtime that is designed to optimize performance of calculations.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="9d343-105">Ce runtime JavaScript permet d’accéder aux API de l’espace de noms qui peuvent être utilisées par les fonctions personnalisées sans interface utilisateur et le volet Des tâches pour stocker `OfficeRuntime` des données.</span><span class="sxs-lookup"><span data-stu-id="9d343-105">This JavaScript runtime provides access to APIs in the `OfficeRuntime` namespace that can be used by UI-less custom functions and the task pane to store data.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="9d343-106">Demande de données externes</span><span class="sxs-lookup"><span data-stu-id="9d343-106">Requesting external data</span></span>

<span data-ttu-id="9d343-107">Dans une fonction personnalisée sans interface utilisateur, vous pouvez demander des données externes à l’aide d’une API telle que [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) ou en utilisant [XmlHttpRequest (XHR),](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)une API web standard qui émettre des demandes HTTP pour interagir avec les serveurs.</span><span class="sxs-lookup"><span data-stu-id="9d343-107">Within a UI-less custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="9d343-108">N’ignorez pas que les fonctions sans interface utilisateur doivent utiliser des mesures [](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) de sécurité supplémentaires lors de la génération de XmlHttpRequests, nécessitant une stratégie d’origine identique et [un CORS](https://www.w3.org/TR/cors/)simple.</span><span class="sxs-lookup"><span data-stu-id="9d343-108">Be aware that UI-less functions must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="9d343-109">Une implémentation CORS simple ne peut pas utiliser de cookies et prend uniquement en charge les méthodes simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="9d343-109">A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="9d343-110">Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="9d343-110">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="9d343-111">Vous pouvez également utiliser un `Content-Type` en-tête dans CORS simple, à condition que le type de contenu `application/x-www-form-urlencoded` soit , ou `text/plain` `multipart/form-data` .</span><span class="sxs-lookup"><span data-stu-id="9d343-111">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

## <a name="storing-and-accessing-data"></a><span data-ttu-id="9d343-112">Accès aux données et stockage</span><span class="sxs-lookup"><span data-stu-id="9d343-112">Storing and accessing data</span></span>

<span data-ttu-id="9d343-113">Dans une fonction personnalisée sans interface utilisateur, vous pouvez stocker et accéder aux données à l’aide de `OfficeRuntime.storage` l’objet.</span><span class="sxs-lookup"><span data-stu-id="9d343-113">Within a UI-less custom function, you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="9d343-114">`Storage` est un système de stockage persistant, non chiffré et à valeur clé qui fournit une alternative à [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé par des fonctions personnalisées sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9d343-114">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), which cannot be used by UI-less custom functions.</span></span> <span data-ttu-id="9d343-115">`Storage` offre 10 Mo de données par domaine.</span><span class="sxs-lookup"><span data-stu-id="9d343-115">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="9d343-116">Les domaines peuvent être partagés par plusieurs modules.</span><span class="sxs-lookup"><span data-stu-id="9d343-116">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="9d343-117">`Storage` est conçu comme une solution de stockage partagé, ce qui signifie que plusieurs parties d’un complément ont accès aux mêmes données.</span><span class="sxs-lookup"><span data-stu-id="9d343-117">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="9d343-118">Par exemple, les jetons pour l’authentification des utilisateurs peuvent être stockés, car ils sont accessibles à la fois par une fonction personnalisée sans interface utilisateur et par des éléments d’interface utilisateur de add-in tels qu’un volet Des `storage` tâches.</span><span class="sxs-lookup"><span data-stu-id="9d343-118">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a UI-less custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="9d343-119">De même, si deux modules complémentaires partagent le même domaine (par exemple, , ), ils sont également autorisés à partager des informations entre `www.contoso.com/addin1` `www.contoso.com/addin2` `storage` eux.</span><span class="sxs-lookup"><span data-stu-id="9d343-119">Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="9d343-120">Notez que les add-ins qui ont différents sous-domaine auront différentes instances `storage` de (par exemple, `subdomain.contoso.com/addin1` , `differentsubdomain.contoso.com/addin2` ).</span><span class="sxs-lookup"><span data-stu-id="9d343-120">Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span></span>

<span data-ttu-id="9d343-121">Comme `storage` peut être un emplacement partagé, il est important de savoir qu’il est possible de remplacer des paires clé-valeur.</span><span class="sxs-lookup"><span data-stu-id="9d343-121">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="9d343-122">Les méthodes suivantes sont disponibles sur `storage` l’objet.</span><span class="sxs-lookup"><span data-stu-id="9d343-122">The following methods are available on the `storage` object.</span></span>

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> <span data-ttu-id="9d343-123">Il n’existe aucune méthode pour effacer toutes les informations (par `clear` exemple).</span><span class="sxs-lookup"><span data-stu-id="9d343-123">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="9d343-124">À la place, vous devez utiliser l’objet `removeItems` pour supprimer plusieurs entrées à la fois.</span><span class="sxs-lookup"><span data-stu-id="9d343-124">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="9d343-125">Exemple OfficeRuntime.storage</span><span class="sxs-lookup"><span data-stu-id="9d343-125">OfficeRuntime.storage example</span></span>

<span data-ttu-id="9d343-126">L’exemple de code suivant appelle `OfficeRuntime.storage.setItem` la fonction pour définir une clé et une valeur dans `storage` .</span><span class="sxs-lookup"><span data-stu-id="9d343-126">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="9d343-127">Considérations supplémentaires</span><span class="sxs-lookup"><span data-stu-id="9d343-127">Additional considerations</span></span>

<span data-ttu-id="9d343-128">Si votre add-in utilise uniquement des fonctions personnalisées sans interface utilisateur, notez que vous ne pouvez pas accéder au modèle DOM (Document Object Model) avec des fonctions personnalisées sans interface utilisateur ou utiliser des bibliothèques telles que jQuery qui reposent sur le DOM.</span><span class="sxs-lookup"><span data-stu-id="9d343-128">If your add-in only uses UI-less custom functions, note that you can't access the Document Object Model (DOM) with UI-less custom functions or use libraries like jQuery that rely on the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="9d343-129">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="9d343-129">Next steps</span></span>
<span data-ttu-id="9d343-130">Découvrez comment [déboguer des](custom-functions-debugging.md)fonctions personnalisées sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9d343-130">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="9d343-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9d343-131">See also</span></span>

* [<span data-ttu-id="9d343-132">Authentifier les fonctions personnalisées sans interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="9d343-132">Authenticate UI-less custom functions</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="9d343-133">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="9d343-133">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="9d343-134">Didacticiel sur les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="9d343-134">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
