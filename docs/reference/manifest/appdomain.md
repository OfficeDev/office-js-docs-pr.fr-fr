---
title: Élément AppDomain dans le fichier manifeste
description: Spécifie les domaines supplémentaires utilisés par votre complément et doit être approuvé par Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778647"
---
# <a name="appdomain-element"></a><span data-ttu-id="ed20e-103">AppDomain, élément</span><span class="sxs-lookup"><span data-stu-id="ed20e-103">AppDomain element</span></span>

<span data-ttu-id="ed20e-104">Spécifie un domaine supplémentaire qu’Office doit approuver, en plus de celui spécifié dans l' [élément SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="ed20e-104">Specifies an additional domain that Office should trust, in addition to the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="ed20e-105">La spécification d’un domaine a les effets suivants :</span><span class="sxs-lookup"><span data-stu-id="ed20e-105">Specifying a domain has these effects:</span></span>

- <span data-ttu-id="ed20e-106">Elle permet l’ouverture directe des pages, des itinéraires ou d’autres ressources dans le domaine dans le volet Office racine du complément sur les plateformes de bureau.</span><span class="sxs-lookup"><span data-stu-id="ed20e-106">It enables pages, routes, or other resources in the domain to be opened directly in the root task pane of the add-in on desktop Office platforms.</span></span> <span data-ttu-id="ed20e-107">(Il n’est pas nécessaire de spécifier un domaine dans un **AppDomain** pour Office sur le Web ou d’ouvrir une ressource dans un IFRAME, et il n’est pas nécessaire d’ouvrir une ressource dans une boîte de dialogue ouverte avec l' [API Dialog](../../develop/dialog-api-in-office-add-ins.md).)</span><span class="sxs-lookup"><span data-stu-id="ed20e-107">(Specifying a domain in an **AppDomain** isn't necessary for Office on the web or to open a resource in an IFrame, nor it is necessary for opening a resource in a dialog opened with the [Dialog API](../../develop/dialog-api-in-office-add-ins.md).)</span></span>
- <span data-ttu-id="ed20e-108">Elle permet aux pages du domaine d’effectuer des appels d’API Office.js à partir d’IFrames dans le complément.</span><span class="sxs-lookup"><span data-stu-id="ed20e-108">It enables pages in the domain to make Office.js API calls from IFrames within the add-in.</span></span>

<span data-ttu-id="ed20e-109">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="ed20e-109">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ed20e-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ed20e-110">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="ed20e-111">La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain.com</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="ed20e-111">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain.com</AppDomain>`).</span></span>
> 2. <span data-ttu-id="ed20e-112">S’il existe un port explicite pour le domaine, incluez-le (par exemple, `<AppDomain>https://myappdomain.com:9999</AppDomain>` ).</span><span class="sxs-lookup"><span data-stu-id="ed20e-112">If there is an explicit port for the domain, include it (e.g.,`<AppDomain>https://myappdomain.com:9999</AppDomain>`).</span></span>
> 3. <span data-ttu-id="ed20e-113">Si un sous-domaine doit être approuvé, incluez-le (par exemple, `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ).</span><span class="sxs-lookup"><span data-stu-id="ed20e-113">If a subdomain needs to be trusted, include it (e.g.,`<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>`).</span></span> <span data-ttu-id="ed20e-114">Le sous-domaine `mysubdomain.mydomain.com` et `mydomain.com` sont des domaines différents.</span><span class="sxs-lookup"><span data-stu-id="ed20e-114">The subdomain `mysubdomain.mydomain.com` and `mydomain.com` are different domains.</span></span> <span data-ttu-id="ed20e-115">Si les deux doivent être approuvés, les deux doivent se trouver dans des éléments **AppDomain** distincts.</span><span class="sxs-lookup"><span data-stu-id="ed20e-115">If both need to be trusted, then both need to be in separate **AppDomain** elements.</span></span>
> 4. <span data-ttu-id="ed20e-116">Le fait de répertorier le même domaine que celui spécifié dans l' [élément SourceLocation](sourcelocation.md) n’a aucun effet et peut être trompeur.</span><span class="sxs-lookup"><span data-stu-id="ed20e-116">Listing the same domain as the one specified in the [SourceLocation element](sourcelocation.md) has no effect and may be misleading.</span></span> <span data-ttu-id="ed20e-117">En particulier, lorsque vous développez sur `localhost` , vous n’avez pas besoin de créer un élément **AppDomain** pour `localhost` .</span><span class="sxs-lookup"><span data-stu-id="ed20e-117">In particular, when you are developing on `localhost`, you don't need to create an **AppDomain** element for `localhost`.</span></span>
> 5. <span data-ttu-id="ed20e-118">N’incluez pas de segments d’URL au-delà du domaine.</span><span class="sxs-lookup"><span data-stu-id="ed20e-118">Don't include any segments of a URL past the domain.</span></span> <span data-ttu-id="ed20e-119">Par exemple, n’incluez pas l’URL complète d’une page.</span><span class="sxs-lookup"><span data-stu-id="ed20e-119">For example, don't include the full URL of a page.</span></span>
> 6. <span data-ttu-id="ed20e-120">Ne placez *pas* de barre oblique (« / ») sur la valeur.</span><span class="sxs-lookup"><span data-stu-id="ed20e-120">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="ed20e-121">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ed20e-121">Contained in</span></span>

[<span data-ttu-id="ed20e-122">AppDomains</span><span class="sxs-lookup"><span data-stu-id="ed20e-122">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="ed20e-123">Remarques</span><span class="sxs-lookup"><span data-stu-id="ed20e-123">Remarks</span></span>

<span data-ttu-id="ed20e-124">Pour plus d’informations, voir le [manifeste XML de compléments Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="ed20e-124">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
