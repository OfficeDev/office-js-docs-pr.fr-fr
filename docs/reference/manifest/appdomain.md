---
title: Élément AppDomain dans le fichier manifeste
description: Spécifie les domaines supplémentaires qui chargent des pages dans la fenêtre du complément.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 6990f759df806f24b1d617c036bc1a452e6da38f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718453"
---
# <a name="appdomain-element"></a><span data-ttu-id="d137f-103">AppDomain, élément</span><span class="sxs-lookup"><span data-stu-id="d137f-103">AppDomain element</span></span>

<span data-ttu-id="d137f-104">Spécifie les domaines supplémentaires qui chargent des pages dans la fenêtre du complément.</span><span class="sxs-lookup"><span data-stu-id="d137f-104">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="d137f-105">Il répertorie également les domaines approuvés à partir desquels les appels de l’API Office. js peuvent être effectués depuis des IFrames au sein du complément.</span><span class="sxs-lookup"><span data-stu-id="d137f-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="d137f-106">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="d137f-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d137f-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="d137f-107">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="d137f-108">La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="d137f-108">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="d137f-109">Ne placez *pas* de barre oblique (« / ») sur la valeur.</span><span class="sxs-lookup"><span data-stu-id="d137f-109">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="d137f-110">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="d137f-110">Contained in</span></span>

[<span data-ttu-id="d137f-111">AppDomains</span><span class="sxs-lookup"><span data-stu-id="d137f-111">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="d137f-112">Remarques</span><span class="sxs-lookup"><span data-stu-id="d137f-112">Remarks</span></span>

<span data-ttu-id="d137f-113">Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="d137f-113">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="d137f-114">Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="d137f-114">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
