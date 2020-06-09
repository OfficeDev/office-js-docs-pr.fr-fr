---
title: Élément AppDomain dans le fichier manifeste
description: Spécifie les domaines supplémentaires qui chargent des pages dans la fenêtre du complément.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: ddacae6d8aa45ccccd3a8acbb42de48b152fb9d2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608774"
---
# <a name="appdomain-element"></a><span data-ttu-id="e2fc8-103">AppDomain, élément</span><span class="sxs-lookup"><span data-stu-id="e2fc8-103">AppDomain element</span></span>

<span data-ttu-id="e2fc8-104">Spécifie les domaines supplémentaires qui chargent des pages dans la fenêtre du complément.</span><span class="sxs-lookup"><span data-stu-id="e2fc8-104">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="e2fc8-105">Il répertorie également les domaines approuvés à partir desquels les appels de l’API Office. js peuvent être effectués depuis des IFrames au sein du complément.</span><span class="sxs-lookup"><span data-stu-id="e2fc8-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="e2fc8-106">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="e2fc8-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e2fc8-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="e2fc8-107">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="e2fc8-108">La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="e2fc8-108">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="e2fc8-109">Ne placez *pas* de barre oblique (« / ») sur la valeur.</span><span class="sxs-lookup"><span data-stu-id="e2fc8-109">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="e2fc8-110">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="e2fc8-110">Contained in</span></span>

[<span data-ttu-id="e2fc8-111">AppDomains</span><span class="sxs-lookup"><span data-stu-id="e2fc8-111">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="e2fc8-112">Remarques</span><span class="sxs-lookup"><span data-stu-id="e2fc8-112">Remarks</span></span>

<span data-ttu-id="e2fc8-113">Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="e2fc8-113">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="e2fc8-114">Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="e2fc8-114">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
