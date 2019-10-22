---
title: Élément AppDomain dans le fichier manifeste
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 2f65302d1ac3d85f2867cd13501bc67606cd00b5
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/21/2019
ms.locfileid: "35575638"
---
# <a name="appdomain-element"></a><span data-ttu-id="63b2b-102">AppDomain, élément</span><span class="sxs-lookup"><span data-stu-id="63b2b-102">AppDomain element</span></span>

<span data-ttu-id="63b2b-103">Spécifie les domaines supplémentaires qui chargent des pages dans la fenêtre du complément.</span><span class="sxs-lookup"><span data-stu-id="63b2b-103">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="63b2b-104">Il répertorie également les domaines approuvés à partir desquels les appels de l’API Office. js peuvent être effectués depuis des IFrames au sein du complément.</span><span class="sxs-lookup"><span data-stu-id="63b2b-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="63b2b-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="63b2b-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="63b2b-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="63b2b-106">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="63b2b-107">La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="63b2b-107">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="63b2b-108">Ne placez *pas* de barre oblique (« / ») sur la valeur.</span><span class="sxs-lookup"><span data-stu-id="63b2b-108">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="63b2b-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="63b2b-109">Contained in</span></span>

[<span data-ttu-id="63b2b-110">AppDomains</span><span class="sxs-lookup"><span data-stu-id="63b2b-110">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="63b2b-111">Remarques</span><span class="sxs-lookup"><span data-stu-id="63b2b-111">Remarks</span></span>

<span data-ttu-id="63b2b-112">Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="63b2b-112">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="63b2b-113">Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="63b2b-113">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
