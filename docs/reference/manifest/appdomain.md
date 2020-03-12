---
title: Élément AppDomain dans le fichier manifeste
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: da28b3b4dec5d669462a781db3c0628bd32c7182
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596787"
---
# <a name="appdomain-element"></a><span data-ttu-id="b919e-102">AppDomain, élément</span><span class="sxs-lookup"><span data-stu-id="b919e-102">AppDomain element</span></span>

<span data-ttu-id="b919e-103">Spécifie les domaines supplémentaires qui chargent des pages dans la fenêtre du complément.</span><span class="sxs-lookup"><span data-stu-id="b919e-103">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="b919e-104">Il répertorie également les domaines approuvés à partir desquels les appels de l’API Office. js peuvent être effectués depuis des IFrames au sein du complément.</span><span class="sxs-lookup"><span data-stu-id="b919e-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="b919e-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="b919e-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b919e-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b919e-106">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="b919e-107">La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="b919e-107">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="b919e-108">Ne placez *pas* de barre oblique (« / ») sur la valeur.</span><span class="sxs-lookup"><span data-stu-id="b919e-108">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="b919e-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b919e-109">Contained in</span></span>

[<span data-ttu-id="b919e-110">AppDomains</span><span class="sxs-lookup"><span data-stu-id="b919e-110">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="b919e-111">Remarques</span><span class="sxs-lookup"><span data-stu-id="b919e-111">Remarks</span></span>

<span data-ttu-id="b919e-112">Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="b919e-112">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="b919e-113">Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="b919e-113">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
