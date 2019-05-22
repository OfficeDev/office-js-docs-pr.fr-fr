---
title: Élément AppDomain dans le fichier manifeste
description: ''
ms.date: 05/15/2019
localization_priority: Normal
ms.openlocfilehash: b1d71648cc7646eec246f3d0a8113c843eed2e74
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337194"
---
# <a name="appdomain-element"></a><span data-ttu-id="6f3da-102">AppDomain, élément</span><span class="sxs-lookup"><span data-stu-id="6f3da-102">AppDomain element</span></span>

<span data-ttu-id="6f3da-103">Indique un domaine supplémentaire permettant de charger des pages dans la fenêtre du complément.</span><span class="sxs-lookup"><span data-stu-id="6f3da-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="6f3da-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="6f3da-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6f3da-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="6f3da-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="6f3da-106">La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="6f3da-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="6f3da-107">Ne placez *pas* de barre oblique («/») sur la valeur.</span><span class="sxs-lookup"><span data-stu-id="6f3da-107">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="6f3da-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="6f3da-108">Contained in</span></span>

[<span data-ttu-id="6f3da-109">AppDomains</span><span class="sxs-lookup"><span data-stu-id="6f3da-109">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="6f3da-110">Remarques</span><span class="sxs-lookup"><span data-stu-id="6f3da-110">Remarks</span></span>

<span data-ttu-id="6f3da-111">Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="6f3da-111">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="6f3da-112">Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="6f3da-112">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
