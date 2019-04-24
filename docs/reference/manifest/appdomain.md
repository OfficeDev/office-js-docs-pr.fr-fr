---
title: Élément AppDomain dans le fichier manifeste
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: 8216603c87a7dcafde84d25a82f068c9aa86ed96
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450750"
---
# <a name="appdomain-element"></a><span data-ttu-id="0bbe2-102">AppDomain, élément</span><span class="sxs-lookup"><span data-stu-id="0bbe2-102">AppDomain element</span></span>

<span data-ttu-id="0bbe2-103">Indique un domaine supplémentaire permettant de charger des pages dans la fenêtre du complément.</span><span class="sxs-lookup"><span data-stu-id="0bbe2-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="0bbe2-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="0bbe2-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0bbe2-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="0bbe2-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="0bbe2-106">La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="0bbe2-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="0bbe2-107">Ne placez *pas* de barre oblique («/») sur la valeur.</span><span class="sxs-lookup"><span data-stu-id="0bbe2-107">Do *not* put a closing slash, "/", on the the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="0bbe2-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="0bbe2-108">Contained in</span></span>

[<span data-ttu-id="0bbe2-109">AppDomains</span><span class="sxs-lookup"><span data-stu-id="0bbe2-109">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="0bbe2-110">Remarques</span><span class="sxs-lookup"><span data-stu-id="0bbe2-110">Remarks</span></span>

<span data-ttu-id="0bbe2-111">Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="0bbe2-111">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="0bbe2-112">Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="0bbe2-112">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
