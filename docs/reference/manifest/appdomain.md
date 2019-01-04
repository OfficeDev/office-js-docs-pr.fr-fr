---
title: Élément AppDomain dans le fichier manifeste
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: 2b55f2c1ea7a2a3dc7dec42c913d74006c0f2e3b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433067"
---
# <a name="appdomain-element"></a><span data-ttu-id="1cd15-102">AppDomain, élément</span><span class="sxs-lookup"><span data-stu-id="1cd15-102">AppDomain element</span></span>

<span data-ttu-id="1cd15-103">Indique un domaine supplémentaire permettant de charger des pages dans la fenêtre du complément.</span><span class="sxs-lookup"><span data-stu-id="1cd15-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="1cd15-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="1cd15-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1cd15-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="1cd15-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="1cd15-106">La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="1cd15-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="1cd15-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="1cd15-107">Contained in</span></span>

[<span data-ttu-id="1cd15-108">AppDomains</span><span class="sxs-lookup"><span data-stu-id="1cd15-108">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="1cd15-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="1cd15-109">Remarks</span></span>

<span data-ttu-id="1cd15-110">Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="1cd15-110">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="1cd15-111">Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="1cd15-111">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
