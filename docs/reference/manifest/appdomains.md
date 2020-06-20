---
title: Élément AppDomains dans le fichier manifeste
description: Répertorie tous les domaines en plus du domaine spécifié dans l' `SourceLocation` élément que votre complément Office utilisera et doit être approuvé par Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778654"
---
# <a name="appdomains-element"></a><span data-ttu-id="345f2-103">AppDomains, élément</span><span class="sxs-lookup"><span data-stu-id="345f2-103">AppDomains element</span></span>

<span data-ttu-id="345f2-104">Répertorie tous les domaines, en plus du domaine spécifié dans l' `SourceLocation` élément, que votre complément Office utilisera et qui doit être approuvé par Office.</span><span class="sxs-lookup"><span data-stu-id="345f2-104">Lists any domains, in addition to the domain specified in the `SourceLocation` element, that your Office Add-in will use and that should be trusted by Office.</span></span> <span data-ttu-id="345f2-105">Cela permet aux pages des domaines d’effectuer des appels à Office.js API depuis des IFrames dans le complément et présente d’autres effets.</span><span class="sxs-lookup"><span data-stu-id="345f2-105">This enables pages in the domains to make calls to Office.js APIs from IFrames within the add-in and has other effects.</span></span> <span data-ttu-id="345f2-106">Pour chaque domaine supplémentaire, indiquez un élément **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="345f2-106">For each additional domain, specify an **AppDomain** element.</span></span>

 <span data-ttu-id="345f2-107">**Type de complément :** Application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="345f2-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="345f2-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="345f2-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="345f2-109">Il existe des restrictions quant à ce qui peut être la valeur d’un élément **AppDomain** .</span><span class="sxs-lookup"><span data-stu-id="345f2-109">There are restrictions on what can be the value of a **AppDomain** element.</span></span> <span data-ttu-id="345f2-110">Pour plus d’informations, consultez la rubrique [AppDomain](appdomain.md).</span><span class="sxs-lookup"><span data-stu-id="345f2-110">For more information, see [AppDomain](appdomain.md).</span></span>

## <a name="contained-in"></a><span data-ttu-id="345f2-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="345f2-111">Contained in</span></span>

[<span data-ttu-id="345f2-112">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="345f2-112">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="345f2-113">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="345f2-113">Can contain</span></span>

[<span data-ttu-id="345f2-114">AppDomain</span><span class="sxs-lookup"><span data-stu-id="345f2-114">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="345f2-115">Remarques</span><span class="sxs-lookup"><span data-stu-id="345f2-115">Remarks</span></span>

<span data-ttu-id="345f2-116">Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément[SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="345f2-116">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="345f2-117">Vous devez indiquer une valeur pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="345f2-117">This element can't be empty.</span></span>
