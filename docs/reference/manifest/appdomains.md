---
title: Élément AppDomains dans le fichier manifeste
description: Répertorie tous les domaines en plus du domaine spécifié dans `SourceLocation` l’élément qui sera utilisé par votre complément Office pour charger des pages.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: f60579d773e81a7e8006bafcf1c151874af42aeb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720700"
---
# <a name="appdomains-element"></a><span data-ttu-id="b9dfa-103">AppDomains, élément</span><span class="sxs-lookup"><span data-stu-id="b9dfa-103">AppDomains element</span></span>

<span data-ttu-id="b9dfa-104">Répertorie tous les domaines en plus du domaine spécifié dans `SourceLocation` l’élément qui sera utilisé par votre complément Office pour charger des pages.</span><span class="sxs-lookup"><span data-stu-id="b9dfa-104">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="b9dfa-105">Il répertorie également les domaines approuvés à partir desquels les appels de l’API Office. js peuvent être effectués depuis des IFrames au sein du complément.</span><span class="sxs-lookup"><span data-stu-id="b9dfa-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="b9dfa-106">Pour chaque domaine supplémentaire, indiquez un élément AppDomain.</span><span class="sxs-lookup"><span data-stu-id="b9dfa-106">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="b9dfa-107">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="b9dfa-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b9dfa-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b9dfa-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="b9dfa-109">La valeur de chaque élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="b9dfa-109">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="b9dfa-110">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b9dfa-110">Contained in</span></span>

[<span data-ttu-id="b9dfa-111">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b9dfa-111">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="b9dfa-112">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="b9dfa-112">Can contain</span></span>

[<span data-ttu-id="b9dfa-113">AppDomain</span><span class="sxs-lookup"><span data-stu-id="b9dfa-113">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="b9dfa-114">Remarques</span><span class="sxs-lookup"><span data-stu-id="b9dfa-114">Remarks</span></span>

<span data-ttu-id="b9dfa-115">Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément[SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="b9dfa-115">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="b9dfa-116">Pour charger des pages qui ne sont pas dans le même domaine que le complément, spécifiez les domaines à l’aide des éléments **AppDomains** et **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="b9dfa-116">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="b9dfa-117">Vous devez indiquer une valeur pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="b9dfa-117">This element can't be empty.</span></span>
