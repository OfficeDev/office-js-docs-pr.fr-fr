---
title: Élément AppDomains dans le fichier manifeste
description: Répertorie tous les domaines en plus du domaine spécifié dans l' `SourceLocation` élément qui sera utilisé par votre complément Office pour charger des pages.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 9183f1815e97bd8d4ac1a7e2cf72d5547d153f7e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608767"
---
# <a name="appdomains-element"></a><span data-ttu-id="a4055-103">AppDomains, élément</span><span class="sxs-lookup"><span data-stu-id="a4055-103">AppDomains element</span></span>

<span data-ttu-id="a4055-104">Répertorie tous les domaines en plus du domaine spécifié dans l' `SourceLocation` élément qui sera utilisé par votre complément Office pour charger des pages.</span><span class="sxs-lookup"><span data-stu-id="a4055-104">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="a4055-105">Il répertorie également les domaines approuvés à partir desquels les appels de l’API Office. js peuvent être effectués depuis des IFrames au sein du complément.</span><span class="sxs-lookup"><span data-stu-id="a4055-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="a4055-106">Pour chaque domaine supplémentaire, indiquez un élément AppDomain.</span><span class="sxs-lookup"><span data-stu-id="a4055-106">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="a4055-107">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="a4055-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a4055-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="a4055-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="a4055-109">La valeur de chaque élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="a4055-109">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="a4055-110">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="a4055-110">Contained in</span></span>

[<span data-ttu-id="a4055-111">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="a4055-111">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="a4055-112">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="a4055-112">Can contain</span></span>

[<span data-ttu-id="a4055-113">AppDomain</span><span class="sxs-lookup"><span data-stu-id="a4055-113">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="a4055-114">Remarques</span><span class="sxs-lookup"><span data-stu-id="a4055-114">Remarks</span></span>

<span data-ttu-id="a4055-115">Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément[SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="a4055-115">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="a4055-116">Pour charger des pages qui ne sont pas dans le même domaine que le complément, spécifiez les domaines à l’aide des éléments **AppDomains** et **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="a4055-116">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="a4055-117">Vous devez indiquer une valeur pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="a4055-117">This element can't be empty.</span></span>
