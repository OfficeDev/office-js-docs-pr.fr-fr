---
title: Élément AppDomains dans le fichier manifeste
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: b6db3d46d004021f25edd5733566544010abb457
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575330"
---
# <a name="appdomains-element"></a><span data-ttu-id="37e81-102">AppDomains, élément</span><span class="sxs-lookup"><span data-stu-id="37e81-102">AppDomains element</span></span>

<span data-ttu-id="37e81-103">Répertorie tous les domaines en plus du domaine spécifié dans `SourceLocation` l’élément qui sera utilisé par votre complément Office pour charger des pages.</span><span class="sxs-lookup"><span data-stu-id="37e81-103">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="37e81-104">Il répertorie également les domaines approuvés à partir desquels les appels de l’API Office. js peuvent être effectués depuis des IFrames au sein du complément.</span><span class="sxs-lookup"><span data-stu-id="37e81-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="37e81-105">Pour chaque domaine supplémentaire, indiquez un élément AppDomain.</span><span class="sxs-lookup"><span data-stu-id="37e81-105">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="37e81-106">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="37e81-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="37e81-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="37e81-107">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="37e81-108">La valeur de chaque élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="37e81-108">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="37e81-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="37e81-109">Contained in</span></span>

[<span data-ttu-id="37e81-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="37e81-110">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="37e81-111">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="37e81-111">Can contain</span></span>

[<span data-ttu-id="37e81-112">AppDomain</span><span class="sxs-lookup"><span data-stu-id="37e81-112">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="37e81-113">Remarques</span><span class="sxs-lookup"><span data-stu-id="37e81-113">Remarks</span></span>

<span data-ttu-id="37e81-114">Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément[SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="37e81-114">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="37e81-115">Pour charger des pages qui ne sont pas dans le même domaine que le complément, spécifiez les domaines à l’aide des éléments **AppDomains** et **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="37e81-115">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="37e81-116">Vous devez indiquer une valeur pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="37e81-116">This element can't be empty.</span></span>
