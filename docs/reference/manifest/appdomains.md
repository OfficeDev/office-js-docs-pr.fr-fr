---
title: Élément AppDomains dans le fichier manifeste
description: ''
ms.date: 12/13/2018
localization_priority: Normal
ms.openlocfilehash: 65391c9529e7ddaa9726d0b58accf90c5b9babef
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450645"
---
# <a name="appdomains-element"></a><span data-ttu-id="fcd87-102">AppDomains, élément</span><span class="sxs-lookup"><span data-stu-id="fcd87-102">AppDomains element</span></span>

<span data-ttu-id="fcd87-p101">Répertorie tout domaine supplémentaire qui sera utilisé par votre complément Office pour charger des pages en plus du domaine spécifié dans l’élément SourceLocation. Pour chaque domaine supplémentaire, indiquez un élément AppDomain.</span><span class="sxs-lookup"><span data-stu-id="fcd87-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="fcd87-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="fcd87-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="fcd87-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="fcd87-106">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="fcd87-107">La valeur de chaque élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="fcd87-107">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="fcd87-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="fcd87-108">Contained in</span></span>

[<span data-ttu-id="fcd87-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="fcd87-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="fcd87-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="fcd87-110">Can contain</span></span>

[<span data-ttu-id="fcd87-111">AppDomain</span><span class="sxs-lookup"><span data-stu-id="fcd87-111">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="fcd87-112">Remarques</span><span class="sxs-lookup"><span data-stu-id="fcd87-112">Remarks</span></span>

<span data-ttu-id="fcd87-113">Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément[SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="fcd87-113">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="fcd87-114">Pour charger des pages qui ne sont pas dans le même domaine que le complément, spécifiez les domaines à l’aide des éléments **AppDomains** et **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="fcd87-114">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="fcd87-115">Vous devez indiquer une valeur pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="fcd87-115">This element can't be empty.</span></span>
