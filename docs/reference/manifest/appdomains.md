---
title: Élément AppDomains dans le fichier manifeste
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: cc2f5ade0bdda214c85490f8e474b42f921edbe8
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433677"
---
# <a name="appdomains-element"></a><span data-ttu-id="21416-102">AppDomains, élément</span><span class="sxs-lookup"><span data-stu-id="21416-102">AppDomains element</span></span>

<span data-ttu-id="21416-p101">Répertorie tout domaine supplémentaire qui sera utilisé par votre complément Office pour charger des pages en plus du domaine spécifié dans l’élément SourceLocation. Pour chaque domaine supplémentaire, indiquez un élément AppDomain.</span><span class="sxs-lookup"><span data-stu-id="21416-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="21416-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="21416-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="21416-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="21416-106">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="21416-107">La valeur de chaque élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="21416-107">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="21416-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="21416-108">Contained in</span></span>

[<span data-ttu-id="21416-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="21416-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="21416-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="21416-110">Can contain</span></span>

[<span data-ttu-id="21416-111">AppDomain</span><span class="sxs-lookup"><span data-stu-id="21416-111">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="21416-112">Remarques</span><span class="sxs-lookup"><span data-stu-id="21416-112">Remarks</span></span>

<span data-ttu-id="21416-113">Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément[SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="21416-113">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="21416-114">Pour charger des pages qui ne sont pas dans le même domaine que le complément, spécifiez les domaines à l’aide des éléments **AppDomains** et **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="21416-114">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="21416-115">Vous devez indiquer une valeur pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="21416-115">This element can't be empty.</span></span>
