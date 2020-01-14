---
title: Runtimes dans le fichier manifeste
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: ec2b85a92325eb4e36c61f731369ec54d44ef169
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111176"
---
# <a name="runtimes-element"></a><span data-ttu-id="12475-102">Élément runtimes</span><span class="sxs-lookup"><span data-stu-id="12475-102">Runtimes element</span></span>

<span data-ttu-id="12475-103">Cette fonctionnalité est en aperçu.</span><span class="sxs-lookup"><span data-stu-id="12475-103">This feature is in preview.</span></span> <span data-ttu-id="12475-104">Spécifie le runtime de votre complément et permet aux fonctions personnalisées et au volet Office de partager des données globales et d’effectuer des appels de fonction.</span><span class="sxs-lookup"><span data-stu-id="12475-104">Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other.</span></span> <span data-ttu-id="12475-105">Doit suivre l' `<Host>` élément dans votre fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="12475-105">Should follow the `<Host>` element in your manifest file.</span></span>

<span data-ttu-id="12475-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="12475-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="12475-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="12475-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a><span data-ttu-id="12475-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="12475-108">Child elements</span></span>

|  <span data-ttu-id="12475-109">Élément</span><span class="sxs-lookup"><span data-stu-id="12475-109">Element</span></span> |  <span data-ttu-id="12475-110">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="12475-110">Required</span></span>  |  <span data-ttu-id="12475-111">Description</span><span class="sxs-lookup"><span data-stu-id="12475-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="12475-112">**Runtime**</span><span class="sxs-lookup"><span data-stu-id="12475-112">**Runtime**</span></span>     | <span data-ttu-id="12475-113">Oui</span><span class="sxs-lookup"><span data-stu-id="12475-113">Yes</span></span> |  <span data-ttu-id="12475-114">Le runtime de votre complément, souvent utilisé avec des fonctions personnalisées Excel.</span><span class="sxs-lookup"><span data-stu-id="12475-114">The Runtime for your add-in, often used with Excel custom functions.</span></span>

## <a name="see-also"></a><span data-ttu-id="12475-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="12475-115">See also</span></span>

<span data-ttu-id="12475-116">-[Runtimes](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="12475-116">-[Runtimes](runtimes.md)</span></span>
