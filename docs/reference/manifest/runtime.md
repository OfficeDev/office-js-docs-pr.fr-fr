---
title: Runtime dans le fichier manifeste
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111169"
---
# <a name="runtime-element"></a><span data-ttu-id="616f5-102">Élément Runtime</span><span class="sxs-lookup"><span data-stu-id="616f5-102">Runtime element</span></span>

<span data-ttu-id="616f5-103">Cette fonctionnalité est en aperçu.</span><span class="sxs-lookup"><span data-stu-id="616f5-103">This feature is in preview.</span></span> <span data-ttu-id="616f5-104">Élément enfant de l' [`<Runtimes>`](runtime.md) élément.</span><span class="sxs-lookup"><span data-stu-id="616f5-104">Child element of the [`<Runtimes>`](runtime.md) element.</span></span> <span data-ttu-id="616f5-105">Cet élément facilite le partage des données globales et des appels de fonction entre des fonctions personnalisées Excel et le volet Office de votre complément.</span><span class="sxs-lookup"><span data-stu-id="616f5-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span> 

## <a name="contained-in"></a><span data-ttu-id="616f5-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="616f5-106">Contained in</span></span>

<span data-ttu-id="616f5-107">-[Runtimes](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="616f5-107">-[Runtimes](runtimes.md)</span></span>

<span data-ttu-id="616f5-108">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="616f5-108">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="616f5-109">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="616f5-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a><span data-ttu-id="616f5-110">Attributs</span><span class="sxs-lookup"><span data-stu-id="616f5-110">Attributes</span></span>

|  <span data-ttu-id="616f5-111">Attribut</span><span class="sxs-lookup"><span data-stu-id="616f5-111">Attribute</span></span>  |  <span data-ttu-id="616f5-112">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="616f5-112">Required</span></span>  |  <span data-ttu-id="616f5-113">Description</span><span class="sxs-lookup"><span data-stu-id="616f5-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="616f5-114">**Lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="616f5-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="616f5-115">Oui</span><span class="sxs-lookup"><span data-stu-id="616f5-115">Yes</span></span>  | <span data-ttu-id="616f5-116">Doit toujours être mentionné si vous souhaitez que les fonctions personnalisées Excel fonctionnent pendant la fermeture du volet Office de votre complément.</span><span class="sxs-lookup"><span data-stu-id="616f5-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="616f5-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="616f5-117">**resid**</span></span>  |  <span data-ttu-id="616f5-118">Oui</span><span class="sxs-lookup"><span data-stu-id="616f5-118">Yes</span></span>  | <span data-ttu-id="616f5-119">S’il est utilisé pour les fonctions personnalisées Excel `resid` , `TaskPaneAndCustomFunction.Url`le doit pointer vers.</span><span class="sxs-lookup"><span data-stu-id="616f5-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="616f5-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="616f5-120">See also</span></span>

<span data-ttu-id="616f5-121">-[Runtime](runtime.md)</span><span class="sxs-lookup"><span data-stu-id="616f5-121">-[Runtime](runtime.md)</span></span>
