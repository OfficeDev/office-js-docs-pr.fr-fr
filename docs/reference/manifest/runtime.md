---
title: Runtime dans le fichier manifeste
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 945a30527632b23a594d7bfb82cec94e74754249
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120634"
---
# <a name="runtime-element"></a><span data-ttu-id="c905d-102">Élément Runtime</span><span class="sxs-lookup"><span data-stu-id="c905d-102">Runtime element</span></span>

<span data-ttu-id="c905d-103">Cette fonctionnalité est en aperçu.</span><span class="sxs-lookup"><span data-stu-id="c905d-103">This feature is in preview.</span></span> <span data-ttu-id="c905d-104">Élément enfant de l' [`<Runtimes>`](runtime.md) élément.</span><span class="sxs-lookup"><span data-stu-id="c905d-104">Child element of the [`<Runtimes>`](runtime.md) element.</span></span> <span data-ttu-id="c905d-105">Cet élément facilite le partage des données globales et des appels de fonction entre des fonctions personnalisées Excel et le volet Office de votre complément.</span><span class="sxs-lookup"><span data-stu-id="c905d-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span>

<span data-ttu-id="c905d-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="c905d-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="c905d-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="c905d-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="c905d-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="c905d-108">Contained in</span></span>

<span data-ttu-id="c905d-109">-[Runtimes](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="c905d-109">-[Runtimes](runtimes.md)</span></span>

## <a name="attributes"></a><span data-ttu-id="c905d-110">Attributs</span><span class="sxs-lookup"><span data-stu-id="c905d-110">Attributes</span></span>

|  <span data-ttu-id="c905d-111">Attribut</span><span class="sxs-lookup"><span data-stu-id="c905d-111">Attribute</span></span>  |  <span data-ttu-id="c905d-112">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="c905d-112">Required</span></span>  |  <span data-ttu-id="c905d-113">Description</span><span class="sxs-lookup"><span data-stu-id="c905d-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c905d-114">**Lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="c905d-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="c905d-115">Oui</span><span class="sxs-lookup"><span data-stu-id="c905d-115">Yes</span></span>  | <span data-ttu-id="c905d-116">Doit toujours être mentionné si vous souhaitez que les fonctions personnalisées Excel fonctionnent pendant la fermeture du volet Office de votre complément.</span><span class="sxs-lookup"><span data-stu-id="c905d-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="c905d-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="c905d-117">**resid**</span></span>  |  <span data-ttu-id="c905d-118">Oui</span><span class="sxs-lookup"><span data-stu-id="c905d-118">Yes</span></span>  | <span data-ttu-id="c905d-119">S’il est utilisé pour les fonctions personnalisées Excel `resid` , `TaskPaneAndCustomFunction.Url`le doit pointer vers.</span><span class="sxs-lookup"><span data-stu-id="c905d-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="c905d-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c905d-120">See also</span></span>

<span data-ttu-id="c905d-121">-[Runtime](runtime.md)</span><span class="sxs-lookup"><span data-stu-id="c905d-121">-[Runtime](runtime.md)</span></span>
