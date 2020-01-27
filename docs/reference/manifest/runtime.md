---
title: Runtime dans le fichier manifeste
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 8fbad8276b3e1d64a6c443cf57d498597d729282
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41553998"
---
# <a name="runtime-element"></a><span data-ttu-id="8ed1b-102">Élément Runtime</span><span class="sxs-lookup"><span data-stu-id="8ed1b-102">Runtime element</span></span>

<span data-ttu-id="8ed1b-103">Cette fonctionnalité est en aperçu.</span><span class="sxs-lookup"><span data-stu-id="8ed1b-103">This feature is in preview.</span></span> <span data-ttu-id="8ed1b-104">Élément enfant de l' [`<Runtimes>`](runtimes.md) élément.</span><span class="sxs-lookup"><span data-stu-id="8ed1b-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="8ed1b-105">Cet élément facilite le partage des données globales et des appels de fonction entre des fonctions personnalisées Excel et le volet Office de votre complément.</span><span class="sxs-lookup"><span data-stu-id="8ed1b-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span>

<span data-ttu-id="8ed1b-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="8ed1b-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="8ed1b-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8ed1b-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="8ed1b-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8ed1b-108">Contained in</span></span>

- [<span data-ttu-id="8ed1b-109">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="8ed1b-109">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="8ed1b-110">Attributs</span><span class="sxs-lookup"><span data-stu-id="8ed1b-110">Attributes</span></span>

|  <span data-ttu-id="8ed1b-111">Attribut</span><span class="sxs-lookup"><span data-stu-id="8ed1b-111">Attribute</span></span>  |  <span data-ttu-id="8ed1b-112">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="8ed1b-112">Required</span></span>  |  <span data-ttu-id="8ed1b-113">Description</span><span class="sxs-lookup"><span data-stu-id="8ed1b-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="8ed1b-114">**Lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="8ed1b-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="8ed1b-115">Oui</span><span class="sxs-lookup"><span data-stu-id="8ed1b-115">Yes</span></span>  | <span data-ttu-id="8ed1b-116">Doit toujours être mentionné si vous souhaitez que les fonctions personnalisées Excel fonctionnent pendant la fermeture du volet Office de votre complément.</span><span class="sxs-lookup"><span data-stu-id="8ed1b-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="8ed1b-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="8ed1b-117">**resid**</span></span>  |  <span data-ttu-id="8ed1b-118">Oui</span><span class="sxs-lookup"><span data-stu-id="8ed1b-118">Yes</span></span>  | <span data-ttu-id="8ed1b-119">S’il est utilisé pour les fonctions personnalisées Excel `resid` , `TaskPaneAndCustomFunction.Url`le doit pointer vers.</span><span class="sxs-lookup"><span data-stu-id="8ed1b-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="8ed1b-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8ed1b-120">See also</span></span>

- [<span data-ttu-id="8ed1b-121">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="8ed1b-121">Runtimes</span></span>](runtimes.md)
