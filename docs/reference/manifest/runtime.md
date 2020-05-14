---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé pour son ruban, son volet de tâches et ses fonctions personnalisées.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217759"
---
# <a name="runtime-element"></a><span data-ttu-id="de604-103">Élément Runtime</span><span class="sxs-lookup"><span data-stu-id="de604-103">Runtime element</span></span>

<span data-ttu-id="de604-104">Élément enfant de l' [`<Runtimes>`](runtimes.md) élément.</span><span class="sxs-lookup"><span data-stu-id="de604-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="de604-105">Cet élément configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé de sorte que votre ruban, votre volet de tâches et vos fonctions personnalisées s’exécutent dans le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="de604-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="de604-106">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="de604-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="de604-107">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="de604-107">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="de604-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="de604-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="de604-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="de604-109">Contained in</span></span>

- [<span data-ttu-id="de604-110">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="de604-110">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="de604-111">Attributs</span><span class="sxs-lookup"><span data-stu-id="de604-111">Attributes</span></span>

|  <span data-ttu-id="de604-112">Attribut</span><span class="sxs-lookup"><span data-stu-id="de604-112">Attribute</span></span>  |  <span data-ttu-id="de604-113">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="de604-113">Required</span></span>  |  <span data-ttu-id="de604-114">Description</span><span class="sxs-lookup"><span data-stu-id="de604-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="de604-115">**Lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="de604-115">**lifetime="long"**</span></span>  |  <span data-ttu-id="de604-116">Oui</span><span class="sxs-lookup"><span data-stu-id="de604-116">Yes</span></span>  | <span data-ttu-id="de604-117">Doit toujours être `long` utilisé pour utiliser un runtime partagé pour le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="de604-117">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="de604-118">**resid**</span><span class="sxs-lookup"><span data-stu-id="de604-118">**resid**</span></span>  |  <span data-ttu-id="de604-119">Oui</span><span class="sxs-lookup"><span data-stu-id="de604-119">Yes</span></span>  | <span data-ttu-id="de604-120">Spécifie l’URL de la page HTML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="de604-120">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="de604-121">L' `resid` doit correspondre à un `id` attribut d’un `Url` élément dans l' `Resources` élément.</span><span class="sxs-lookup"><span data-stu-id="de604-121">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="de604-122">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="de604-122">See also</span></span>

- [<span data-ttu-id="de604-123">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="de604-123">Runtimes</span></span>](runtimes.md)
