---
title: Runtimes dans le fichier manifeste (aperçu)
description: L’élément runtimes spécifie le runtime de votre complément.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 5797aa78ae3667461de48de481ff44f14c307ced
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720420"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="9d2e2-103">Runtimes, élément (aperçu)</span><span class="sxs-lookup"><span data-stu-id="9d2e2-103">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="9d2e2-104">Spécifie le runtime de votre complément et active des fonctions personnalisées, des boutons du ruban et le volet des tâches pour utiliser le même Runtime JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9d2e2-104">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="9d2e2-105">Enfant de l' `<Host>` élément dans votre fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="9d2e2-105">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="9d2e2-106">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="9d2e2-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="9d2e2-107">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="9d2e2-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9d2e2-108">Le runtime partagé est actuellement en préversion et n’est disponible que sur Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="9d2e2-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="9d2e2-109">Pour essayer les fonctionnalités d’aperçu, vous devrez rejoindre [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="9d2e2-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="9d2e2-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="9d2e2-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="9d2e2-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="9d2e2-111">Contained in</span></span> 
[<span data-ttu-id="9d2e2-112">Host</span><span class="sxs-lookup"><span data-stu-id="9d2e2-112">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="9d2e2-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9d2e2-113">Child elements</span></span>

|  <span data-ttu-id="9d2e2-114">Élément</span><span class="sxs-lookup"><span data-stu-id="9d2e2-114">Element</span></span> |  <span data-ttu-id="9d2e2-115">Requis</span><span class="sxs-lookup"><span data-stu-id="9d2e2-115">Required</span></span>  |  <span data-ttu-id="9d2e2-116">Description</span><span class="sxs-lookup"><span data-stu-id="9d2e2-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="9d2e2-117">**Runtime**</span><span class="sxs-lookup"><span data-stu-id="9d2e2-117">**Runtime**</span></span>     | <span data-ttu-id="9d2e2-118">Oui</span><span class="sxs-lookup"><span data-stu-id="9d2e2-118">Yes</span></span> |  <span data-ttu-id="9d2e2-119">Le runtime de votre complément.</span><span class="sxs-lookup"><span data-stu-id="9d2e2-119">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="9d2e2-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9d2e2-120">See also</span></span>

- [<span data-ttu-id="9d2e2-121">Runtime</span><span class="sxs-lookup"><span data-stu-id="9d2e2-121">Runtime</span></span>](runtime.md)
