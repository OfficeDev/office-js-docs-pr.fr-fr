---
title: Runtime dans le fichier manifeste (aperçu)
description: L’élément Runtime configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé pour son ruban, son volet de tâches et ses fonctions personnalisées.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 6237f64fec47ed22b0105bf74c8eb7e2b7c38afe
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717928"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="e4d19-103">Élément Runtime (aperçu)</span><span class="sxs-lookup"><span data-stu-id="e4d19-103">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="e4d19-104">Élément enfant de l' [`<Runtimes>`](runtimes.md) élément.</span><span class="sxs-lookup"><span data-stu-id="e4d19-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="e4d19-105">Cet élément configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé de sorte que votre ruban, votre volet de tâches et vos fonctions personnalisées s’exécutent dans le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="e4d19-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="e4d19-106">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="e4d19-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="e4d19-107">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="e4d19-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e4d19-108">Le runtime partagé est actuellement en préversion et n’est disponible que sur Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="e4d19-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="e4d19-109">Pour essayer les fonctionnalités d’aperçu, vous devrez rejoindre [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="e4d19-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="e4d19-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="e4d19-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="e4d19-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="e4d19-111">Contained in</span></span>

- [<span data-ttu-id="e4d19-112">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="e4d19-112">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="e4d19-113">Attributs</span><span class="sxs-lookup"><span data-stu-id="e4d19-113">Attributes</span></span>

|  <span data-ttu-id="e4d19-114">Attribut</span><span class="sxs-lookup"><span data-stu-id="e4d19-114">Attribute</span></span>  |  <span data-ttu-id="e4d19-115">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="e4d19-115">Required</span></span>  |  <span data-ttu-id="e4d19-116">Description</span><span class="sxs-lookup"><span data-stu-id="e4d19-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e4d19-117">**Lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="e4d19-117">**lifetime="long"**</span></span>  |  <span data-ttu-id="e4d19-118">Oui</span><span class="sxs-lookup"><span data-stu-id="e4d19-118">Yes</span></span>  | <span data-ttu-id="e4d19-119">Doit toujours être `long` utilisé pour utiliser un runtime partagé pour le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="e4d19-119">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="e4d19-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="e4d19-120">**resid**</span></span>  |  <span data-ttu-id="e4d19-121">Oui</span><span class="sxs-lookup"><span data-stu-id="e4d19-121">Yes</span></span>  | <span data-ttu-id="e4d19-122">Spécifie l’URL de la page HTML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="e4d19-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="e4d19-123">L `resid` 'doit correspondre `id` à un attribut `Url` d’un élément `Resources` dans l’élément.</span><span class="sxs-lookup"><span data-stu-id="e4d19-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="e4d19-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e4d19-124">See also</span></span>

- [<span data-ttu-id="e4d19-125">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="e4d19-125">Runtimes</span></span>](runtimes.md)
