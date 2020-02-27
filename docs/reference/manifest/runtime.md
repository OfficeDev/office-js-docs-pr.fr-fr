---
title: Runtime dans le fichier manifeste (aperçu)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 26702896604f9ecf4c69296e5110efe5cdf4218b
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283883"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="4f38a-102">Élément Runtime (aperçu)</span><span class="sxs-lookup"><span data-stu-id="4f38a-102">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="4f38a-103">Élément enfant de l' [`<Runtimes>`](runtimes.md) élément.</span><span class="sxs-lookup"><span data-stu-id="4f38a-103">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="4f38a-104">Cet élément configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé de sorte que votre ruban, votre volet de tâches et vos fonctions personnalisées s’exécutent dans le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="4f38a-104">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="4f38a-105">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="4f38a-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="4f38a-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="4f38a-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
<span data-ttu-id="4f38a-107"><<<<<<< en-tête Shared Runtime est actuellement en préversion et n’est disponible que sur Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="4f38a-107"><<<<<<< HEAD Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="4f38a-108">Pour essayer les fonctionnalités d’aperçu, vous devrez rejoindre [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="4f38a-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="4f38a-109">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="4f38a-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="4f38a-110">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="4f38a-110">Contained in</span></span>

- [<span data-ttu-id="4f38a-111">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="4f38a-111">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="4f38a-112">Attributs</span><span class="sxs-lookup"><span data-stu-id="4f38a-112">Attributes</span></span>

|  <span data-ttu-id="4f38a-113">Attribut</span><span class="sxs-lookup"><span data-stu-id="4f38a-113">Attribute</span></span>  |  <span data-ttu-id="4f38a-114">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="4f38a-114">Required</span></span>  |  <span data-ttu-id="4f38a-115">Description</span><span class="sxs-lookup"><span data-stu-id="4f38a-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4f38a-116">**Lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="4f38a-116">**lifetime="long"**</span></span>  |  <span data-ttu-id="4f38a-117">Oui</span><span class="sxs-lookup"><span data-stu-id="4f38a-117">Yes</span></span>  | <span data-ttu-id="4f38a-118">Doit toujours être `long` utilisé pour utiliser un runtime partagé pour le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="4f38a-118">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="4f38a-119">**resid**</span><span class="sxs-lookup"><span data-stu-id="4f38a-119">**resid**</span></span>  |  <span data-ttu-id="4f38a-120">Oui</span><span class="sxs-lookup"><span data-stu-id="4f38a-120">Yes</span></span>  | <span data-ttu-id="4f38a-121">Spécifie l’URL de la page HTML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="4f38a-121">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="4f38a-122">L `resid` 'doit correspondre `id` à un attribut `Url` d’un élément `Resources` dans l’élément.</span><span class="sxs-lookup"><span data-stu-id="4f38a-122">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="4f38a-123">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4f38a-123">See also</span></span>

- [<span data-ttu-id="4f38a-124">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="4f38a-124">Runtimes</span></span>](runtimes.md)
