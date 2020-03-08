---
title: Runtime dans le fichier manifeste (aperçu)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: dd51c5b317700f92ee74c94835e68523371789f8
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561827"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="8aba3-102">Élément Runtime (aperçu)</span><span class="sxs-lookup"><span data-stu-id="8aba3-102">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="8aba3-103">Élément enfant de l' [`<Runtimes>`](runtimes.md) élément.</span><span class="sxs-lookup"><span data-stu-id="8aba3-103">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="8aba3-104">Cet élément configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé de sorte que votre ruban, votre volet de tâches et vos fonctions personnalisées s’exécutent dans le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="8aba3-104">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="8aba3-105">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="8aba3-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="8aba3-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="8aba3-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8aba3-107">Le runtime partagé est actuellement en préversion et n’est disponible que sur Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="8aba3-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="8aba3-108">Pour essayer les fonctionnalités d’aperçu, vous devrez rejoindre [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="8aba3-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="8aba3-109">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8aba3-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="8aba3-110">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8aba3-110">Contained in</span></span>

- [<span data-ttu-id="8aba3-111">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="8aba3-111">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="8aba3-112">Attributs</span><span class="sxs-lookup"><span data-stu-id="8aba3-112">Attributes</span></span>

|  <span data-ttu-id="8aba3-113">Attribut</span><span class="sxs-lookup"><span data-stu-id="8aba3-113">Attribute</span></span>  |  <span data-ttu-id="8aba3-114">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="8aba3-114">Required</span></span>  |  <span data-ttu-id="8aba3-115">Description</span><span class="sxs-lookup"><span data-stu-id="8aba3-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="8aba3-116">**Lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="8aba3-116">**lifetime="long"**</span></span>  |  <span data-ttu-id="8aba3-117">Oui</span><span class="sxs-lookup"><span data-stu-id="8aba3-117">Yes</span></span>  | <span data-ttu-id="8aba3-118">Doit toujours être `long` utilisé pour utiliser un runtime partagé pour le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="8aba3-118">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="8aba3-119">**resid**</span><span class="sxs-lookup"><span data-stu-id="8aba3-119">**resid**</span></span>  |  <span data-ttu-id="8aba3-120">Oui</span><span class="sxs-lookup"><span data-stu-id="8aba3-120">Yes</span></span>  | <span data-ttu-id="8aba3-121">Spécifie l’URL de la page HTML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="8aba3-121">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="8aba3-122">L `resid` 'doit correspondre `id` à un attribut `Url` d’un élément `Resources` dans l’élément.</span><span class="sxs-lookup"><span data-stu-id="8aba3-122">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="8aba3-123">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8aba3-123">See also</span></span>

- [<span data-ttu-id="8aba3-124">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="8aba3-124">Runtimes</span></span>](runtimes.md)
