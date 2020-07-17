---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé pour ses différents composants, par exemple le ruban, le volet Office, les fonctions personnalisées.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 9e6e13f83db363fb5485c8d8defbc381c80e32d6
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159366"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="f0921-103">Élément Runtime (aperçu)</span><span class="sxs-lookup"><span data-stu-id="f0921-103">Runtime element (preview)</span></span>

<span data-ttu-id="f0921-104">Configure votre complément pour qu’il utilise un Runtime JavaScript partagé afin que les différents composants s’exécutent tous dans le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="f0921-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="f0921-105">Enfant de l' [`<Runtimes>`](runtimes.md) élément.</span><span class="sxs-lookup"><span data-stu-id="f0921-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="f0921-106">Dans Excel, cet élément active le ruban, le volet des tâches et les fonctions personnalisées pour utiliser le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="f0921-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="f0921-107">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="f0921-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="f0921-108">Dans Outlook, cet élément active l’activation de complément basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="f0921-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="f0921-109">Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="f0921-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="f0921-110">**Type de complément :** Volet Office, messagerie</span><span class="sxs-lookup"><span data-stu-id="f0921-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f0921-111">**Outlook**: l’activation basée sur un événement est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="f0921-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="f0921-112">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="f0921-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="f0921-113">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="f0921-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="f0921-114">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="f0921-114">Contained in</span></span>

- [<span data-ttu-id="f0921-115">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="f0921-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="f0921-116">Attributs</span><span class="sxs-lookup"><span data-stu-id="f0921-116">Attributes</span></span>

|  <span data-ttu-id="f0921-117">Attribut</span><span class="sxs-lookup"><span data-stu-id="f0921-117">Attribute</span></span>  |  <span data-ttu-id="f0921-118">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="f0921-118">Required</span></span>  |  <span data-ttu-id="f0921-119">Description</span><span class="sxs-lookup"><span data-stu-id="f0921-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f0921-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="f0921-120">**resid**</span></span>  |  <span data-ttu-id="f0921-121">Oui</span><span class="sxs-lookup"><span data-stu-id="f0921-121">Yes</span></span>  | <span data-ttu-id="f0921-122">Spécifie l’URL de la page HTML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="f0921-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="f0921-123">L' `resid` doit correspondre à un `id` attribut d’un `Url` élément dans l' `Resources` élément.</span><span class="sxs-lookup"><span data-stu-id="f0921-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="f0921-124">**vie**</span><span class="sxs-lookup"><span data-stu-id="f0921-124">**lifetime**</span></span>  |  <span data-ttu-id="f0921-125">Non</span><span class="sxs-lookup"><span data-stu-id="f0921-125">No</span></span>  | <span data-ttu-id="f0921-126">La valeur par défaut de `lifetime` est `short` et n’a pas besoin d’être spécifiée.</span><span class="sxs-lookup"><span data-stu-id="f0921-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="f0921-127">Les compléments Outlook utilisent uniquement la `short` valeur.</span><span class="sxs-lookup"><span data-stu-id="f0921-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="f0921-128">Si vous souhaitez utiliser un runtime partagé dans un complément Excel, définissez explicitement la valeur sur `long` .</span><span class="sxs-lookup"><span data-stu-id="f0921-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="f0921-129">Consultez également</span><span class="sxs-lookup"><span data-stu-id="f0921-129">See also</span></span>

- [<span data-ttu-id="f0921-130">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="f0921-130">Runtimes</span></span>](runtimes.md)
