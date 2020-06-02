---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé pour ses différents composants, par exemple le ruban, le volet Office, les fonctions personnalisées.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: a463b72f22b41f74e2fe98acca467762bb00cf39
ms.sourcegitcommit: 09a8683ff29cf06d0d1d822be83cf0798f1ccdf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/01/2020
ms.locfileid: "44471337"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="80275-103">Élément Runtime (aperçu)</span><span class="sxs-lookup"><span data-stu-id="80275-103">Runtime element (preview)</span></span>

<span data-ttu-id="80275-104">Configure votre complément pour qu’il utilise un Runtime JavaScript partagé afin que les différents composants s’exécutent tous dans le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="80275-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="80275-105">Enfant de l' [`<Runtimes>`](runtimes.md) élément.</span><span class="sxs-lookup"><span data-stu-id="80275-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="80275-106">Dans Excel, cet élément active le ruban, le volet des tâches et les fonctions personnalisées pour utiliser le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="80275-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="80275-107">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="80275-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="80275-108">Dans Outlook, cet élément active l’activation de complément basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="80275-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="80275-109">Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="80275-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="80275-110">**Type de complément :** Volet Office, messagerie</span><span class="sxs-lookup"><span data-stu-id="80275-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="80275-111">**Excel**: le runtime partagé est actuellement disponible uniquement dans Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="80275-111">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="80275-112">**Outlook**: l’activation basée sur un événement est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="80275-112">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="80275-113">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="80275-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="80275-114">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="80275-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="80275-115">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="80275-115">Contained in</span></span>

- [<span data-ttu-id="80275-116">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="80275-116">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="80275-117">Attributs</span><span class="sxs-lookup"><span data-stu-id="80275-117">Attributes</span></span>

|  <span data-ttu-id="80275-118">Attribut</span><span class="sxs-lookup"><span data-stu-id="80275-118">Attribute</span></span>  |  <span data-ttu-id="80275-119">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="80275-119">Required</span></span>  |  <span data-ttu-id="80275-120">Description</span><span class="sxs-lookup"><span data-stu-id="80275-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="80275-121">**resid**</span><span class="sxs-lookup"><span data-stu-id="80275-121">**resid**</span></span>  |  <span data-ttu-id="80275-122">Oui</span><span class="sxs-lookup"><span data-stu-id="80275-122">Yes</span></span>  | <span data-ttu-id="80275-123">Spécifie l’URL de la page HTML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="80275-123">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="80275-124">L' `resid` doit correspondre à un `id` attribut d’un `Url` élément dans l' `Resources` élément.</span><span class="sxs-lookup"><span data-stu-id="80275-124">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="80275-125">**vie**</span><span class="sxs-lookup"><span data-stu-id="80275-125">**lifetime**</span></span>  |  <span data-ttu-id="80275-126">Non</span><span class="sxs-lookup"><span data-stu-id="80275-126">No</span></span>  | <span data-ttu-id="80275-127">La valeur par défaut de `lifetime` est `short` et n’a pas besoin d’être spécifiée.</span><span class="sxs-lookup"><span data-stu-id="80275-127">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="80275-128">Les compléments Outlook utilisent uniquement la `short` valeur.</span><span class="sxs-lookup"><span data-stu-id="80275-128">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="80275-129">Si vous souhaitez utiliser un runtime partagé dans un complément Excel, définissez explicitement la valeur sur `long` .</span><span class="sxs-lookup"><span data-stu-id="80275-129">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="80275-130">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="80275-130">See also</span></span>

- [<span data-ttu-id="80275-131">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="80275-131">Runtimes</span></span>](runtimes.md)
