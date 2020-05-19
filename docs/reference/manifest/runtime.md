---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé pour ses différents composants, par exemple le ruban, le volet Office, les fonctions personnalisées.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: c2c404bcaad6e24af58f5c0ed8835343abb97e5f
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278412"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="d6305-103">Élément Runtime (aperçu)</span><span class="sxs-lookup"><span data-stu-id="d6305-103">Runtime element (preview)</span></span>

<span data-ttu-id="d6305-104">Configure votre complément pour qu’il utilise un Runtime JavaScript partagé afin que les différents composants s’exécutent tous dans le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="d6305-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="d6305-105">Enfant de l' [`<Runtimes>`](runtimes.md) élément.</span><span class="sxs-lookup"><span data-stu-id="d6305-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="d6305-106">Dans Excel, cet élément active le ruban, le volet des tâches et les fonctions personnalisées pour utiliser le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="d6305-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="d6305-107">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="d6305-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="d6305-108">Dans Outlook, cet élément active l’activation de complément basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="d6305-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="d6305-109">Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="d6305-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="d6305-110">**Type de complément :** Volet Office, messagerie</span><span class="sxs-lookup"><span data-stu-id="d6305-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d6305-111">**Excel**: le runtime partagé est actuellement en préversion et disponible uniquement dans Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="d6305-111">**Excel**: Shared runtime is currently in preview and only available in Excel on Windows.</span></span> <span data-ttu-id="d6305-112">Pour essayer les fonctionnalités d’aperçu, vous devrez rejoindre [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="d6305-112">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>
>
> <span data-ttu-id="d6305-113">**Outlook**: l’activation basée sur un événement est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="d6305-113">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="d6305-114">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="d6305-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="d6305-115">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="d6305-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="d6305-116">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="d6305-116">Contained in</span></span>

- [<span data-ttu-id="d6305-117">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="d6305-117">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="d6305-118">Attributs</span><span class="sxs-lookup"><span data-stu-id="d6305-118">Attributes</span></span>

|  <span data-ttu-id="d6305-119">Attribut</span><span class="sxs-lookup"><span data-stu-id="d6305-119">Attribute</span></span>  |  <span data-ttu-id="d6305-120">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="d6305-120">Required</span></span>  |  <span data-ttu-id="d6305-121">Description</span><span class="sxs-lookup"><span data-stu-id="d6305-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d6305-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="d6305-122">**resid**</span></span>  |  <span data-ttu-id="d6305-123">Oui</span><span class="sxs-lookup"><span data-stu-id="d6305-123">Yes</span></span>  | <span data-ttu-id="d6305-124">Spécifie l’URL de la page HTML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="d6305-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="d6305-125">L' `resid` doit correspondre à un `id` attribut d’un `Url` élément dans l' `Resources` élément.</span><span class="sxs-lookup"><span data-stu-id="d6305-125">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="d6305-126">**vie**</span><span class="sxs-lookup"><span data-stu-id="d6305-126">**lifetime**</span></span>  |  <span data-ttu-id="d6305-127">Non</span><span class="sxs-lookup"><span data-stu-id="d6305-127">No</span></span>  | <span data-ttu-id="d6305-128">La valeur par défaut de `lifetime` est `short` et n’a pas besoin d’être spécifiée.</span><span class="sxs-lookup"><span data-stu-id="d6305-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="d6305-129">Les compléments Outlook utilisent uniquement la `short` valeur.</span><span class="sxs-lookup"><span data-stu-id="d6305-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="d6305-130">Si vous souhaitez utiliser un runtime partagé dans un complément Excel, définissez explicitement la valeur sur `long` .</span><span class="sxs-lookup"><span data-stu-id="d6305-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="d6305-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d6305-131">See also</span></span>

- [<span data-ttu-id="d6305-132">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="d6305-132">Runtimes</span></span>](runtimes.md)
