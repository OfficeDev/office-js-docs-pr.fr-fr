---
title: Runtimes dans le fichier manifeste
description: L’élément runtimes spécifie le runtime de votre complément.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: a8598a8f926e6d6905c147f5c554f1d40a692ad9
ms.sourcegitcommit: 09a8683ff29cf06d0d1d822be83cf0798f1ccdf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/01/2020
ms.locfileid: "44471323"
---
# <a name="runtimes-element"></a><span data-ttu-id="7f5ee-103">Élément runtimes</span><span class="sxs-lookup"><span data-stu-id="7f5ee-103">Runtimes element</span></span>

<span data-ttu-id="7f5ee-104">Spécifie le runtime de votre complément.</span><span class="sxs-lookup"><span data-stu-id="7f5ee-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="7f5ee-105">Enfant de l' [`<Host>`](host.md) élément.</span><span class="sxs-lookup"><span data-stu-id="7f5ee-105">Child of the [`<Host>`](host.md) element.</span></span> <span data-ttu-id="7f5ee-106">Si l' `Runtimes` élément est présent dans votre manifeste, votre complément utilise par défaut le navigateur Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="7f5ee-106">If the `Runtimes` element is present in your manifest, your add-in will by default use the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="7f5ee-107">Dans Excel, cet élément active le ruban, le volet des tâches et les fonctions personnalisées pour utiliser le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="7f5ee-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="7f5ee-108">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="7f5ee-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="7f5ee-109">Dans Outlook, cet élément active l’activation de complément basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="7f5ee-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="7f5ee-110">Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="7f5ee-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="7f5ee-111">**Type de complément :** Volet Office, messagerie</span><span class="sxs-lookup"><span data-stu-id="7f5ee-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7f5ee-112">**Excel**: le runtime partagé est actuellement disponible uniquement dans Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="7f5ee-112">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="7f5ee-113">**Outlook**: la fonctionnalité d’activation basée sur un événement est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="7f5ee-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="7f5ee-114">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="7f5ee-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="7f5ee-115">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="7f5ee-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="7f5ee-116">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="7f5ee-116">Contained in</span></span>

[<span data-ttu-id="7f5ee-117">Hôte</span><span class="sxs-lookup"><span data-stu-id="7f5ee-117">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="7f5ee-118">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="7f5ee-118">Child elements</span></span>

|  <span data-ttu-id="7f5ee-119">Élément</span><span class="sxs-lookup"><span data-stu-id="7f5ee-119">Element</span></span> |  <span data-ttu-id="7f5ee-120">Requis</span><span class="sxs-lookup"><span data-stu-id="7f5ee-120">Required</span></span>  |  <span data-ttu-id="7f5ee-121">Description</span><span class="sxs-lookup"><span data-stu-id="7f5ee-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="7f5ee-122">Runtime</span><span class="sxs-lookup"><span data-stu-id="7f5ee-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="7f5ee-123">Oui</span><span class="sxs-lookup"><span data-stu-id="7f5ee-123">Yes</span></span> |  <span data-ttu-id="7f5ee-124">Le runtime de votre complément.</span><span class="sxs-lookup"><span data-stu-id="7f5ee-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="7f5ee-125">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7f5ee-125">See also</span></span>

- [<span data-ttu-id="7f5ee-126">Runtime</span><span class="sxs-lookup"><span data-stu-id="7f5ee-126">Runtime</span></span>](runtime.md)
