---
title: Runtimes dans le fichier manifeste
description: L’élément runtimes spécifie le runtime de votre complément.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 95549d88df24a7d7c54cf27c92c15693491bdf29
ms.sourcegitcommit: 9229102c16a1864e3a8724aaf9b0dc68b1428094
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/03/2020
ms.locfileid: "44520338"
---
# <a name="runtimes-element"></a><span data-ttu-id="c5e59-103">Élément runtimes</span><span class="sxs-lookup"><span data-stu-id="c5e59-103">Runtimes element</span></span>

<span data-ttu-id="c5e59-104">Spécifie le runtime de votre complément.</span><span class="sxs-lookup"><span data-stu-id="c5e59-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="c5e59-105">Enfant de l' [`<Host>`](host.md) élément.</span><span class="sxs-lookup"><span data-stu-id="c5e59-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="c5e59-106">Lors de l’exécution dans Office sur Windows, votre complément utilise le navigateur Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="c5e59-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="c5e59-107">Dans Excel, cet élément active le ruban, le volet des tâches et les fonctions personnalisées pour utiliser le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="c5e59-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="c5e59-108">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="c5e59-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="c5e59-109">Dans Outlook, cet élément active l’activation de complément basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="c5e59-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="c5e59-110">Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="c5e59-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="c5e59-111">**Type de complément :** Volet Office, messagerie</span><span class="sxs-lookup"><span data-stu-id="c5e59-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c5e59-112">**Excel**: le runtime partagé est actuellement disponible uniquement dans Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="c5e59-112">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="c5e59-113">**Outlook**: la fonctionnalité d’activation basée sur un événement est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="c5e59-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="c5e59-114">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="c5e59-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="c5e59-115">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="c5e59-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="c5e59-116">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="c5e59-116">Contained in</span></span>

[<span data-ttu-id="c5e59-117">Hôte</span><span class="sxs-lookup"><span data-stu-id="c5e59-117">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="c5e59-118">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="c5e59-118">Child elements</span></span>

|  <span data-ttu-id="c5e59-119">Élément</span><span class="sxs-lookup"><span data-stu-id="c5e59-119">Element</span></span> |  <span data-ttu-id="c5e59-120">Requis</span><span class="sxs-lookup"><span data-stu-id="c5e59-120">Required</span></span>  |  <span data-ttu-id="c5e59-121">Description</span><span class="sxs-lookup"><span data-stu-id="c5e59-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="c5e59-122">Runtime</span><span class="sxs-lookup"><span data-stu-id="c5e59-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="c5e59-123">Oui</span><span class="sxs-lookup"><span data-stu-id="c5e59-123">Yes</span></span> |  <span data-ttu-id="c5e59-124">Le runtime de votre complément.</span><span class="sxs-lookup"><span data-stu-id="c5e59-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="c5e59-125">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c5e59-125">See also</span></span>

- [<span data-ttu-id="c5e59-126">Runtime</span><span class="sxs-lookup"><span data-stu-id="c5e59-126">Runtime</span></span>](runtime.md)
