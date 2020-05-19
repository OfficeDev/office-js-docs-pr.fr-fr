---
title: Runtimes dans le fichier manifeste
description: L’élément runtimes spécifie le runtime de votre complément.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 22156a171ca2f423024efb1b3d2a6fdae07dfef6
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278363"
---
# <a name="runtimes-element"></a><span data-ttu-id="9b728-103">Élément runtimes</span><span class="sxs-lookup"><span data-stu-id="9b728-103">Runtimes element</span></span>

<span data-ttu-id="9b728-104">Spécifie le runtime de votre complément.</span><span class="sxs-lookup"><span data-stu-id="9b728-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="9b728-105">Enfant de l' [`<Host>`](host.md) élément.</span><span class="sxs-lookup"><span data-stu-id="9b728-105">Child of the [`<Host>`](host.md) element.</span></span>

<span data-ttu-id="9b728-106">Dans Excel, cet élément active le ruban, le volet des tâches et les fonctions personnalisées pour utiliser le même Runtime.</span><span class="sxs-lookup"><span data-stu-id="9b728-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="9b728-107">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="9b728-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="9b728-108">Dans Outlook, cet élément active l’activation de complément basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="9b728-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="9b728-109">Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="9b728-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="9b728-110">**Type de complément :** Volet Office, messagerie</span><span class="sxs-lookup"><span data-stu-id="9b728-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9b728-111">**Excel**: le runtime partagé est actuellement en préversion et disponible uniquement dans Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="9b728-111">**Excel**: Shared runtime is currently in preview and only available in Excel on Windows.</span></span> <span data-ttu-id="9b728-112">Pour essayer les fonctionnalités d’aperçu, vous devrez rejoindre [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="9b728-112">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>
>
> <span data-ttu-id="9b728-113">**Outlook**: la fonctionnalité d’activation basée sur un événement est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="9b728-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="9b728-114">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="9b728-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="9b728-115">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="9b728-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="9b728-116">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="9b728-116">Contained in</span></span>

[<span data-ttu-id="9b728-117">Hôte</span><span class="sxs-lookup"><span data-stu-id="9b728-117">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="9b728-118">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9b728-118">Child elements</span></span>

|  <span data-ttu-id="9b728-119">Élément</span><span class="sxs-lookup"><span data-stu-id="9b728-119">Element</span></span> |  <span data-ttu-id="9b728-120">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="9b728-120">Required</span></span>  |  <span data-ttu-id="9b728-121">Description</span><span class="sxs-lookup"><span data-stu-id="9b728-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="9b728-122">Runtime</span><span class="sxs-lookup"><span data-stu-id="9b728-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="9b728-123">Oui</span><span class="sxs-lookup"><span data-stu-id="9b728-123">Yes</span></span> |  <span data-ttu-id="9b728-124">Le runtime de votre complément.</span><span class="sxs-lookup"><span data-stu-id="9b728-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="9b728-125">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9b728-125">See also</span></span>

- [<span data-ttu-id="9b728-126">Runtime</span><span class="sxs-lookup"><span data-stu-id="9b728-126">Runtime</span></span>](runtime.md)
