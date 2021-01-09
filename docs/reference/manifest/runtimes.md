---
title: Runtimes dans le fichier manifeste
description: L’élément Runtimes spécifie le runtime de votre add-in.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: afbcc6a909c51d2ed56292ef1541193f7f698d28
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789162"
---
# <a name="runtimes-element"></a><span data-ttu-id="0c5ef-103">Élément Runtimes</span><span class="sxs-lookup"><span data-stu-id="0c5ef-103">Runtimes element</span></span>

<span data-ttu-id="0c5ef-104">Spécifie le runtime de votre add-in.</span><span class="sxs-lookup"><span data-stu-id="0c5ef-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="0c5ef-105">Enfant de [`<Host>`](host.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="0c5ef-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="0c5ef-106">Lorsque vous exécutez Office sur Windows, votre application utilise le navigateur Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="0c5ef-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="0c5ef-107">Dans Excel, cet élément permet au ruban, au volet Des tâches et aux fonctions personnalisées d’utiliser le même runtime.</span><span class="sxs-lookup"><span data-stu-id="0c5ef-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="0c5ef-108">Pour plus d’informations, voir Configurer votre add-in Excel pour utiliser [un runtime JavaScript partagé.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="0c5ef-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="0c5ef-109">Dans Outlook, cet élément active l’activation des compléments basés sur des événements.</span><span class="sxs-lookup"><span data-stu-id="0c5ef-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="0c5ef-110">Pour plus d’informations, voir Configurer votre complément [Outlook pour l’activation basée sur des événements.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="0c5ef-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="0c5ef-111">**Type de add-in :** Volet De tâches, Courrier</span><span class="sxs-lookup"><span data-stu-id="0c5ef-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0c5ef-112">**Outlook**: la fonctionnalité d’activation basée sur des événements est actuellement en [prévisualisation](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) et disponible uniquement dans Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="0c5ef-112">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="0c5ef-113">Pour plus d’informations, [voir Comment afficher un aperçu de la fonctionnalité d’activation basée sur des événements.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="0c5ef-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="0c5ef-114">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="0c5ef-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="0c5ef-115">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="0c5ef-115">Contained in</span></span>

[<span data-ttu-id="0c5ef-116">Host</span><span class="sxs-lookup"><span data-stu-id="0c5ef-116">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="0c5ef-117">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="0c5ef-117">Child elements</span></span>

|  <span data-ttu-id="0c5ef-118">Élément</span><span class="sxs-lookup"><span data-stu-id="0c5ef-118">Element</span></span> |  <span data-ttu-id="0c5ef-119">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="0c5ef-119">Required</span></span>  |  <span data-ttu-id="0c5ef-120">Description</span><span class="sxs-lookup"><span data-stu-id="0c5ef-120">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="0c5ef-121">Runtime</span><span class="sxs-lookup"><span data-stu-id="0c5ef-121">Runtime</span></span>](runtime.md) | <span data-ttu-id="0c5ef-122">Oui</span><span class="sxs-lookup"><span data-stu-id="0c5ef-122">Yes</span></span> |  <span data-ttu-id="0c5ef-123">Runtime de votre add-in.</span><span class="sxs-lookup"><span data-stu-id="0c5ef-123">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="0c5ef-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0c5ef-124">See also</span></span>

- [<span data-ttu-id="0c5ef-125">Runtime</span><span class="sxs-lookup"><span data-stu-id="0c5ef-125">Runtime</span></span>](runtime.md)
