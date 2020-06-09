---
title: LaunchEvents dans le fichier manifeste (aperçu)
description: L’élément LaunchEvents configure votre complément de sorte qu’il s’active en fonction des événements pris en charge.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 92416f8c646326410a8cd9ee7831e17a5c5f1ffc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611770"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="cd07b-103">Élément LaunchEvents (aperçu)</span><span class="sxs-lookup"><span data-stu-id="cd07b-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="cd07b-104">Configure votre complément pour qu’il s’active en fonction des événements pris en charge.</span><span class="sxs-lookup"><span data-stu-id="cd07b-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="cd07b-105">Enfant de l' [`<ExtensionPoint>`](extensionpoint.md) élément.</span><span class="sxs-lookup"><span data-stu-id="cd07b-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="cd07b-106">Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="cd07b-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="cd07b-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="cd07b-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cd07b-108">L’activation basée sur les événements est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="cd07b-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="cd07b-109">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="cd07b-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="cd07b-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="cd07b-110">Syntax</span></span>

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a><span data-ttu-id="cd07b-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="cd07b-111">Contained in</span></span>

<span data-ttu-id="cd07b-112">[ExtensionPoint](extensionpoint.md) (complément de messagerie**LaunchEvent** )</span><span class="sxs-lookup"><span data-stu-id="cd07b-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="cd07b-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="cd07b-113">Child elements</span></span>

|  <span data-ttu-id="cd07b-114">Élément</span><span class="sxs-lookup"><span data-stu-id="cd07b-114">Element</span></span> |  <span data-ttu-id="cd07b-115">Requis</span><span class="sxs-lookup"><span data-stu-id="cd07b-115">Required</span></span>  |  <span data-ttu-id="cd07b-116">Description</span><span class="sxs-lookup"><span data-stu-id="cd07b-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="cd07b-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="cd07b-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="cd07b-118">Oui</span><span class="sxs-lookup"><span data-stu-id="cd07b-118">Yes</span></span> |  <span data-ttu-id="cd07b-119">Mappez l’événement pris en charge à sa fonction dans le fichier JavaScript pour l’activation des compléments.</span><span class="sxs-lookup"><span data-stu-id="cd07b-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="cd07b-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cd07b-120">See also</span></span>

- [<span data-ttu-id="cd07b-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="cd07b-121">LaunchEvent</span></span>](launchevent.md)
