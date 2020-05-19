---
title: LaunchEvents dans le fichier manifeste (aperçu)
description: L’élément LaunchEvents configure votre complément de sorte qu’il s’active en fonction des événements pris en charge.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 2e1ad56d405fca0f85fad500a113fba7d0448caf
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278541"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="6c9d5-103">Élément LaunchEvents (aperçu)</span><span class="sxs-lookup"><span data-stu-id="6c9d5-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="6c9d5-104">Configure votre complément pour qu’il s’active en fonction des événements pris en charge.</span><span class="sxs-lookup"><span data-stu-id="6c9d5-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="6c9d5-105">Enfant de l' [`<ExtensionPoint>`](extensionpoint.md) élément.</span><span class="sxs-lookup"><span data-stu-id="6c9d5-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="6c9d5-106">Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="6c9d5-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="6c9d5-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="6c9d5-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6c9d5-108">L’activation basée sur les événements est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="6c9d5-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="6c9d5-109">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="6c9d5-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="6c9d5-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="6c9d5-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="6c9d5-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="6c9d5-111">Contained in</span></span>

<span data-ttu-id="6c9d5-112">[ExtensionPoint](extensionpoint.md) (complément de messagerie**LaunchEvent** )</span><span class="sxs-lookup"><span data-stu-id="6c9d5-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="6c9d5-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="6c9d5-113">Child elements</span></span>

|  <span data-ttu-id="6c9d5-114">Élément</span><span class="sxs-lookup"><span data-stu-id="6c9d5-114">Element</span></span> |  <span data-ttu-id="6c9d5-115">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="6c9d5-115">Required</span></span>  |  <span data-ttu-id="6c9d5-116">Description</span><span class="sxs-lookup"><span data-stu-id="6c9d5-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="6c9d5-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="6c9d5-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="6c9d5-118">Oui</span><span class="sxs-lookup"><span data-stu-id="6c9d5-118">Yes</span></span> |  <span data-ttu-id="6c9d5-119">Mappez l’événement pris en charge à sa fonction dans le fichier JavaScript pour l’activation des compléments.</span><span class="sxs-lookup"><span data-stu-id="6c9d5-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="6c9d5-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6c9d5-120">See also</span></span>

- [<span data-ttu-id="6c9d5-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="6c9d5-121">LaunchEvent</span></span>](launchevent.md)
