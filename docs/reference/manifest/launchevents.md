---
title: LaunchEvents dans le fichier manifeste
description: L’élément LaunchEvents configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 16d721ca6d9402d2bd5d19787707e146358044f0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590915"
---
# <a name="launchevents-element"></a><span data-ttu-id="65e4e-103">Élément LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="65e4e-103">LaunchEvents element</span></span>

<span data-ttu-id="65e4e-104">Configure votre add-in pour qu’il s’active en fonction des événements pris en charge.</span><span class="sxs-lookup"><span data-stu-id="65e4e-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="65e4e-105">Enfant de [`<ExtensionPoint>`](extensionpoint.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="65e4e-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="65e4e-106">Pour plus d’informations, [voir Configurer Outlook complément pour l’activation basée sur des événements.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="65e4e-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="65e4e-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="65e4e-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="65e4e-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="65e4e-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="65e4e-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="65e4e-109">Contained in</span></span>

<span data-ttu-id="65e4e-110">[ExtensionPoint](extensionpoint.md) (**launchEvent** mail add-in)</span><span class="sxs-lookup"><span data-stu-id="65e4e-110">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="65e4e-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="65e4e-111">Child elements</span></span>

|  <span data-ttu-id="65e4e-112">Élément</span><span class="sxs-lookup"><span data-stu-id="65e4e-112">Element</span></span> |  <span data-ttu-id="65e4e-113">Requis</span><span class="sxs-lookup"><span data-stu-id="65e4e-113">Required</span></span>  |  <span data-ttu-id="65e4e-114">Description</span><span class="sxs-lookup"><span data-stu-id="65e4e-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="65e4e-115">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="65e4e-115">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="65e4e-116">Oui</span><span class="sxs-lookup"><span data-stu-id="65e4e-116">Yes</span></span> |  <span data-ttu-id="65e4e-117">Mapz l’événement pris en charge à sa fonction dans le fichier JavaScript pour l’activation du complément.</span><span class="sxs-lookup"><span data-stu-id="65e4e-117">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="65e4e-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="65e4e-118">See also</span></span>

- [<span data-ttu-id="65e4e-119">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="65e4e-119">LaunchEvent</span></span>](launchevent.md)
