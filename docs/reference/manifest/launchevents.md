---
title: LaunchEvents dans le fichier manifeste (aperçu)
description: L’élément LaunchEvents configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 59c52aa3f60e69e2bdda84718c6123f02942fedc
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237979"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="626c2-103">Élément LaunchEvents (aperçu)</span><span class="sxs-lookup"><span data-stu-id="626c2-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="626c2-104">Configure votre add-in pour qu’il s’active en fonction des événements pris en charge.</span><span class="sxs-lookup"><span data-stu-id="626c2-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="626c2-105">Enfant de [`<ExtensionPoint>`](extensionpoint.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="626c2-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="626c2-106">Pour plus d’informations, voir Configurer votre complément [Outlook pour l’activation basée sur des événements.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="626c2-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="626c2-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="626c2-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="626c2-108">L’activation basée sur des événements est actuellement [en prévisualisation](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) et disponible uniquement dans Outlook sur le web et sur Windows.</span><span class="sxs-lookup"><span data-stu-id="626c2-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and on Windows.</span></span> <span data-ttu-id="626c2-109">Pour plus d’informations, [voir Comment afficher un aperçu de la fonctionnalité d’activation basée sur des événements.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="626c2-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="626c2-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="626c2-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="626c2-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="626c2-111">Contained in</span></span>

<span data-ttu-id="626c2-112">[ExtensionPoint](extensionpoint.md) (**launchEvent** mail add-in)</span><span class="sxs-lookup"><span data-stu-id="626c2-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="626c2-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="626c2-113">Child elements</span></span>

|  <span data-ttu-id="626c2-114">Élément</span><span class="sxs-lookup"><span data-stu-id="626c2-114">Element</span></span> |  <span data-ttu-id="626c2-115">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="626c2-115">Required</span></span>  |  <span data-ttu-id="626c2-116">Description</span><span class="sxs-lookup"><span data-stu-id="626c2-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="626c2-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="626c2-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="626c2-118">Oui</span><span class="sxs-lookup"><span data-stu-id="626c2-118">Yes</span></span> |  <span data-ttu-id="626c2-119">Mapz l’événement pris en charge à sa fonction dans le fichier JavaScript pour l’activation du complément.</span><span class="sxs-lookup"><span data-stu-id="626c2-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="626c2-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="626c2-120">See also</span></span>

- [<span data-ttu-id="626c2-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="626c2-121">LaunchEvent</span></span>](launchevent.md)
