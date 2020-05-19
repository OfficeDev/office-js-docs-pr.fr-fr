---
title: LaunchEvent dans le fichier manifeste (aperçu)
description: L’élément LaunchEvent configure votre complément de sorte qu’il s’active en fonction des événements pris en charge.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: a4f5208ec7f735d926c3a878cae34973c3992cf9
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278542"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="22608-103">Élément LaunchEvent (aperçu)</span><span class="sxs-lookup"><span data-stu-id="22608-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="22608-104">Configure votre complément pour qu’il s’active en fonction des événements pris en charge.</span><span class="sxs-lookup"><span data-stu-id="22608-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="22608-105">Enfant de l' [`<LaunchEvents>`](launchevents.md) élément.</span><span class="sxs-lookup"><span data-stu-id="22608-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="22608-106">Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="22608-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="22608-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="22608-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="22608-108">L’activation basée sur les événements est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="22608-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="22608-109">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="22608-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="22608-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="22608-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="22608-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="22608-111">Contained in</span></span>

- [<span data-ttu-id="22608-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="22608-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="22608-113">Attributs</span><span class="sxs-lookup"><span data-stu-id="22608-113">Attributes</span></span>

|  <span data-ttu-id="22608-114">Attribut</span><span class="sxs-lookup"><span data-stu-id="22608-114">Attribute</span></span>  |  <span data-ttu-id="22608-115">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="22608-115">Required</span></span>  |  <span data-ttu-id="22608-116">Description</span><span class="sxs-lookup"><span data-stu-id="22608-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="22608-117">**Type**</span><span class="sxs-lookup"><span data-stu-id="22608-117">**Type**</span></span>  |  <span data-ttu-id="22608-118">Oui</span><span class="sxs-lookup"><span data-stu-id="22608-118">Yes</span></span>  | <span data-ttu-id="22608-119">Spécifie un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="22608-119">Specifies a supported event type.</span></span> <span data-ttu-id="22608-120">Les types disponibles sont `OnNewMessageCompose` et `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="22608-120">Available types are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> |
|  <span data-ttu-id="22608-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="22608-121">**FunctionName**</span></span>  |  <span data-ttu-id="22608-122">Oui</span><span class="sxs-lookup"><span data-stu-id="22608-122">Yes</span></span>  | <span data-ttu-id="22608-123">Spécifie le nom de la fonction JavaScript permettant de gérer l’événement spécifié dans l' `Type` attribut.</span><span class="sxs-lookup"><span data-stu-id="22608-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="22608-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="22608-124">See also</span></span>

- [<span data-ttu-id="22608-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="22608-125">LaunchEvents</span></span>](launchevents.md)
