---
title: LaunchEvent dans le fichier manifeste
description: L’élément LaunchEvent configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: c866a085ed6b7a33c8d7bf02d25e6ec748629e07
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591078"
---
# <a name="launchevent-element"></a><span data-ttu-id="9290a-103">Élément LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="9290a-103">LaunchEvent element</span></span>

<span data-ttu-id="9290a-104">Configure votre add-in pour qu’il s’active en fonction des événements pris en charge.</span><span class="sxs-lookup"><span data-stu-id="9290a-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="9290a-105">Enfant de [`<LaunchEvents>`](launchevents.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="9290a-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="9290a-106">Pour plus d’informations, [voir Configurer Outlook complément pour l’activation basée sur des événements.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="9290a-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="9290a-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="9290a-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9290a-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="9290a-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="9290a-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="9290a-109">Contained in</span></span>

- [<span data-ttu-id="9290a-110">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="9290a-110">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="9290a-111">Attributs</span><span class="sxs-lookup"><span data-stu-id="9290a-111">Attributes</span></span>

|  <span data-ttu-id="9290a-112">Attribut</span><span class="sxs-lookup"><span data-stu-id="9290a-112">Attribute</span></span>  |  <span data-ttu-id="9290a-113">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="9290a-113">Required</span></span>  |  <span data-ttu-id="9290a-114">Description</span><span class="sxs-lookup"><span data-stu-id="9290a-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="9290a-115">**Type**</span><span class="sxs-lookup"><span data-stu-id="9290a-115">**Type**</span></span>  |  <span data-ttu-id="9290a-116">Oui</span><span class="sxs-lookup"><span data-stu-id="9290a-116">Yes</span></span>  | <span data-ttu-id="9290a-117">Spécifie un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="9290a-117">Specifies a supported event type.</span></span> <span data-ttu-id="9290a-118">Pour obtenir l’ensemble des types pris en charge, voir [Configurer Outlook complément pour l’activation basée sur des événements.](../../outlook/autolaunch.md#supported-events)</span><span class="sxs-lookup"><span data-stu-id="9290a-118">For the set of supported types, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events).</span></span> |
|  <span data-ttu-id="9290a-119">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="9290a-119">**FunctionName**</span></span>  |  <span data-ttu-id="9290a-120">Oui</span><span class="sxs-lookup"><span data-stu-id="9290a-120">Yes</span></span>  | <span data-ttu-id="9290a-121">Spécifie le nom de la fonction JavaScript pour gérer l’événement spécifié dans `Type` l’attribut.</span><span class="sxs-lookup"><span data-stu-id="9290a-121">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="9290a-122">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9290a-122">See also</span></span>

- [<span data-ttu-id="9290a-123">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="9290a-123">LaunchEvents</span></span>](launchevents.md)
