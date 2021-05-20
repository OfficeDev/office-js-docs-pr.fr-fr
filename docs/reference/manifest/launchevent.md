---
title: LaunchEvent dans le fichier manifeste (aperçu)
description: L’élément LaunchEvent configure votre module d’activation en fonction des événements pris en charge.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 7283e9aba9ca57793019ffe027a7f4d6e3243aa8
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555310"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="a7997-103">LaunchEvent élément (aperçu)</span><span class="sxs-lookup"><span data-stu-id="a7997-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="a7997-104">Configure votre module d’activation en fonction des événements pris en charge.</span><span class="sxs-lookup"><span data-stu-id="a7997-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="a7997-105">Enfant de [`<LaunchEvents>`](launchevents.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="a7997-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="a7997-106">Pour plus d’informations, [consultez Configurez votre Outlook pour l’activation basée sur l’événement.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="a7997-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="a7997-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="a7997-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a7997-108">L’activation basée sur [l’événement est actuellement en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) avant-première et n’est disponible Outlook sur le Web et sur Windows.</span><span class="sxs-lookup"><span data-stu-id="a7997-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and on Windows.</span></span> <span data-ttu-id="a7997-109">Pour plus d’informations, voir [Comment prévisualiser la fonction d’activation basée sur l’événement](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="a7997-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="a7997-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="a7997-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="a7997-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="a7997-111">Contained in</span></span>

- [<span data-ttu-id="a7997-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="a7997-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="a7997-113">Attributs</span><span class="sxs-lookup"><span data-stu-id="a7997-113">Attributes</span></span>

|  <span data-ttu-id="a7997-114">Attribut</span><span class="sxs-lookup"><span data-stu-id="a7997-114">Attribute</span></span>  |  <span data-ttu-id="a7997-115">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a7997-115">Required</span></span>  |  <span data-ttu-id="a7997-116">Description</span><span class="sxs-lookup"><span data-stu-id="a7997-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a7997-117">**Type**</span><span class="sxs-lookup"><span data-stu-id="a7997-117">**Type**</span></span>  |  <span data-ttu-id="a7997-118">Oui</span><span class="sxs-lookup"><span data-stu-id="a7997-118">Yes</span></span>  | <span data-ttu-id="a7997-119">Spécifie un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="a7997-119">Specifies a supported event type.</span></span> <span data-ttu-id="a7997-120">Pour l’ensemble des types pris en charge, voir [Comment prévisualiser la fonction d’activation basée sur l’événement](../../outlook/autolaunch.md#supported-events).</span><span class="sxs-lookup"><span data-stu-id="a7997-120">For the set of supported types, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#supported-events).</span></span> |
|  <span data-ttu-id="a7997-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="a7997-121">**FunctionName**</span></span>  |  <span data-ttu-id="a7997-122">Oui</span><span class="sxs-lookup"><span data-stu-id="a7997-122">Yes</span></span>  | <span data-ttu-id="a7997-123">Spécifie le nom de la fonction JavaScript pour gérer l’événement spécifié dans `Type` l’attribut.</span><span class="sxs-lookup"><span data-stu-id="a7997-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="a7997-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a7997-124">See also</span></span>

- [<span data-ttu-id="a7997-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="a7997-125">LaunchEvents</span></span>](launchevents.md)
