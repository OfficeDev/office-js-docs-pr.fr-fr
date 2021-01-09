---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre add-in pour utiliser un runtime JavaScript partagé pour ses différents composants, par exemple, ruban, volet Des tâches, fonctions personnalisées.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 3cabfacc665ccf6c0e4e796cb0e1fbc70c770ee3
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789183"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="b1506-103">Élément Runtime (aperçu)</span><span class="sxs-lookup"><span data-stu-id="b1506-103">Runtime element (preview)</span></span>

<span data-ttu-id="b1506-104">Configure votre add-in pour utiliser un runtime JavaScript partagé afin que différents composants s’exécutent tous dans le même runtime.</span><span class="sxs-lookup"><span data-stu-id="b1506-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="b1506-105">Enfant de [`<Runtimes>`](runtimes.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="b1506-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="b1506-106">Dans Excel, cet élément permet au ruban, au volet Des tâches et aux fonctions personnalisées d’utiliser le même runtime.</span><span class="sxs-lookup"><span data-stu-id="b1506-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="b1506-107">Pour plus d’informations, voir Configurer votre add-in Excel pour utiliser [un runtime JavaScript partagé.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="b1506-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="b1506-108">Dans Outlook, cet élément active l’activation des compléments basés sur des événements.</span><span class="sxs-lookup"><span data-stu-id="b1506-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="b1506-109">Pour plus d’informations, voir Configurer votre complément [Outlook pour l’activation basée sur des événements.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="b1506-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="b1506-110">**Type de add-in :** Volet De tâches, Courrier</span><span class="sxs-lookup"><span data-stu-id="b1506-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b1506-111">**Outlook**: l’activation basée sur des événements est actuellement [en prévisualisation](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) et disponible uniquement dans Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="b1506-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="b1506-112">Pour plus d’informations, [voir Comment afficher un aperçu de la fonctionnalité d’activation basée sur des événements.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="b1506-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="b1506-113">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b1506-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="b1506-114">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b1506-114">Contained in</span></span>

- [<span data-ttu-id="b1506-115">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="b1506-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="b1506-116">Attributs</span><span class="sxs-lookup"><span data-stu-id="b1506-116">Attributes</span></span>

|  <span data-ttu-id="b1506-117">Attribut</span><span class="sxs-lookup"><span data-stu-id="b1506-117">Attribute</span></span>  |  <span data-ttu-id="b1506-118">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="b1506-118">Required</span></span>  |  <span data-ttu-id="b1506-119">Description</span><span class="sxs-lookup"><span data-stu-id="b1506-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b1506-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="b1506-120">**resid**</span></span>  |  <span data-ttu-id="b1506-121">Oui</span><span class="sxs-lookup"><span data-stu-id="b1506-121">Yes</span></span>  | <span data-ttu-id="b1506-122">Spécifie l’emplacement URL de la page HTML de votre application.</span><span class="sxs-lookup"><span data-stu-id="b1506-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="b1506-123">Il ne peut pas y avoir plus de 32 caractères et doit correspondre à un `resid` `id` attribut `Url` d’un élément dans `Resources` l’élément.</span><span class="sxs-lookup"><span data-stu-id="b1506-123">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="b1506-124">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="b1506-124">**lifetime**</span></span>  |  <span data-ttu-id="b1506-125">Non</span><span class="sxs-lookup"><span data-stu-id="b1506-125">No</span></span>  | <span data-ttu-id="b1506-126">La valeur par `lifetime` défaut est et n’a pas besoin `short` d’être spécifiée.</span><span class="sxs-lookup"><span data-stu-id="b1506-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="b1506-127">Les add-ins Outlook utilisent uniquement la `short` valeur.</span><span class="sxs-lookup"><span data-stu-id="b1506-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="b1506-128">Si vous souhaitez utiliser un runtime partagé dans un add-in Excel, définissez explicitement la valeur sur `long` .</span><span class="sxs-lookup"><span data-stu-id="b1506-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="b1506-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b1506-129">See also</span></span>

- [<span data-ttu-id="b1506-130">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="b1506-130">Runtimes</span></span>](runtimes.md)
