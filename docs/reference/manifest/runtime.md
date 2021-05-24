---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre add-in pour utiliser un runtime JavaScript partagé pour ses différents composants, par exemple, ruban, volet des tâches, fonctions personnalisées.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590911"
---
# <a name="runtime-element"></a><span data-ttu-id="e64c3-103">Élément Runtime</span><span class="sxs-lookup"><span data-stu-id="e64c3-103">Runtime element</span></span>

<span data-ttu-id="e64c3-104">Configure votre add-in pour utiliser un runtime JavaScript partagé afin que différents composants s’exécutent tous dans le même runtime.</span><span class="sxs-lookup"><span data-stu-id="e64c3-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="e64c3-105">Enfant de [`<Runtimes>`](runtimes.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="e64c3-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="e64c3-106">**Type de add-in :** Volet De tâches, Courrier</span><span class="sxs-lookup"><span data-stu-id="e64c3-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="e64c3-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="e64c3-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="e64c3-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="e64c3-108">Contained in</span></span>

- [<span data-ttu-id="e64c3-109">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="e64c3-109">Runtimes</span></span>](runtimes.md)

## <a name="child-elements"></a><span data-ttu-id="e64c3-110">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="e64c3-110">Child elements</span></span>

|  <span data-ttu-id="e64c3-111">Élément</span><span class="sxs-lookup"><span data-stu-id="e64c3-111">Element</span></span> |  <span data-ttu-id="e64c3-112">Requis</span><span class="sxs-lookup"><span data-stu-id="e64c3-112">Required</span></span>  |  <span data-ttu-id="e64c3-113">Description</span><span class="sxs-lookup"><span data-stu-id="e64c3-113">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="e64c3-114">Override</span><span class="sxs-lookup"><span data-stu-id="e64c3-114">Override</span></span>](override.md) | <span data-ttu-id="e64c3-115">Non</span><span class="sxs-lookup"><span data-stu-id="e64c3-115">No</span></span> | <span data-ttu-id="e64c3-116">**Outlook**: spécifie l’emplacement d’URL du fichier JavaScript dont Outlook Desktop a besoin pour les handleurs de [point d’extension LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent)</span><span class="sxs-lookup"><span data-stu-id="e64c3-116">**Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span> <span data-ttu-id="e64c3-117">**Important**: Pour le moment, vous ne pouvez définir qu’un seul élément et `<Override>` il doit être de type `javascript` .</span><span class="sxs-lookup"><span data-stu-id="e64c3-117">**Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.</span></span>|

## <a name="attributes"></a><span data-ttu-id="e64c3-118">Attributs</span><span class="sxs-lookup"><span data-stu-id="e64c3-118">Attributes</span></span>

|  <span data-ttu-id="e64c3-119">Attribut</span><span class="sxs-lookup"><span data-stu-id="e64c3-119">Attribute</span></span>  |  <span data-ttu-id="e64c3-120">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="e64c3-120">Required</span></span>  |  <span data-ttu-id="e64c3-121">Description</span><span class="sxs-lookup"><span data-stu-id="e64c3-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e64c3-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="e64c3-122">**resid**</span></span>  |  <span data-ttu-id="e64c3-123">Oui</span><span class="sxs-lookup"><span data-stu-id="e64c3-123">Yes</span></span>  | <span data-ttu-id="e64c3-124">Spécifie l’emplacement URL de la page HTML de votre application.</span><span class="sxs-lookup"><span data-stu-id="e64c3-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="e64c3-125">Il ne peut pas y avoir plus de 32 caractères et doit correspondre à un `resid` `id` attribut `Url` d’un élément dans `Resources` l’élément.</span><span class="sxs-lookup"><span data-stu-id="e64c3-125">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="e64c3-126">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="e64c3-126">**lifetime**</span></span>  |  <span data-ttu-id="e64c3-127">Non</span><span class="sxs-lookup"><span data-stu-id="e64c3-127">No</span></span>  | <span data-ttu-id="e64c3-128">La valeur par `lifetime` défaut est `short` et n’a pas besoin d’être spécifiée.</span><span class="sxs-lookup"><span data-stu-id="e64c3-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="e64c3-129">Outlook’utilisent que la `short` valeur.</span><span class="sxs-lookup"><span data-stu-id="e64c3-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="e64c3-130">Si vous souhaitez utiliser un runtime partagé dans un Excel, définissez explicitement la valeur sur `long` .</span><span class="sxs-lookup"><span data-stu-id="e64c3-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="e64c3-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e64c3-131">See also</span></span>

- [<span data-ttu-id="e64c3-132">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="e64c3-132">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="e64c3-133">Configurer votre complément Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="e64c3-133">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="e64c3-134">Configurer votre complément Outlook pour l’activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="e64c3-134">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
