---
title: Temps d’exécution dans le fichier manifeste
description: L’élément Runtime configure votre module d’ajout pour utiliser un temps d’exécution JavaScript partagé pour ses différents composants, par exemple, ruban, volet de tâches, fonctions personnalisées.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: c59e5a23e53940aea46c758d710b4a455cb5c0cc
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555303"
---
# <a name="runtime-element"></a><span data-ttu-id="df4ba-103">Élément runtime</span><span class="sxs-lookup"><span data-stu-id="df4ba-103">Runtime element</span></span>

<span data-ttu-id="df4ba-104">Configure votre module d’ajout pour utiliser un temps d’exécution JavaScript partagé afin que les différents composants s’exécutent tous dans le même temps d’exécution.</span><span class="sxs-lookup"><span data-stu-id="df4ba-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="df4ba-105">Enfant de [`<Runtimes>`](runtimes.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="df4ba-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="df4ba-106">**Type d’add-in :** Volet de tâche, Courrier</span><span class="sxs-lookup"><span data-stu-id="df4ba-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="df4ba-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="df4ba-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="df4ba-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="df4ba-108">Contained in</span></span>

- [<span data-ttu-id="df4ba-109">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="df4ba-109">Runtimes</span></span>](runtimes.md)

## <a name="child-elements"></a><span data-ttu-id="df4ba-110">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="df4ba-110">Child elements</span></span>

|  <span data-ttu-id="df4ba-111">Élément</span><span class="sxs-lookup"><span data-stu-id="df4ba-111">Element</span></span> |  <span data-ttu-id="df4ba-112">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="df4ba-112">Required</span></span>  |  <span data-ttu-id="df4ba-113">Description</span><span class="sxs-lookup"><span data-stu-id="df4ba-113">Description</span></span>  |
|:-----|:-----|:-----|
| <span data-ttu-id="df4ba-114">[Override](override.md) (aperçu)</span><span class="sxs-lookup"><span data-stu-id="df4ba-114">[Override](override.md) (preview)</span></span> | <span data-ttu-id="df4ba-115">Non</span><span class="sxs-lookup"><span data-stu-id="df4ba-115">No</span></span> | <span data-ttu-id="df4ba-116">**Outlook**: Spécifie l’emplacement de l’URL du fichier JavaScript Outlook Desktop nécessite pour [les gestionnaires de points d’extension LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview)</span><span class="sxs-lookup"><span data-stu-id="df4ba-116">**Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent-preview) handlers.</span></span> <span data-ttu-id="df4ba-117">**Important**: À l’heure actuelle, vous ne pouvez définir `<Override>` qu’un seul élément et il doit être de type `javascript` .</span><span class="sxs-lookup"><span data-stu-id="df4ba-117">**Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.</span></span>|

## <a name="attributes"></a><span data-ttu-id="df4ba-118">Attributs</span><span class="sxs-lookup"><span data-stu-id="df4ba-118">Attributes</span></span>

|  <span data-ttu-id="df4ba-119">Attribut</span><span class="sxs-lookup"><span data-stu-id="df4ba-119">Attribute</span></span>  |  <span data-ttu-id="df4ba-120">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="df4ba-120">Required</span></span>  |  <span data-ttu-id="df4ba-121">Description</span><span class="sxs-lookup"><span data-stu-id="df4ba-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="df4ba-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="df4ba-122">**resid**</span></span>  |  <span data-ttu-id="df4ba-123">Oui</span><span class="sxs-lookup"><span data-stu-id="df4ba-123">Yes</span></span>  | <span data-ttu-id="df4ba-124">Spécifie l’emplacement de l’URL de la page HTML pour votre module d’ajout.</span><span class="sxs-lookup"><span data-stu-id="df4ba-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="df4ba-125">Le `resid` ne peut pas être plus de 32 caractères et doit correspondre à un attribut `id` d’un `Url` élément dans `Resources` l’élément.</span><span class="sxs-lookup"><span data-stu-id="df4ba-125">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="df4ba-126">**vie**</span><span class="sxs-lookup"><span data-stu-id="df4ba-126">**lifetime**</span></span>  |  <span data-ttu-id="df4ba-127">Non</span><span class="sxs-lookup"><span data-stu-id="df4ba-127">No</span></span>  | <span data-ttu-id="df4ba-128">La valeur par défaut `lifetime` pour est `short` et n’a pas besoin d’être spécifiée.</span><span class="sxs-lookup"><span data-stu-id="df4ba-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="df4ba-129">Outlook add-ins n’utilisent que la `short` valeur.</span><span class="sxs-lookup"><span data-stu-id="df4ba-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="df4ba-130">Si vous souhaitez utiliser un temps d’exécution partagé dans un Excel add-in, définissez explicitement la valeur à `long` .</span><span class="sxs-lookup"><span data-stu-id="df4ba-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="df4ba-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="df4ba-131">See also</span></span>

- [<span data-ttu-id="df4ba-132">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="df4ba-132">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="df4ba-133">Configurer votre complément Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="df4ba-133">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="df4ba-134">Configurez votre Outlook add-in pour l’activation basée sur l’événement</span><span class="sxs-lookup"><span data-stu-id="df4ba-134">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
