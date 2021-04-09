---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre add-in pour utiliser un runtime JavaScript partagé pour ses différents composants, par exemple, ruban, volet des tâches, fonctions personnalisées.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652243"
---
# <a name="runtime-element"></a><span data-ttu-id="c4b0c-103">Élément Runtime</span><span class="sxs-lookup"><span data-stu-id="c4b0c-103">Runtime element</span></span>

<span data-ttu-id="c4b0c-104">Configure votre add-in pour utiliser un runtime JavaScript partagé afin que différents composants s’exécutent tous dans le même runtime.</span><span class="sxs-lookup"><span data-stu-id="c4b0c-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="c4b0c-105">Enfant de [`<Runtimes>`](runtimes.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="c4b0c-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="c4b0c-106">**Type de add-in :** Volet De tâches, Courrier</span><span class="sxs-lookup"><span data-stu-id="c4b0c-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="c4b0c-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="c4b0c-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="c4b0c-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="c4b0c-108">Contained in</span></span>

- [<span data-ttu-id="c4b0c-109">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="c4b0c-109">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="c4b0c-110">Attributs</span><span class="sxs-lookup"><span data-stu-id="c4b0c-110">Attributes</span></span>

|  <span data-ttu-id="c4b0c-111">Attribut</span><span class="sxs-lookup"><span data-stu-id="c4b0c-111">Attribute</span></span>  |  <span data-ttu-id="c4b0c-112">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="c4b0c-112">Required</span></span>  |  <span data-ttu-id="c4b0c-113">Description</span><span class="sxs-lookup"><span data-stu-id="c4b0c-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c4b0c-114">**resid**</span><span class="sxs-lookup"><span data-stu-id="c4b0c-114">**resid**</span></span>  |  <span data-ttu-id="c4b0c-115">Oui</span><span class="sxs-lookup"><span data-stu-id="c4b0c-115">Yes</span></span>  | <span data-ttu-id="c4b0c-116">Spécifie l’emplacement URL de la page HTML de votre application.</span><span class="sxs-lookup"><span data-stu-id="c4b0c-116">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="c4b0c-117">Il ne peut pas y avoir plus de 32 caractères et doit correspondre à un `resid` `id` attribut `Url` d’un élément dans `Resources` l’élément.</span><span class="sxs-lookup"><span data-stu-id="c4b0c-117">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="c4b0c-118">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="c4b0c-118">**lifetime**</span></span>  |  <span data-ttu-id="c4b0c-119">Non</span><span class="sxs-lookup"><span data-stu-id="c4b0c-119">No</span></span>  | <span data-ttu-id="c4b0c-120">La valeur par `lifetime` défaut est `short` et n’a pas besoin d’être spécifiée.</span><span class="sxs-lookup"><span data-stu-id="c4b0c-120">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="c4b0c-121">Les add-ins Outlook utilisent uniquement la `short` valeur.</span><span class="sxs-lookup"><span data-stu-id="c4b0c-121">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="c4b0c-122">Si vous souhaitez utiliser un runtime partagé dans un add-in Excel, définissez explicitement la valeur sur `long` .</span><span class="sxs-lookup"><span data-stu-id="c4b0c-122">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="c4b0c-123">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c4b0c-123">See also</span></span>

- [<span data-ttu-id="c4b0c-124">Services d’exécution</span><span class="sxs-lookup"><span data-stu-id="c4b0c-124">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="c4b0c-125">Configurer votre complément Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="c4b0c-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="c4b0c-126">Configurer votre complément Outlook pour l’activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="c4b0c-126">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
