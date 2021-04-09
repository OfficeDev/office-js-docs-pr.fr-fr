---
title: Runtimes dans le fichier manifeste
description: L’élément Runtimes spécifie le runtime de votre add-in.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: a5cd05a0890615375bf3466caf70d22f9912d951
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652232"
---
# <a name="runtimes-element"></a><span data-ttu-id="4810e-103">Élément Runtimes</span><span class="sxs-lookup"><span data-stu-id="4810e-103">Runtimes element</span></span>

<span data-ttu-id="4810e-104">Spécifie le runtime de votre add-in.</span><span class="sxs-lookup"><span data-stu-id="4810e-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="4810e-105">Enfant de [`<Host>`](host.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="4810e-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="4810e-106">Lorsque vous exécutez Office sur Windows, votre application utilise le navigateur Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="4810e-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="4810e-107">**Type de add-in :** Volet De tâches, Courrier</span><span class="sxs-lookup"><span data-stu-id="4810e-107">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="4810e-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="4810e-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="4810e-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="4810e-109">Contained in</span></span>

[<span data-ttu-id="4810e-110">Host</span><span class="sxs-lookup"><span data-stu-id="4810e-110">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="4810e-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="4810e-111">Child elements</span></span>

|  <span data-ttu-id="4810e-112">Élément</span><span class="sxs-lookup"><span data-stu-id="4810e-112">Element</span></span> |  <span data-ttu-id="4810e-113">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="4810e-113">Required</span></span>  |  <span data-ttu-id="4810e-114">Description</span><span class="sxs-lookup"><span data-stu-id="4810e-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="4810e-115">Runtime</span><span class="sxs-lookup"><span data-stu-id="4810e-115">Runtime</span></span>](runtime.md) | <span data-ttu-id="4810e-116">Oui</span><span class="sxs-lookup"><span data-stu-id="4810e-116">Yes</span></span> |  <span data-ttu-id="4810e-117">Runtime de votre add-in.</span><span class="sxs-lookup"><span data-stu-id="4810e-117">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="4810e-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4810e-118">See also</span></span>

- [<span data-ttu-id="4810e-119">Runtime</span><span class="sxs-lookup"><span data-stu-id="4810e-119">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="4810e-120">Configurer votre complément Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="4810e-120">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="4810e-121">Configurer votre complément Outlook pour l’activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="4810e-121">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
