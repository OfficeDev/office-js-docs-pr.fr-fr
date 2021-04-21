---
title: Runtimes dans le fichier manifeste
description: L'élément Runtimes spécifie le runtime de votre add-in.
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917085"
---
# <a name="runtimes-element"></a><span data-ttu-id="52b7d-103">Élément Runtimes</span><span class="sxs-lookup"><span data-stu-id="52b7d-103">Runtimes element</span></span>

<span data-ttu-id="52b7d-104">Spécifie le runtime de votre add-in.</span><span class="sxs-lookup"><span data-stu-id="52b7d-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="52b7d-105">Enfant de [`<Host>`](host.md) l'élément.</span><span class="sxs-lookup"><span data-stu-id="52b7d-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="52b7d-106">Lors de l'exécution dans Office sur Windows, un add-in qui possède un élément dans son manifeste ne s'exécute pas nécessairement dans le même contrôle webview que dans le `<Runtimes>` cas contraire.</span><span class="sxs-lookup"><span data-stu-id="52b7d-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="52b7d-107">Pour plus d'informations sur la façon dont les versions de Windows et d'Office déterminent le contrôle webview utilisé normalement, voir Navigateurs utilisés par les [applications Office.](../../concepts/browsers-used-by-office-web-add-ins.md) Si les conditions décrites ici pour l'utilisation de Microsoft Edge avec WebView2 (basé sur Chromium) sont remplies, le add-in utilise ce navigateur, qu'il ait ou non un `<Runtimes>` élément.</span><span class="sxs-lookup"><span data-stu-id="52b7d-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="52b7d-108">Toutefois, lorsque ces conditions ne sont pas remplies, un add-in avec un élément utilise toujours Internet Explorer 11, quelle que soit la version de Windows ou `<Runtimes>` de Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="52b7d-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="52b7d-109">**Type de add-in :** Volet De tâches, Courrier</span><span class="sxs-lookup"><span data-stu-id="52b7d-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="52b7d-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="52b7d-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="52b7d-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="52b7d-111">Contained in</span></span>

[<span data-ttu-id="52b7d-112">Host</span><span class="sxs-lookup"><span data-stu-id="52b7d-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="52b7d-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="52b7d-113">Child elements</span></span>

|  <span data-ttu-id="52b7d-114">Élément</span><span class="sxs-lookup"><span data-stu-id="52b7d-114">Element</span></span> |  <span data-ttu-id="52b7d-115">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="52b7d-115">Required</span></span>  |  <span data-ttu-id="52b7d-116">Description</span><span class="sxs-lookup"><span data-stu-id="52b7d-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="52b7d-117">Runtime</span><span class="sxs-lookup"><span data-stu-id="52b7d-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="52b7d-118">Oui</span><span class="sxs-lookup"><span data-stu-id="52b7d-118">Yes</span></span> |  <span data-ttu-id="52b7d-119">Runtime de votre add-in.</span><span class="sxs-lookup"><span data-stu-id="52b7d-119">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="52b7d-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="52b7d-120">See also</span></span>

- [<span data-ttu-id="52b7d-121">Runtime</span><span class="sxs-lookup"><span data-stu-id="52b7d-121">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="52b7d-122">Configurer votre complément Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="52b7d-122">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="52b7d-123">Configurer votre complément Outlook pour l'activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="52b7d-123">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
