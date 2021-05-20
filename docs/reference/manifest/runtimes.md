---
title: Temps d’exécution dans le fichier manifeste
description: L’élément Runtimes spécifie le temps d’exécution de votre module.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555296"
---
# <a name="runtimes-element"></a><span data-ttu-id="06530-103">Élément runtimes</span><span class="sxs-lookup"><span data-stu-id="06530-103">Runtimes element</span></span>

<span data-ttu-id="06530-104">Spécifie le temps d’exécution de votre module d’exécution.</span><span class="sxs-lookup"><span data-stu-id="06530-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="06530-105">Enfant de [`<Host>`](host.md) l’élément.</span><span class="sxs-lookup"><span data-stu-id="06530-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="06530-106">Lors de l’exécution Office sur Windows, un add-in qui a un élément dans son manifeste ne fonctionne `<Runtimes>` pas nécessairement dans le même contrôle webview qu’il le ferait autrement.</span><span class="sxs-lookup"><span data-stu-id="06530-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="06530-107">Pour plus d’informations sur la façon dont les versions de Windows et Office déterminent quel contrôle webview est normalement utilisé, [voir Navigateurs utilisés par Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). Si les conditions décrites pour l’utilisation de Microsoft Edge avec WebView2 (basé sur Chromium) sont remplies, alors l’add-in utilise ce navigateur, qu’il ait ou non un `<Runtimes>` élément.</span><span class="sxs-lookup"><span data-stu-id="06530-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="06530-108">Toutefois, lorsque ces conditions ne sont pas remplies, un module avec un `<Runtimes>` élément utilise toujours Internet Explorer 11 indépendamment de la version Windows ou Microsoft 365 version.</span><span class="sxs-lookup"><span data-stu-id="06530-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="06530-109">**Type d’add-in :** Volet de tâche, Courrier</span><span class="sxs-lookup"><span data-stu-id="06530-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="06530-110">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="06530-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="06530-111">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="06530-111">Contained in</span></span>

[<span data-ttu-id="06530-112">Host</span><span class="sxs-lookup"><span data-stu-id="06530-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="06530-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="06530-113">Child elements</span></span>

|  <span data-ttu-id="06530-114">Élément</span><span class="sxs-lookup"><span data-stu-id="06530-114">Element</span></span> |  <span data-ttu-id="06530-115">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="06530-115">Required</span></span>  |  <span data-ttu-id="06530-116">Description</span><span class="sxs-lookup"><span data-stu-id="06530-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="06530-117">Runtime</span><span class="sxs-lookup"><span data-stu-id="06530-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="06530-118">Oui</span><span class="sxs-lookup"><span data-stu-id="06530-118">Yes</span></span> |  <span data-ttu-id="06530-119">Le temps d’exécution de votre add-in.</span><span class="sxs-lookup"><span data-stu-id="06530-119">The runtime for your add-in.</span></span> <span data-ttu-id="06530-120">**Important :** À l’heure actuelle, vous ne pouvez définir qu’un `<Runtime>` seul élément.</span><span class="sxs-lookup"><span data-stu-id="06530-120">**Important**: At present, you can only define one `<Runtime>` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="06530-121">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="06530-121">See also</span></span>

- [<span data-ttu-id="06530-122">Runtime</span><span class="sxs-lookup"><span data-stu-id="06530-122">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="06530-123">Configurer votre complément Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="06530-123">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="06530-124">Configurez votre Outlook add-in pour l’activation basée sur l’événement</span><span class="sxs-lookup"><span data-stu-id="06530-124">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
