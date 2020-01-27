---
title: Runtimes dans le fichier manifeste
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 6682887935ee6894b5a311ad519408067452bb23
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554005"
---
# <a name="runtimes-element"></a><span data-ttu-id="01b65-102">Élément runtimes</span><span class="sxs-lookup"><span data-stu-id="01b65-102">Runtimes element</span></span>

<span data-ttu-id="01b65-103">Cette fonctionnalité est en aperçu.</span><span class="sxs-lookup"><span data-stu-id="01b65-103">This feature is in preview.</span></span> <span data-ttu-id="01b65-104">Spécifie le runtime de votre complément et permet aux fonctions personnalisées et au volet Office de partager des données globales et d’effectuer des appels de fonction.</span><span class="sxs-lookup"><span data-stu-id="01b65-104">Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other.</span></span> <span data-ttu-id="01b65-105">Doit suivre l' `<Host>` élément dans votre fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="01b65-105">Should follow the `<Host>` element in your manifest file.</span></span>

<span data-ttu-id="01b65-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="01b65-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="01b65-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="01b65-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a><span data-ttu-id="01b65-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="01b65-108">Child elements</span></span>

|  <span data-ttu-id="01b65-109">Élément</span><span class="sxs-lookup"><span data-stu-id="01b65-109">Element</span></span> |  <span data-ttu-id="01b65-110">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="01b65-110">Required</span></span>  |  <span data-ttu-id="01b65-111">Description</span><span class="sxs-lookup"><span data-stu-id="01b65-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="01b65-112">**Runtime**</span><span class="sxs-lookup"><span data-stu-id="01b65-112">**Runtime**</span></span>     | <span data-ttu-id="01b65-113">Oui</span><span class="sxs-lookup"><span data-stu-id="01b65-113">Yes</span></span> |  <span data-ttu-id="01b65-114">Le runtime de votre complément, souvent utilisé avec des fonctions personnalisées Excel.</span><span class="sxs-lookup"><span data-stu-id="01b65-114">The Runtime for your add-in, often used with Excel custom functions.</span></span>

## <a name="see-also"></a><span data-ttu-id="01b65-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="01b65-115">See also</span></span>

- [<span data-ttu-id="01b65-116">Runtime</span><span class="sxs-lookup"><span data-stu-id="01b65-116">Runtime</span></span>](runtime.md)
