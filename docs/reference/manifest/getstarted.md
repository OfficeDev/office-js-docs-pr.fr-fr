---
title: Élément GetStarted dans le fichier manifeste
description: Fournit des informations utilisées par la légende qui s’affiche lorsque le complément est installé dans des hôtes Word, Excel, PowerPoint et OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7b4eb0e0413eff97751eac01c3dc6e5a452722aa
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718187"
---
# <a name="getstarted-element"></a><span data-ttu-id="55055-103">GetStarted, élément</span><span class="sxs-lookup"><span data-stu-id="55055-103">GetStarted element</span></span>

<span data-ttu-id="55055-p101">Fournit des informations utilisées par la légende qui s’affiche lorsque le complément est installé dans des hôtes Word, Excel, PowerPoint et OneNote. L’élément **GetStarted** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="55055-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="55055-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="55055-106">Child elements</span></span>

| <span data-ttu-id="55055-107">Élément</span><span class="sxs-lookup"><span data-stu-id="55055-107">Element</span></span>                       | <span data-ttu-id="55055-108">Requis</span><span class="sxs-lookup"><span data-stu-id="55055-108">Required</span></span> | <span data-ttu-id="55055-109">Description</span><span class="sxs-lookup"><span data-stu-id="55055-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="55055-110">Titre</span><span class="sxs-lookup"><span data-stu-id="55055-110">Title</span></span>](#title)               | <span data-ttu-id="55055-111">Oui</span><span class="sxs-lookup"><span data-stu-id="55055-111">Yes</span></span>      | <span data-ttu-id="55055-112">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="55055-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="55055-113">Description</span><span class="sxs-lookup"><span data-stu-id="55055-113">Description</span></span>](#description)   | <span data-ttu-id="55055-114">Oui</span><span class="sxs-lookup"><span data-stu-id="55055-114">Yes</span></span>      | <span data-ttu-id="55055-115">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="55055-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="55055-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="55055-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="55055-117">Oui</span><span class="sxs-lookup"><span data-stu-id="55055-117">Yes</span></span>       | <span data-ttu-id="55055-118">URL vers une page qui décrit le complément de façon plus détaillée.</span><span class="sxs-lookup"><span data-stu-id="55055-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="55055-119">Titre</span><span class="sxs-lookup"><span data-stu-id="55055-119">Title</span></span> 

<span data-ttu-id="55055-p102">Obligatoire. Le titre est utilisé pour la partie supérieure de la légende. L’attribut **resid** fait référence à un ID valide de l’élément **ShortStrings** dans la section [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="55055-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="55055-123">Description</span><span class="sxs-lookup"><span data-stu-id="55055-123">Description</span></span>

<span data-ttu-id="55055-p103">Obligatoire. Description/Contenu du corps de la légende. L’attribut **resid** fait référence à un ID valide de l’élément **LongStrings** dans la section [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="55055-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="55055-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="55055-127">LearnMoreUrl</span></span>

<span data-ttu-id="55055-p104">Obligatoire. URL vers une page dans laquelle l’utilisateur peut obtenir des informations sur votre complément. L’attribut **resid** fait référence à un ID valide de l’élément **Urls** dans la section [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="55055-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="55055-131">**LearnMoreUrl** n’est pas actuellement restitué dans les clients Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="55055-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="55055-132">Nous vous recommandons d’ajouter cette URL pour tous les clients afin que l’URL soit restituée lorsqu’elle est disponible.</span><span class="sxs-lookup"><span data-stu-id="55055-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="55055-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="55055-133">See also</span></span>

<span data-ttu-id="55055-134">Les exemples de code suivants utilisent l’élément **GetStarted** :</span><span class="sxs-lookup"><span data-stu-id="55055-134">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="55055-135">Complément web Excel pour manipuler la mise en forme de tableau et de graphique</span><span class="sxs-lookup"><span data-stu-id="55055-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="55055-136">Complément Word JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="55055-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="55055-137">Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint</span><span class="sxs-lookup"><span data-stu-id="55055-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
