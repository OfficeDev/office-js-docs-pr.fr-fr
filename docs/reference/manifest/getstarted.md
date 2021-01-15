---
title: Élément GetStarted dans le fichier manifeste
description: Fournit des informations utilisées par la légende qui apparaît lorsque le complément est installé dans Word, Excel, PowerPoint et OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0ad6196dc45e4ea06c2b43ac5da66a560ab0b899
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771413"
---
# <a name="getstarted-element"></a><span data-ttu-id="7fd04-103">Élément GetStarted</span><span class="sxs-lookup"><span data-stu-id="7fd04-103">GetStarted element</span></span>

<span data-ttu-id="7fd04-104">Fournit des informations utilisées par la légende qui apparaît lorsque le complément est installé dans Word, Excel, PowerPoint et OneNote.</span><span class="sxs-lookup"><span data-stu-id="7fd04-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="7fd04-105">L’élément **GetStarted** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="7fd04-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="7fd04-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="7fd04-106">Child elements</span></span>

| <span data-ttu-id="7fd04-107">Élément</span><span class="sxs-lookup"><span data-stu-id="7fd04-107">Element</span></span>                       | <span data-ttu-id="7fd04-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7fd04-108">Required</span></span> | <span data-ttu-id="7fd04-109">Description</span><span class="sxs-lookup"><span data-stu-id="7fd04-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="7fd04-110">Titre</span><span class="sxs-lookup"><span data-stu-id="7fd04-110">Title</span></span>](#title)               | <span data-ttu-id="7fd04-111">Oui</span><span class="sxs-lookup"><span data-stu-id="7fd04-111">Yes</span></span>      | <span data-ttu-id="7fd04-112">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="7fd04-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="7fd04-113">Description</span><span class="sxs-lookup"><span data-stu-id="7fd04-113">Description</span></span>](#description)   | <span data-ttu-id="7fd04-114">Oui</span><span class="sxs-lookup"><span data-stu-id="7fd04-114">Yes</span></span>      | <span data-ttu-id="7fd04-115">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="7fd04-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="7fd04-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="7fd04-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="7fd04-117">Oui</span><span class="sxs-lookup"><span data-stu-id="7fd04-117">Yes</span></span>       | <span data-ttu-id="7fd04-118">URL vers une page qui décrit le complément de façon plus détaillée.</span><span class="sxs-lookup"><span data-stu-id="7fd04-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="7fd04-119">Titre</span><span class="sxs-lookup"><span data-stu-id="7fd04-119">Title</span></span> 

<span data-ttu-id="7fd04-120">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="7fd04-120">Required.</span></span> <span data-ttu-id="7fd04-121">Le titre est utilisé pour la partie supérieure de la légende.</span><span class="sxs-lookup"><span data-stu-id="7fd04-121">The title used for the top of the callout.</span></span> <span data-ttu-id="7fd04-122">L’attribut **RESID** fait référence à un ID valide dans l’élément **ShortStrings** dans la section [Resources](resources.md) et ne peut pas contenir plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="7fd04-122">The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="description"></a><span data-ttu-id="7fd04-123">Description</span><span class="sxs-lookup"><span data-stu-id="7fd04-123">Description</span></span>

<span data-ttu-id="7fd04-124">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="7fd04-124">Required.</span></span> <span data-ttu-id="7fd04-125">Description/Contenu du corps de la légende.</span><span class="sxs-lookup"><span data-stu-id="7fd04-125">The description / body content for the callout.</span></span> <span data-ttu-id="7fd04-126">L’attribut **RESID** fait référence à un ID valide dans l’élément **LongStrings** dans la section [Resources](resources.md) et ne peut pas contenir plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="7fd04-126">The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="7fd04-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="7fd04-127">LearnMoreUrl</span></span>

<span data-ttu-id="7fd04-128">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="7fd04-128">Required.</span></span> <span data-ttu-id="7fd04-129">URL vers une page dans laquelle l’utilisateur peut obtenir des informations sur votre complément.</span><span class="sxs-lookup"><span data-stu-id="7fd04-129">The URL to a page where the user can learn more about your add-in.</span></span> <span data-ttu-id="7fd04-130">L’attribut **RESID** fait référence à un ID valide dans l’élément **URLs** de la section [Resources](resources.md) et ne peut pas contenir plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="7fd04-130">The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

> [!NOTE]
> <span data-ttu-id="7fd04-131">**LearnMoreUrl** n’est pas actuellement restitué dans les clients Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="7fd04-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="7fd04-132">Nous vous recommandons d’ajouter cette URL pour tous les clients afin que l’URL soit restituée lorsqu’elle est disponible.</span><span class="sxs-lookup"><span data-stu-id="7fd04-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="7fd04-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7fd04-133">See also</span></span>

<span data-ttu-id="7fd04-134">Les exemples de code suivants utilisent l’élément **GetStarted** :</span><span class="sxs-lookup"><span data-stu-id="7fd04-134">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="7fd04-135">Complément web Excel pour manipuler la mise en forme de tableau et de graphique</span><span class="sxs-lookup"><span data-stu-id="7fd04-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="7fd04-136">Complément Word JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="7fd04-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="7fd04-137">Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7fd04-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
