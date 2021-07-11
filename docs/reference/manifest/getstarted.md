---
title: Élément GetStarted dans le fichier manifeste
description: Fournit des informations utilisées par la callout qui s’affiche lorsque le add-in est installé dans Word, Excel, PowerPoint et OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a637f3f9031d9f8e09d14f17f2095ca0647c4d50
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348684"
---
# <a name="getstarted-element"></a><span data-ttu-id="edcfe-103">Élément GetStarted</span><span class="sxs-lookup"><span data-stu-id="edcfe-103">GetStarted element</span></span>

<span data-ttu-id="edcfe-104">Fournit des informations utilisées par la callout qui s’affiche lorsque le add-in est installé dans Word, Excel, PowerPoint et OneNote.</span><span class="sxs-lookup"><span data-stu-id="edcfe-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="edcfe-105">L’élément **GetStarted** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="edcfe-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="edcfe-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="edcfe-106">Child elements</span></span>

| <span data-ttu-id="edcfe-107">Élément</span><span class="sxs-lookup"><span data-stu-id="edcfe-107">Element</span></span>                       | <span data-ttu-id="edcfe-108">Requis</span><span class="sxs-lookup"><span data-stu-id="edcfe-108">Required</span></span> | <span data-ttu-id="edcfe-109">Description</span><span class="sxs-lookup"><span data-stu-id="edcfe-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="edcfe-110">Titre</span><span class="sxs-lookup"><span data-stu-id="edcfe-110">Title</span></span>](#title)               | <span data-ttu-id="edcfe-111">Oui</span><span class="sxs-lookup"><span data-stu-id="edcfe-111">Yes</span></span>      | <span data-ttu-id="edcfe-112">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="edcfe-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="edcfe-113">Description</span><span class="sxs-lookup"><span data-stu-id="edcfe-113">Description</span></span>](#description)   | <span data-ttu-id="edcfe-114">Oui</span><span class="sxs-lookup"><span data-stu-id="edcfe-114">Yes</span></span>      | <span data-ttu-id="edcfe-115">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="edcfe-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="edcfe-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="edcfe-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="edcfe-117">Oui</span><span class="sxs-lookup"><span data-stu-id="edcfe-117">Yes</span></span>       | <span data-ttu-id="edcfe-118">URL vers une page qui décrit le complément de façon plus détaillée.</span><span class="sxs-lookup"><span data-stu-id="edcfe-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="edcfe-119">Titre</span><span class="sxs-lookup"><span data-stu-id="edcfe-119">Title</span></span> 

<span data-ttu-id="edcfe-120">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="edcfe-120">Required.</span></span> <span data-ttu-id="edcfe-121">Le titre est utilisé pour la partie supérieure de la légende.</span><span class="sxs-lookup"><span data-stu-id="edcfe-121">The title used for the top of the callout.</span></span> <span data-ttu-id="edcfe-122">**L’attribut resid** fait référence à un ID valide dans l’élément **ShortStrings** de la section [Resources](resources.md) et ne peut pas avoir plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="edcfe-122">The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="description"></a><span data-ttu-id="edcfe-123">Description</span><span class="sxs-lookup"><span data-stu-id="edcfe-123">Description</span></span>

<span data-ttu-id="edcfe-124">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="edcfe-124">Required.</span></span> <span data-ttu-id="edcfe-125">Description/Contenu du corps de la légende.</span><span class="sxs-lookup"><span data-stu-id="edcfe-125">The description / body content for the callout.</span></span> <span data-ttu-id="edcfe-126">**L’attribut resid** fait référence à un ID valide dans l’élément **LongStrings** de la section [Resources](resources.md) et ne peut pas être plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="edcfe-126">The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="edcfe-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="edcfe-127">LearnMoreUrl</span></span>

<span data-ttu-id="edcfe-128">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="edcfe-128">Required.</span></span> <span data-ttu-id="edcfe-129">URL vers une page dans laquelle l’utilisateur peut obtenir des informations sur votre complément.</span><span class="sxs-lookup"><span data-stu-id="edcfe-129">The URL to a page where the user can learn more about your add-in.</span></span> <span data-ttu-id="edcfe-130">**L’attribut resid** fait référence à un ID valide dans l’élément **Urls** de la section [Resources](resources.md) et ne peut pas avoir plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="edcfe-130">The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

> [!NOTE]
> <span data-ttu-id="edcfe-131">**LearnMoreUrl** n’est pas actuellement restitué dans les clients Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="edcfe-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="edcfe-132">Nous vous recommandons d’ajouter cette URL pour tous les clients afin que l’URL soit restituée lorsqu’elle est disponible.</span><span class="sxs-lookup"><span data-stu-id="edcfe-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="edcfe-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="edcfe-133">See also</span></span>

<span data-ttu-id="edcfe-134">Les exemples de code suivants utilisent **l’élément GetStarted.**</span><span class="sxs-lookup"><span data-stu-id="edcfe-134">The following code samples use the **GetStarted** element.</span></span>

* [<span data-ttu-id="edcfe-135">Complément web Excel pour manipuler la mise en forme de tableau et de graphique</span><span class="sxs-lookup"><span data-stu-id="edcfe-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="edcfe-136">Complément Word JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="edcfe-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="edcfe-137">Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint</span><span class="sxs-lookup"><span data-stu-id="edcfe-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
