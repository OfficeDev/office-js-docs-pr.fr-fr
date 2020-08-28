---
title: Élément GetStarted dans le fichier manifeste
description: Fournit des informations utilisées par la légende qui apparaît lorsque le complément est installé dans Word, Excel, PowerPoint et OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 01b10b8316c87b046cf816d6f86551bf1a349267
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292292"
---
# <a name="getstarted-element"></a><span data-ttu-id="8e773-103">Élément GetStarted</span><span class="sxs-lookup"><span data-stu-id="8e773-103">GetStarted element</span></span>

<span data-ttu-id="8e773-104">Fournit des informations utilisées par la légende qui apparaît lorsque le complément est installé dans Word, Excel, PowerPoint et OneNote.</span><span class="sxs-lookup"><span data-stu-id="8e773-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="8e773-105">L’élément **GetStarted** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="8e773-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="8e773-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="8e773-106">Child elements</span></span>

| <span data-ttu-id="8e773-107">Élément</span><span class="sxs-lookup"><span data-stu-id="8e773-107">Element</span></span>                       | <span data-ttu-id="8e773-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="8e773-108">Required</span></span> | <span data-ttu-id="8e773-109">Description</span><span class="sxs-lookup"><span data-stu-id="8e773-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="8e773-110">Titre</span><span class="sxs-lookup"><span data-stu-id="8e773-110">Title</span></span>](#title)               | <span data-ttu-id="8e773-111">Oui</span><span class="sxs-lookup"><span data-stu-id="8e773-111">Yes</span></span>      | <span data-ttu-id="8e773-112">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="8e773-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="8e773-113">Description</span><span class="sxs-lookup"><span data-stu-id="8e773-113">Description</span></span>](#description)   | <span data-ttu-id="8e773-114">Oui</span><span class="sxs-lookup"><span data-stu-id="8e773-114">Yes</span></span>      | <span data-ttu-id="8e773-115">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8e773-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="8e773-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="8e773-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="8e773-117">Oui</span><span class="sxs-lookup"><span data-stu-id="8e773-117">Yes</span></span>       | <span data-ttu-id="8e773-118">URL vers une page qui décrit le complément de façon plus détaillée.</span><span class="sxs-lookup"><span data-stu-id="8e773-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="8e773-119">Titre</span><span class="sxs-lookup"><span data-stu-id="8e773-119">Title</span></span> 

<span data-ttu-id="8e773-p102">Obligatoire. Le titre est utilisé pour la partie supérieure de la légende. L’attribut **resid** fait référence à un ID valide de l’élément **ShortStrings** dans la section [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="8e773-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="8e773-123">Description</span><span class="sxs-lookup"><span data-stu-id="8e773-123">Description</span></span>

<span data-ttu-id="8e773-p103">Obligatoire. Description/Contenu du corps de la légende. L’attribut **resid** fait référence à un ID valide de l’élément **LongStrings** dans la section [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="8e773-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="8e773-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="8e773-127">LearnMoreUrl</span></span>

<span data-ttu-id="8e773-p104">Obligatoire. URL vers une page dans laquelle l’utilisateur peut obtenir des informations sur votre complément. L’attribut **resid** fait référence à un ID valide de l’élément **Urls** dans la section [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="8e773-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="8e773-131">**LearnMoreUrl** n’est pas actuellement restitué dans les clients Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="8e773-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="8e773-132">Nous vous recommandons d’ajouter cette URL pour tous les clients afin que l’URL soit restituée lorsqu’elle est disponible.</span><span class="sxs-lookup"><span data-stu-id="8e773-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="8e773-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8e773-133">See also</span></span>

<span data-ttu-id="8e773-134">Les exemples de code suivants utilisent l’élément **GetStarted** :</span><span class="sxs-lookup"><span data-stu-id="8e773-134">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="8e773-135">Complément web Excel pour manipuler la mise en forme de tableau et de graphique</span><span class="sxs-lookup"><span data-stu-id="8e773-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="8e773-136">Complément Word JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="8e773-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="8e773-137">Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8e773-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
