---
title: Élément GetStarted dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 82fa1b9b62674adfb05c07536a7fdf2bbabf8f45
ms.sourcegitcommit: e5a5ec4ba32bacd0ccd13291b4e7f4bfc42901a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2019
ms.locfileid: "37429736"
---
# <a name="getstarted-element"></a><span data-ttu-id="99a78-102">GetStarted, élément</span><span class="sxs-lookup"><span data-stu-id="99a78-102">GetStarted element</span></span>

<span data-ttu-id="99a78-p101">Fournit des informations utilisées par la légende qui s’affiche lorsque le complément est installé dans des hôtes Word, Excel, PowerPoint et OneNote. L’élément **GetStarted** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="99a78-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="99a78-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="99a78-105">Child elements</span></span>

| <span data-ttu-id="99a78-106">Élément</span><span class="sxs-lookup"><span data-stu-id="99a78-106">Element</span></span>                       | <span data-ttu-id="99a78-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="99a78-107">Required</span></span> | <span data-ttu-id="99a78-108">Description</span><span class="sxs-lookup"><span data-stu-id="99a78-108">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="99a78-109">Titre</span><span class="sxs-lookup"><span data-stu-id="99a78-109">Title</span></span>](#title)               | <span data-ttu-id="99a78-110">Oui</span><span class="sxs-lookup"><span data-stu-id="99a78-110">Yes</span></span>      | <span data-ttu-id="99a78-111">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="99a78-111">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="99a78-112">Description</span><span class="sxs-lookup"><span data-stu-id="99a78-112">Description</span></span>](#description)   | <span data-ttu-id="99a78-113">Oui</span><span class="sxs-lookup"><span data-stu-id="99a78-113">Yes</span></span>      | <span data-ttu-id="99a78-114">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="99a78-114">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="99a78-115">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="99a78-115">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="99a78-116">Oui</span><span class="sxs-lookup"><span data-stu-id="99a78-116">Yes</span></span>       | <span data-ttu-id="99a78-117">URL vers une page qui décrit le complément de façon plus détaillée.</span><span class="sxs-lookup"><span data-stu-id="99a78-117">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="99a78-118">Titre</span><span class="sxs-lookup"><span data-stu-id="99a78-118">Title</span></span> 

<span data-ttu-id="99a78-p102">Obligatoire. Le titre est utilisé pour la partie supérieure de la légende. L’attribut **resid** fait référence à un ID valide de l’élément **ShortStrings** dans la section [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="99a78-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="99a78-122">Description</span><span class="sxs-lookup"><span data-stu-id="99a78-122">Description</span></span>

<span data-ttu-id="99a78-p103">Obligatoire. Description/Contenu du corps de la légende. L’attribut **resid** fait référence à un ID valide de l’élément **LongStrings** dans la section [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="99a78-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="99a78-126">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="99a78-126">LearnMoreUrl</span></span>

<span data-ttu-id="99a78-p104">Obligatoire. URL vers une page dans laquelle l’utilisateur peut obtenir des informations sur votre complément. L’attribut **resid** fait référence à un ID valide de l’élément **Urls** dans la section [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="99a78-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="99a78-130">**LearnMoreUrl** n’est pas actuellement restitué dans les clients Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="99a78-130">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="99a78-131">Nous vous recommandons d’ajouter cette URL pour tous les clients afin que l’URL soit restituée lorsqu’elle est disponible.</span><span class="sxs-lookup"><span data-stu-id="99a78-131">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="99a78-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="99a78-132">See also</span></span>

<span data-ttu-id="99a78-133">Les exemples de code suivants utilisent l’élément **GetStarted** :</span><span class="sxs-lookup"><span data-stu-id="99a78-133">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="99a78-134">Complément web Excel pour manipuler la mise en forme de tableau et de graphique</span><span class="sxs-lookup"><span data-stu-id="99a78-134">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="99a78-135">Complément Word JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="99a78-135">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="99a78-136">Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint</span><span class="sxs-lookup"><span data-stu-id="99a78-136">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
