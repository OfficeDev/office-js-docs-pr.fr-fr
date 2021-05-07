---
title: Office UI Fabric dans des compléments Office
description: Obtenez une vue d’ensemble de l’utilisation Office composants UI Fabric dans Office des composants.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 20f926913335197a65ac24e4ec30ed0106b81bae
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253367"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="17ba5-103">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="17ba5-103">Office UI Fabric in Office Add-ins</span></span>

<span data-ttu-id="17ba5-104">Office UI Fabric est une infrastructure frontale JavaScript permettant de créer des expériences utilisateur pour Office.</span><span class="sxs-lookup"><span data-stu-id="17ba5-104">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office.</span></span> <span data-ttu-id="17ba5-105">Fabric propose des composants axés sur des visuels que vous pouvez étendre, retravailler et utiliser dans votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="17ba5-105">Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in.</span></span> <span data-ttu-id="17ba5-106">Fabric utilisant le langage de création d’Office, ses composants d’expérience utilisateur ressemblent à une extension naturelle d’Office.</span><span class="sxs-lookup"><span data-stu-id="17ba5-106">Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span>

<span data-ttu-id="17ba5-p102">Si vous créez un complément, nous vous encourageons à utiliser Office UI Fabric pour mettre au point l’expérience utilisateur. L’utilisation d’Office UI Fabric est facultative.</span><span class="sxs-lookup"><span data-stu-id="17ba5-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="17ba5-109">Les sections suivantes expliquent comment commencer à utiliser Fabric en fonction de vos besoins.</span><span class="sxs-lookup"><span data-stu-id="17ba5-109">The following sections explain how to get started using Fabric to meet your requirements.</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="17ba5-110">Utiliser Fabric Core : icônes, polices, couleurs</span><span class="sxs-lookup"><span data-stu-id="17ba5-110">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="17ba5-111">Fabric Core contient les principaux éléments du langage de création tels que les icônes, les couleurs, le type et la grille.</span><span class="sxs-lookup"><span data-stu-id="17ba5-111">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid.</span></span> <span data-ttu-id="17ba5-112">Fabric Core n’est pas dépendant de l’infrastructure.</span><span class="sxs-lookup"><span data-stu-id="17ba5-112">Fabric core is framework independent.</span></span> <span data-ttu-id="17ba5-113">Fabric Core est utilisé par et inclus avec Fabric React.</span><span class="sxs-lookup"><span data-stu-id="17ba5-113">Fabric Core is used by and included with Fabric React.</span></span>

<span data-ttu-id="17ba5-114">Pour commencer à utiliser Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="17ba5-114">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="17ba5-115">Ajoutez la référence CDN au code HTML sur votre page.</span><span class="sxs-lookup"><span data-stu-id="17ba5-115">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="17ba5-116">Utilisez les polices et les icônes Fabric.</span><span class="sxs-lookup"><span data-stu-id="17ba5-116">Use Fabric icons and fonts.</span></span>

    <span data-ttu-id="17ba5-p104">Pour utiliser une icône Fabric, incluez l’élément « i » sur votre page, puis référencez les classes appropriées. Vous pouvez contrôler la taille de l’icône en modifiant la taille de police. Par exemple, le code suivant montre comment créer une icône de tableau extra large qui utilise la couleur themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="17ba5-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="17ba5-p105">Pour rechercher des icônes supplémentaires disponibles dans Office UI Fabric, utilisez la fonctionnalité de recherche de la page [Icônes](https://developer.microsoft.com/fabric#/styles/icons). Lorsque vous trouvez une icône à utiliser dans votre complément, veillez à précéder le nom de l’icône de `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="17ba5-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="17ba5-122">Pour plus d’informations sur les tailles de police et les couleurs disponibles dans Office UI Fabric, voir [Typographie](https://developer.microsoft.com/fabric#/styles/typography) et [Couleurs](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="17ba5-122">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

## <a name="use-fabric-components"></a><span data-ttu-id="17ba5-123">Utiliser les composants Fabric</span><span class="sxs-lookup"><span data-stu-id="17ba5-123">Use Fabric Components</span></span>

<span data-ttu-id="17ba5-124">Fabric fournit une variété de composants UX que vous pouvez utiliser pour créer votre add-in.</span><span class="sxs-lookup"><span data-stu-id="17ba5-124">Fabric provides a variety of UX components that you can use to build your add-in.</span></span> <span data-ttu-id="17ba5-125">Nous ne nous attendons pas à ce que tous les composants fabric soient utilisés par un seul et même composant.</span><span class="sxs-lookup"><span data-stu-id="17ba5-125">We do not expect that all fabric components will be used by a single add-in.</span></span> <span data-ttu-id="17ba5-126">Déterminez les meilleurs composants pour votre scénario et l’expérience utilisateur [](https://developer.microsoft.com/fabric#/components/breadcrumb) (par exemple, il peut être difficile d’afficher correctement une vue d’accès dans le volet Des tâches).</span><span class="sxs-lookup"><span data-stu-id="17ba5-126">Determine the best components for your scenario and user experience (for example, it may be hard to properly display a [Breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) in the task pane).</span></span>

<span data-ttu-id="17ba5-127">Voici une liste des composants [UX courants](https://developer.microsoft.com/fluentui#/controls/web) de Fabric React que nous vous recommandons d’utiliser dans un add-in :</span><span class="sxs-lookup"><span data-stu-id="17ba5-127">The following is a list of common [Fabric React UX components](https://developer.microsoft.com/fluentui#/controls/web) that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="17ba5-128">Bouton</span><span class="sxs-lookup"><span data-stu-id="17ba5-128">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="17ba5-129">Case à cocher</span><span class="sxs-lookup"><span data-stu-id="17ba5-129">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="17ba5-130">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="17ba5-130">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="17ba5-131">Liste déroulante</span><span class="sxs-lookup"><span data-stu-id="17ba5-131">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="17ba5-132">Étiquette</span><span class="sxs-lookup"><span data-stu-id="17ba5-132">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="17ba5-133">Liste</span><span class="sxs-lookup"><span data-stu-id="17ba5-133">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="17ba5-134">Tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="17ba5-134">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="17ba5-135">TextField</span><span class="sxs-lookup"><span data-stu-id="17ba5-135">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="17ba5-136">Bouton bascule</span><span class="sxs-lookup"><span data-stu-id="17ba5-136">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="17ba5-p107">Vous pouvez utiliser différentes infrastructures JavaScript, comme Angular ou React, pour créer votre complément. Pour commencer à utiliser les composants Fabric avec votre infrastructure, consultez les ressources suivantes.</span><span class="sxs-lookup"><span data-stu-id="17ba5-p107">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="17ba5-139">**Infrastructure**</span><span class="sxs-lookup"><span data-stu-id="17ba5-139">**Framework**</span></span>|<span data-ttu-id="17ba5-140">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="17ba5-140">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="17ba5-141">**React**</span><span class="sxs-lookup"><span data-stu-id="17ba5-141">**React**</span></span>|[<span data-ttu-id="17ba5-142">Utilisation d’Office UI Fabric React dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="17ba5-142">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
