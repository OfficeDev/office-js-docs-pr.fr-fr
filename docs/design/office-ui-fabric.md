---
title: Office UI Fabric dans des compléments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 04964d5864eea4a960f7b57e5df6f7bd7c844fde
ms.sourcegitcommit: 4e4f7c095e8f33b06bd8a02534ee901125eb1d17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/28/2018
ms.locfileid: "20084069"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="4c71d-102">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="4c71d-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="4c71d-p101">Office UI Fabric est une infrastructure frontale JavaScript permettant de créer des expériences pour Office et Office 365. Fabric propose des composants axés sur des visuels que vous pouvez étendre, retravailler et utiliser dans votre complément Office. Fabric utilisant le langage de création d’Office, ses composants d’expérience utilisateur ressemblent à une extension naturelle d’Office.</span><span class="sxs-lookup"><span data-stu-id="4c71d-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="4c71d-p102">Si vous créez un complément, nous vous encourageons à utiliser Office UI Fabric pour mettre au point l’expérience utilisateur. L’utilisation d’Office UI Fabric est facultative.</span><span class="sxs-lookup"><span data-stu-id="4c71d-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="4c71d-108">Les sections suivantes expliquent comment commencer à utiliser Fabric en fonction de vos besoins.</span><span class="sxs-lookup"><span data-stu-id="4c71d-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="4c71d-109">Utiliser Fabric Core : icônes, polices, couleurs</span><span class="sxs-lookup"><span data-stu-id="4c71d-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="4c71d-p103">Fabric Core contient les principaux éléments du langage de création tels que les icônes, les couleurs, le type et la grille. Fabric Core n’est pas dépendant de l’infrastructure. Les composants JS et React de la structure utilisent Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="4c71d-p103">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.</span></span>

<span data-ttu-id="4c71d-113">Pour commencer à utiliser Fabric Core :</span><span class="sxs-lookup"><span data-stu-id="4c71d-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="4c71d-114">Ajoutez la référence CDN au code HTML sur votre page.</span><span class="sxs-lookup"><span data-stu-id="4c71d-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="4c71d-115">Utilisez les polices et les icônes Fabric.</span><span class="sxs-lookup"><span data-stu-id="4c71d-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="4c71d-p104">Pour utiliser une icône Fabric, incluez l’élément « i » sur votre page, puis référencez les classes appropriées. Vous pouvez contrôler la taille de l’icône en modifiant la taille de police. Par exemple, le code suivant montre comment créer une icône de tableau extra large qui utilise la couleur themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="4c71d-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="4c71d-p105">Pour rechercher des icônes supplémentaires disponibles dans Office UI Fabric, utilisez la fonctionnalité de recherche de la page [Icônes](https://dev.office.com/fabric#/styles/icons). Lorsque vous trouvez une icône à utiliser dans votre complément, veillez à précéder le nom de l’icône de `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="4c71d-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="4c71d-121">Pour plus d’informations sur les tailles de police et les couleurs disponibles dans Office UI Fabric, voir [Typographie](https://dev.office.com/fabric#/styles/typography) et [Couleurs](https://dev.office.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="4c71d-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="4c71d-122">Utiliser les composants Fabric</span><span class="sxs-lookup"><span data-stu-id="4c71d-122">Use Fabric Components</span></span> 
<span data-ttu-id="4c71d-123">Fabric fournit une variété de composants UX que vous pouvez utiliser pour créer votre complément, y compris les types de composants suivants :</span><span class="sxs-lookup"><span data-stu-id="4c71d-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="4c71d-124">Composants d’entrée - par exemple, bouton, case à cocher et bouton bascule</span><span class="sxs-lookup"><span data-stu-id="4c71d-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="4c71d-125">Composants de navigation - pivot et Breadcrumb à titre d'exemples</span><span class="sxs-lookup"><span data-stu-id="4c71d-125">Navigation components - for example, Pivot Breadcrumb</span></span>
- <span data-ttu-id="4c71d-126">Composants de notification - par exemple, MessageBar et légende</span><span class="sxs-lookup"><span data-stu-id="4c71d-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="4c71d-127">Tous les composants ne sont pas recommandés pour une dans les compléments. Voici une liste des composants Fabric React UX que nous vous recommandons d'utiliser dans un complément :</span><span class="sxs-lookup"><span data-stu-id="4c71d-127">Not all Fabric components are recommended for use in add-ins. Here is a list of Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="4c71d-128">Breadcrumb</span><span class="sxs-lookup"><span data-stu-id="4c71d-128">Breadcrumb</span></span>](https://developer.microsoft.com/en-us/fabric#/components/breadcrumb)
- [<span data-ttu-id="4c71d-129">Bouton</span><span class="sxs-lookup"><span data-stu-id="4c71d-129">Button</span></span>](https://developer.microsoft.com/en-us/fabric#/components/button)
- [<span data-ttu-id="4c71d-130">Case à cocher</span><span class="sxs-lookup"><span data-stu-id="4c71d-130">Checkbox</span></span>](https://developer.microsoft.com/en-us/fabric#/components/checkbox)
- [<span data-ttu-id="4c71d-131">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="4c71d-131">ChoiceGroup</span></span>](https://developer.microsoft.com/en-us/fabric#/components/choicegroup)
- [<span data-ttu-id="4c71d-132">Liste déroulante</span><span class="sxs-lookup"><span data-stu-id="4c71d-132">Dropdown</span></span>](https://developer.microsoft.com/en-us/fabric#/components/dropdown)
- [<span data-ttu-id="4c71d-133">Étiquette</span><span class="sxs-lookup"><span data-stu-id="4c71d-133">Label</span></span>](https://developer.microsoft.com/en-us/fabric#/components/label)
- [<span data-ttu-id="4c71d-134">Liste</span><span class="sxs-lookup"><span data-stu-id="4c71d-134">List</span></span>](https://developer.microsoft.com/en-us/fabric#/components/list)
- [<span data-ttu-id="4c71d-135">Tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="4c71d-135">Pivot</span></span>](https://developer.microsoft.com/en-us/fabric#/components/pivot)
- [<span data-ttu-id="4c71d-136">TextField</span><span class="sxs-lookup"><span data-stu-id="4c71d-136">TextField</span></span>](https://developer.microsoft.com/en-us/fabric#/components/textfield)
- [<span data-ttu-id="4c71d-137">Bascule</span><span class="sxs-lookup"><span data-stu-id="4c71d-137">Toggle</span></span>](https://developer.microsoft.com/en-us/fabric#/components/toggle)

<span data-ttu-id="4c71d-p106">Vous pouvez utiliser différentes infrastructures JavaScript, comme Angular ou React, pour créer votre complément. Pour commencer à utiliser les composants Fabric avec votre infrastructure, consultez les ressources suivantes.</span><span class="sxs-lookup"><span data-stu-id="4c71d-p106">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="4c71d-140">**Infrastructure**</span><span class="sxs-lookup"><span data-stu-id="4c71d-140">**Framework**</span></span>|<span data-ttu-id="4c71d-141">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="4c71d-141">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="4c71d-142">**React**</span><span class="sxs-lookup"><span data-stu-id="4c71d-142">**React**</span></span>|[<span data-ttu-id="4c71d-143">Utilisation d’Office UI Fabric React dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="4c71d-143">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="4c71d-144">**Angular**</span><span class="sxs-lookup"><span data-stu-id="4c71d-144">**Angular**</span></span>| <span data-ttu-id="4c71d-145">Reportez-vous à [ngOfficeUIFabric](http://ngofficeuifabric.com/), qui est un projet communautaire avec des directives Angular 1.5, et [envisagez d’insérer des composants Fabric dans des composants Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components).</span><span class="sxs-lookup"><span data-stu-id="4c71d-145">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
