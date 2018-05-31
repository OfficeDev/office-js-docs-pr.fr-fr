---
title: Office UI Fabric dans des compléments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 8fafe8a68c477868c12bff61c7f9ff23fc7314e0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437366"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="eedb4-102">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="eedb4-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="eedb4-p101">Office UI Fabric est une infrastructure frontale JavaScript permettant de créer des expériences pour Office et Office 365. Fabric propose des composants axés sur des visuels que vous pouvez étendre, retravailler et utiliser dans votre complément Office. Fabric utilisant le langage de création d’Office, ses composants d’expérience utilisateur ressemblent à une extension naturelle d’Office.</span><span class="sxs-lookup"><span data-stu-id="eedb4-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="eedb4-p102">Si vous créez un complément, nous vous encourageons à utiliser Office UI Fabric pour mettre au point l’expérience utilisateur. L’utilisation d’Office UI Fabric est facultative.</span><span class="sxs-lookup"><span data-stu-id="eedb4-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="eedb4-108">Les sections suivantes expliquent comment commencer à utiliser Fabric en fonction de vos besoins.</span><span class="sxs-lookup"><span data-stu-id="eedb4-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="eedb4-109">Utiliser Fabric Core : icônes, polices, couleurs</span><span class="sxs-lookup"><span data-stu-id="eedb4-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="eedb4-p103">Fabric Core contient les principaux éléments du langage de création tels que les icônes, les couleurs, le type et la grille. Fabric Core n’est pas dépendant de l’infrastructure. Les composants JS et React de la structure utilisent Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="eedb4-p103">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.</span></span>

<span data-ttu-id="eedb4-113">Pour commencer à utiliser Fabric Core :</span><span class="sxs-lookup"><span data-stu-id="eedb4-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="eedb4-114">Ajoutez la référence CDN au code HTML sur votre page.</span><span class="sxs-lookup"><span data-stu-id="eedb4-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="eedb4-115">Utilisez les polices et les icônes Fabric.</span><span class="sxs-lookup"><span data-stu-id="eedb4-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="eedb4-p104">Pour utiliser une icône Fabric, incluez l’élément « i » sur votre page, puis référencez les classes appropriées. Vous pouvez contrôler la taille de l’icône en modifiant la taille de police. Par exemple, le code suivant montre comment créer une icône de tableau extra large qui utilise la couleur themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="eedb4-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="eedb4-p105">Pour rechercher des icônes supplémentaires disponibles dans Office UI Fabric, utilisez la fonctionnalité de recherche de la page [Icônes](https://dev.office.com/fabric#/styles/icons). Lorsque vous trouvez une icône à utiliser dans votre complément, veillez à précéder le nom de l’icône de `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="eedb4-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="eedb4-121">Pour plus d’informations sur les tailles de police et les couleurs disponibles dans Office UI Fabric, voir [Typographie](https://dev.office.com/fabric#/styles/typography) et [Couleurs](https://dev.office.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="eedb4-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="eedb4-122">Utiliser les composants Fabric</span><span class="sxs-lookup"><span data-stu-id="eedb4-122">Use Fabric Components</span></span> 
<span data-ttu-id="eedb4-123">Fabric fournit une variété de composants UX que vous pouvez utiliser pour créer votre complément, y compris les types de composants suivants :</span><span class="sxs-lookup"><span data-stu-id="eedb4-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="eedb4-124">Composants d’entrée - par exemple, bouton, case à cocher et bouton bascule</span><span class="sxs-lookup"><span data-stu-id="eedb4-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="eedb4-125">Composants de navigation - par exemple, tableau croisé dynamique, barre de navigation</span><span class="sxs-lookup"><span data-stu-id="eedb4-125">Navigation components - for example, Pivot Breadcrumb</span></span>
- <span data-ttu-id="eedb4-126">Composants de notification - par exemple, MessageBar et légende</span><span class="sxs-lookup"><span data-stu-id="eedb4-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="eedb4-p106">Il n’est pas recommandé d’utiliser tous les composants Fabric dans des compléments. Nous fournissons des conseils sur l’utilisation des composants recommandés dans cette section. Par exemple, pour savoir comment utiliser un bouton Fabric dans votre complément, voir [Bouton](button.md).</span><span class="sxs-lookup"><span data-stu-id="eedb4-p106">Not all Fabric components are recommended for use in add-ins. We provide guidance for how you can use the recommended components in this section. For example, for guidance for using a Fabric button in your add-in, see [Button](button.md).</span></span> 

<span data-ttu-id="eedb4-p107">Vous pouvez utiliser différentes infrastructures JavaScript, comme Angular ou React, pour créer votre complément. Pour commencer à utiliser les composants Fabric avec votre infrastructure, consultez les ressources suivantes.</span><span class="sxs-lookup"><span data-stu-id="eedb4-p107">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="eedb4-131">**Infrastructure**</span><span class="sxs-lookup"><span data-stu-id="eedb4-131">**Framework**</span></span>|<span data-ttu-id="eedb4-132">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="eedb4-132">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="eedb4-133">**React**</span><span class="sxs-lookup"><span data-stu-id="eedb4-133">**React**</span></span>|[<span data-ttu-id="eedb4-134">Utilisation d’Office UI Fabric React dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="eedb4-134">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="eedb4-135">**Angular**</span><span class="sxs-lookup"><span data-stu-id="eedb4-135">**Angular**</span></span>| <span data-ttu-id="eedb4-136">Reportez-vous à [ngOfficeUIFabric](http://ngofficeuifabric.com/), qui est un projet communautaire avec des directives Angular 1.5, et [envisagez d’insérer des composants Fabric dans des composants Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components).</span><span class="sxs-lookup"><span data-stu-id="eedb4-136">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
