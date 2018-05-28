---
title: Office UI Fabric dans des compl?ments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 8fafe8a68c477868c12bff61c7f9ff23fc7314e0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="f6973-102">Office UI Fabric dans des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="f6973-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="f6973-p101">Office UI Fabric est une infrastructure frontale JavaScript permettant de cr?er des exp?riences pour Office et Office 365. Fabric propose des composants ax?s sur des visuels que vous pouvez ?tendre, retravailler et utiliser dans votre compl?ment Office. Fabric utilisant le langage de cr?ation d?Office, ses composants d?exp?rience utilisateur ressemblent ? une extension naturelle d?Office.</span><span class="sxs-lookup"><span data-stu-id="f6973-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="f6973-p102">Si vous cr?ez un compl?ment, nous vous encourageons ? utiliser Office UI Fabric pour mettre au point l?exp?rience utilisateur. L?utilisation d?Office UI Fabric est facultative.</span><span class="sxs-lookup"><span data-stu-id="f6973-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="f6973-108">Les sections suivantes expliquent comment commencer ? utiliser Fabric en fonction de vos besoins.</span><span class="sxs-lookup"><span data-stu-id="f6973-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="f6973-109">Utiliser Fabric Core : ic?nes, polices, couleurs</span><span class="sxs-lookup"><span data-stu-id="f6973-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="f6973-p103">Fabric Core contient les principaux ?l?ments du langage de cr?ation tels que les ic?nes, les couleurs, le type et la grille. Fabric Core n?est pas d?pendant de l?infrastructure. Les composants JS et React de la structure utilisent Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="f6973-p103">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.</span></span>

<span data-ttu-id="f6973-113">Pour commencer ? utiliser Fabric Core :</span><span class="sxs-lookup"><span data-stu-id="f6973-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="f6973-114">Ajoutez la r?f?rence CDN au code HTML sur votre page.</span><span class="sxs-lookup"><span data-stu-id="f6973-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="f6973-115">Utilisez les polices et les ic?nes Fabric.</span><span class="sxs-lookup"><span data-stu-id="f6973-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="f6973-p104">Pour utiliser une ic?ne Fabric, incluez l??l?ment ? i ? sur votre page, puis r?f?rencez les classes appropri?es. Vous pouvez contr?ler la taille de l?ic?ne en modifiant la taille de police. Par exemple, le code suivant montre comment cr?er une ic?ne de tableau extra large qui utilise la couleur themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="f6973-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="f6973-p105">Pour rechercher des ic?nes suppl?mentaires disponibles dans Office UI Fabric, utilisez la fonctionnalit? de recherche de la page [Ic?nes](https://dev.office.com/fabric#/styles/icons). Lorsque vous trouvez une ic?ne ? utiliser dans votre compl?ment, veillez ? pr?c?der le nom de l?ic?ne de `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="f6973-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="f6973-121">Pour plus d?informations sur les tailles de police et les couleurs disponibles dans Office UI Fabric, voir [Typographie](https://dev.office.com/fabric#/styles/typography) et [Couleurs](https://dev.office.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="f6973-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="f6973-122">Utiliser les composants Fabric</span><span class="sxs-lookup"><span data-stu-id="f6973-122">Use Fabric Components</span></span> 
<span data-ttu-id="f6973-123">Fabric fournit une vari?t? de composants UX que vous pouvez utiliser pour cr?er votre compl?ment, y compris les types de composants suivants :</span><span class="sxs-lookup"><span data-stu-id="f6973-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="f6973-124">Composants d?entr?e - par exemple, bouton, case ? cocher et bouton bascule</span><span class="sxs-lookup"><span data-stu-id="f6973-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="f6973-125">Composants de navigation - par exemple, tableau crois? dynamique, barre de navigation</span><span class="sxs-lookup"><span data-stu-id="f6973-125">Navigation components - for example, Pivot Breadcrumb</span></span>
- <span data-ttu-id="f6973-126">Composants de notification - par exemple, MessageBar et l?gende</span><span class="sxs-lookup"><span data-stu-id="f6973-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="f6973-p106">Il n?est pas recommand? d?utiliser tous les composants Fabric dans des compl?ments. Nous fournissons des conseils sur l?utilisation des composants recommand?s dans cette section. Par exemple, pour savoir comment utiliser un bouton Fabric dans votre compl?ment, voir [Bouton](button.md).</span><span class="sxs-lookup"><span data-stu-id="f6973-p106">Not all Fabric components are recommended for use in add-ins. We provide guidance for how you can use the recommended components in this section. For example, for guidance for using a Fabric button in your add-in, see [Button](button.md).</span></span> 

<span data-ttu-id="f6973-p107">Vous pouvez utiliser diff?rentes infrastructures JavaScript, comme Angular ou React, pour cr?er votre compl?ment. Pour commencer ? utiliser les composants Fabric avec votre infrastructure, consultez les ressources suivantes.</span><span class="sxs-lookup"><span data-stu-id="f6973-p107">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="f6973-131">**Infrastructure**</span><span class="sxs-lookup"><span data-stu-id="f6973-131">**Framework**</span></span>|<span data-ttu-id="f6973-132">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="f6973-132">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="f6973-133">**React**</span><span class="sxs-lookup"><span data-stu-id="f6973-133">**React**</span></span>|[<span data-ttu-id="f6973-134">Utilisation d?Office UI Fabric React dans des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="f6973-134">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="f6973-135">**Angular**</span><span class="sxs-lookup"><span data-stu-id="f6973-135">**Angular**</span></span>| <span data-ttu-id="f6973-136">Reportez-vous ? [ngOfficeUIFabric](http://ngofficeuifabric.com/), qui est un projet communautaire avec des directives Angular 1.5, et [envisagez d?ins?rer des composants Fabric dans des composants Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components).</span><span class="sxs-lookup"><span data-stu-id="f6973-136">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
