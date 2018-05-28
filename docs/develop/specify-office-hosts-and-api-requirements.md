---
title: Sp?cification des exigences en mati?re d?h?tes Office et d?API
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd517dee1faf8d3f3009a0b9ce7127f5760e730d
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="specify-office-hosts-and-api-requirements"></a><span data-ttu-id="705cb-102">Sp?cification des exigences en mati?re d?h?tes Office et d?API</span><span class="sxs-lookup"><span data-stu-id="705cb-102">Specify Office hosts and API requirements</span></span>

<span data-ttu-id="705cb-p101">Il se peut que votre compl?ment Office d?pende d?un h?te Office sp?cifique, d?un ensemble de conditions requises, d?un membre d?API ou d?une version de l?API pour fonctionner correctement. Par exemple, votre compl?ment peut :</span><span class="sxs-lookup"><span data-stu-id="705cb-p101">Your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API in order to work as expected. For example, your add-in might:</span></span>

- <span data-ttu-id="705cb-105">ex?cuter une ou plusieurs application Office (Word ou Excel) ;</span><span class="sxs-lookup"><span data-stu-id="705cb-105">Run in a single Office application (Word or Excel), or several applications.</span></span>
    
- <span data-ttu-id="705cb-p102">utiliser des API JavaScript disponibles uniquement dans certaines versions d?Office. Par exemple, vous pouvez utiliser les API JavaScript d?Excel dans un compl?ment qui fonctionne dans Excel 2016 ;</span><span class="sxs-lookup"><span data-stu-id="705cb-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span> 
    
- <span data-ttu-id="705cb-108">s?ex?cuter uniquement dans les versions d?Office qui prennent en charge les membres d?API utilis?s par votre compl?ment.</span><span class="sxs-lookup"><span data-stu-id="705cb-108">Run only in versions of Office that support API members that your add-in uses.</span></span>
    
<span data-ttu-id="705cb-109">Cet article vous aidera ? comprendre les options que vous devez choisir afin de vous assurer que votre compl?ment fonctionne comme pr?vu et atteint l?audience la plus large possible.</span><span class="sxs-lookup"><span data-stu-id="705cb-109">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="705cb-110">Pour savoir de mani?re d?taill?e quelle version d?Office prend en charge les compl?ments Office, consultez la page relative ? la [disponibilit? des compl?ments Office sur les plateformes et les h?tes](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="705cb-110">For a high-level view of where Office Add-ins are currently supported, see the [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span> 

<span data-ttu-id="705cb-111">Le tableau suivant r?pertorie les concepts principaux d?crits dans cet article.</span><span class="sxs-lookup"><span data-stu-id="705cb-111">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="705cb-112">**Concept**</span><span class="sxs-lookup"><span data-stu-id="705cb-112">**Concept**</span></span>|<span data-ttu-id="705cb-113">**Description**</span><span class="sxs-lookup"><span data-stu-id="705cb-113">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="705cb-114">Application Office, application h?te Office ou h?te Office</span><span class="sxs-lookup"><span data-stu-id="705cb-114">Office application, Office host application, Office host, or host</span></span>|<span data-ttu-id="705cb-p103">Application Office utilis?e pour ex?cuter votre compl?ment. Par exemple, Word, Word Online ou Excel.</span><span class="sxs-lookup"><span data-stu-id="705cb-p103">The Office application used to run your add-in. For example, Word, Word Online, Excel, and so on.</span></span>|
|<span data-ttu-id="705cb-117">Plateforme</span><span class="sxs-lookup"><span data-stu-id="705cb-117">Platform</span></span>|<span data-ttu-id="705cb-118">Application sur laquelle l?h?te Office est ex?cut?, comme Office Online ou Office pour iPad.</span><span class="sxs-lookup"><span data-stu-id="705cb-118">Where the Office host runs, such as Office Online or Office for iPad.</span></span>|
|<span data-ttu-id="705cb-119">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="705cb-119">Requirement set</span></span>|<span data-ttu-id="705cb-p104">Groupe nomm? de membres d?API associ?s. Les compl?ments utilisent des ensembles de conditions requises pour d?terminer si l?h?te Office prend en charge les membres d?API utilis?s par votre compl?ment. Il est plus facile de tester la prise en charge d?un ensemble de conditions requises, plut?t que la prise en charge de membres individuels d?API. La prise en charge de l?ensemble des conditions requises varie selon l?h?te Office et la version de ce dernier. </span><span class="sxs-lookup"><span data-stu-id="705cb-p104">A named group of related API members. Add-ins use requirement sets to determine whether the Office host supports API members used by your add-in. It's easier to test for the support of a requirement set than for the support of individual API members. Requirement set support varies by Office host and the version of the Office host. </span></span><br ><span data-ttu-id="705cb-p105">Les ensembles de conditions requises sont sp?cifi?s dans le fichier manifeste. Quand vous d?finissez des ensembles de conditions requises dans le fichier manifeste, vous d?finissez le niveau minimal de prise en charge de l?API que l?h?te Office doit fournir pour ex?cuter votre compl?ment. Les h?tes Office qui ne prennent pas en charge les ensembles de conditions requises sp?cifi?s dans le manifeste ne peuvent pas ex?cuter votre compl?ment, et votre compl?ment ne sera pas affich? dans <span class="ui">Mes compl?ments</span>. Cela limite les emplacements o? votre compl?ment sera disponible. Dans le code utilisant les v?rifications ? l?ex?cution. Pour obtenir la liste compl?te des ensembles de conditions requises, voir [Ensemble de conditions requises pour les compl?ments Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="705cb-p105">Requirement sets are specified in the manifest file. When you specify requirement sets in the manifest, you set the minimum level of API support that the Office host must provide in order to run your add-in. Office hosts that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.In code using runtime checks. For the complete list of requirement sets, see [Office add-in requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>|
|<span data-ttu-id="705cb-128">V?rification ? l?ex?cution</span><span class="sxs-lookup"><span data-stu-id="705cb-128">Runtime check</span></span>|<span data-ttu-id="705cb-p106">Test effectu? ? l?ex?cution pour d?terminer si l?h?te Office qui ex?cute votre compl?ment prend en charge les ensembles de conditions requises ou les m?thodes utilis?s par votre compl?ment. Pour effectuer une v?rification ? l?ex?cution, vous pouvez utiliser une instruction **if** avec la m?thode **isSetSupported**, les ensembles de conditions requises ou les noms de m?thode qui ne font pas partie d?un ensemble de conditions requises. Les v?rifications ? l?ex?cution permettent de veiller ? ce que votre compl?ment atteigne le plus grand nombre possible de clients. Contrairement aux ensembles de conditions requises, les v?rifications ? l?ex?cution ne pr?cisent pas le niveau minimal de prise en charge de l?API que l?h?te Office doit fournir pour l?ex?cution de votre compl?ment. Au lieu de cela, vous devez utiliser l?instruction **if** afin de d?terminer si un membre d?API est pris en charge. Si c?est le cas, vous pouvez fournir des fonctionnalit?s suppl?mentaires dans votre compl?ment. Votre compl?ment s?affiche toujours dans **Mes compl?ments** quand vous effectuez des v?rifications ? l?ex?cution.</span><span class="sxs-lookup"><span data-stu-id="705cb-p106">A test that is performed at runtime to determine whether the Office host running your add-in supports requirement sets or methods used by your add-in. To perform a runtime check, you use an  **if** statement with the **isSetSupported** method, the requirement sets, or the method names that aren't part of a requirement set.Use runtime checks to ensure that your add-in reaches the broadest number of customers. Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office host must provide for your add-in to run. Instead, you use the  **if** statement to determine whether an API member is supported. If it is, you can provide additional functionality in your add-in. Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="705cb-135">Avant de commencer</span><span class="sxs-lookup"><span data-stu-id="705cb-135">Before you begin</span></span>

<span data-ttu-id="705cb-p107">Votre compl?ment doit utiliser la version la plus r?cente du sch?ma de manifeste de compl?ment. Si vous utilisez les v?rifications ? l?ex?cution dans votre compl?ment, assurez-vous que vous utilisez la derni?re API JavaScript pour la biblioth?que Office (office.js).</span><span class="sxs-lookup"><span data-stu-id="705cb-p107">Your add-in must use the most current version of the add-in manifest schema. If you use runtime checks in your add-in, ensure that you use the latest JavaScript API for Office (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="705cb-138">Indication du sch?ma de manifeste de compl?ment le plus r?cent</span><span class="sxs-lookup"><span data-stu-id="705cb-138">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="705cb-p108">Le manifeste de votre du compl?ment doit utiliser la version 1.1 du sch?ma de manifeste de compl?ment. D?finissez l??l?ment **App_office** dans votre manifeste compl?ment comme suit.</span><span class="sxs-lookup"><span data-stu-id="705cb-p108">Your add-in's manifest must use version 1.1 of the add-in manifest schema. Set the  **OfficeApp** element in your add-in manifest as follows.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-javascript-api-for-office-library"></a><span data-ttu-id="705cb-141">Indication de l?API JavaScript la plus r?cente pour la biblioth?que Office</span><span class="sxs-lookup"><span data-stu-id="705cb-141">Specify the latest JavaScript API for Office library</span></span>

<span data-ttu-id="705cb-p109">Si vous utilisez des v?rifications ? l?ex?cution, r?f?rencez la version la plus r?cente de l?API JavaScript pour la biblioth?que Office ? partir du r?seau de livraison de contenu (CDN). Pour ce faire, ajoutez la balise `script` suivante ? votre code HTML. L?utilisation de `/1/` dans l?URL CDN garantit que vous r?f?rencez la version d?Office.js la plus r?cente.</span><span class="sxs-lookup"><span data-stu-id="705cb-p109">If you use runtime checks, reference the most current version of the JavaScript API for Office library from the content delivery network (CDN). To do this, add the following  `script` tag to your HTML. Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a><span data-ttu-id="705cb-145">Options pour sp?cifier des h?tes Office ou les conditions requises d?API</span><span class="sxs-lookup"><span data-stu-id="705cb-145">Options to specify Office hosts or API requirements</span></span>

<span data-ttu-id="705cb-p110">Lors de la sp?cification des h?tes Office ou des conditions requises d?API, vous devez tenir compte de plusieurs facteurs. Le diagramme suivant montre comment choisir la technique ? utiliser dans votre compl?ment.</span><span class="sxs-lookup"><span data-stu-id="705cb-p110">When you specify Office hosts or API requirements, there are several factors to consider. The following diagram shows how to decide which technique to use in your add-in.</span></span>

![Optez pour la meilleure solution pour votre compl?ment lorsque vous sp?cifiez des h?tes Office ou des exigences d?API](../images/options-for-office-hosts.png)

- <span data-ttu-id="705cb-p111">Si votre compl?ment s?ex?cute dans un h?te Office, d?finissez l??l?ment **Hosts** dans le manifeste. Pour plus d?informations, consultez [D?finition de l??l?ment Hosts](#set-the-hosts-element).</span><span class="sxs-lookup"><span data-stu-id="705cb-p111">If your add-in runs in one Office host, set the **Hosts** element in the manifest. For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>
    
- <span data-ttu-id="705cb-p112">Pour d?finir l?ensemble minimal de conditions requises ou les membres minimaux d?API qu?un h?te Office doit prendre en charge pour ex?cuter votre compl?ment, d?finissez l??l?ment **Requirements** dans le manifeste. Pour plus d?informations, consultez la section [ D?finition de l??l?ment Requirements dans le manifeste](#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="705cb-p112">To set the minimum requirement set or API members that an Office host must support to run your add-in, set the  **Requirements** element in the manifest. For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>
    
- <span data-ttu-id="705cb-p113">Si vous souhaitez proposer des fonctionnalit?s suppl?mentaires lorsque des ensembles de conditions requises sp?cifiques ou des membres d?API sont disponibles dans l?h?te Office, effectuez une v?rification ? l?ex?cution dans le code JavaScript de votre compl?ment. Par exemple, si votre compl?ment est ex?cut? dans Excel 2016, utilisez les membres d?API de la nouvelle API JavaScript pour Excel pour fournir des fonctionnalit?s suppl?mentaires. Pour plus d?informations, consultez la section [Utilisation des v?rifications ? l?ex?cution dans votre code JavaScript](#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="705cb-p113">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office host, perform a runtime check in your add-in's JavaScript code. For example, if your add-in runs in Excel 2016, use API members from the new JavaScript API for Excel to provide additional functionality. For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>
    
## <a name="set-the-hosts-element"></a><span data-ttu-id="705cb-156">D?finition de l??l?ment Hosts</span><span class="sxs-lookup"><span data-stu-id="705cb-156">Set the Hosts element</span></span>

<span data-ttu-id="705cb-p114">Pour ex?cuter votre compl?ment dans une application h?te Office, utilisez les ?l?ments **Hosts** et **Host** dans le manifeste. Si vous ne d?finissez pas l??l?ment **Hosts**, votre compl?ment sera ex?cut? dans tous les h?tes.</span><span class="sxs-lookup"><span data-stu-id="705cb-p114">To make your add-in run in one Office host application, use the  **Hosts** and **Host** elements in the manifest. If you don't specify the **Hosts** element, your add-in will run in all hosts.</span></span>

<span data-ttu-id="705cb-159">Par exemple, les d?clarations  **Hosts** et **Host** suivantes indiquent que le compl?ment fonctionnera avec n?importe quelle version d?Excel, y compris Excel pour Windows, Excel Online et Excel pour iPad.</span><span class="sxs-lookup"><span data-stu-id="705cb-159">For example, the following  **Hosts** and **Host** declaration specifies that the add-in will work with any release of Excel, which includes Excel for Windows, Excel Online, and Excel for iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="705cb-p115">L??l?ment  **Hosts** peut contenir un ou plusieurs ?l?ments  **Host**. L??l?ment  **Host** indique l?h?te Office requis par votre compl?ment. L?attribut **Name** est requis et peut ?tre d?fini sur l?une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="705cb-p115">The  **Hosts** element can contain one or more **Host** elements. The **Host** element specifies the Office host your add-in requires. The **Name** attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="705cb-163">Nom</span><span class="sxs-lookup"><span data-stu-id="705cb-163">Name</span></span>          | <span data-ttu-id="705cb-164">Applications h?tes Office</span><span class="sxs-lookup"><span data-stu-id="705cb-164">Office host applications</span></span>                      |
|:--------------|:----------------------------------------------|
| <span data-ttu-id="705cb-165">Base de donn?es</span><span class="sxs-lookup"><span data-stu-id="705cb-165">Database</span></span>      | <span data-ttu-id="705cb-166">applications web Access</span><span class="sxs-lookup"><span data-stu-id="705cb-166">Access web apps</span></span>                               |
| <span data-ttu-id="705cb-167">Document</span><span class="sxs-lookup"><span data-stu-id="705cb-167">Document</span></span>      | <span data-ttu-id="705cb-168">Word pour Windows, Mac, iPad et Online</span><span class="sxs-lookup"><span data-stu-id="705cb-168">Word for Windows, Mac, iPad and Online</span></span>        |
| <span data-ttu-id="705cb-169">Bo?te aux lettres</span><span class="sxs-lookup"><span data-stu-id="705cb-169">Mailbox</span></span>       | <span data-ttu-id="705cb-170">Outlook pour Windows, Mac, Web et Outlook.com</span><span class="sxs-lookup"><span data-stu-id="705cb-170">Outlook for Windows, Mac, Web and Outlook.com</span></span> | 
| <span data-ttu-id="705cb-171">Pr?sentation</span><span class="sxs-lookup"><span data-stu-id="705cb-171">Presentation</span></span>  | <span data-ttu-id="705cb-172">PowerPoint pour Windows, Mac, iPad et Online</span><span class="sxs-lookup"><span data-stu-id="705cb-172">PowerPoint for Windows, Mac, iPad and Online</span></span>  |
| <span data-ttu-id="705cb-173">Projet</span><span class="sxs-lookup"><span data-stu-id="705cb-173">Project</span></span>       | <span data-ttu-id="705cb-174">Projet</span><span class="sxs-lookup"><span data-stu-id="705cb-174">Project</span></span>                                       |
| <span data-ttu-id="705cb-175">Classeur</span><span class="sxs-lookup"><span data-stu-id="705cb-175">Workbook</span></span>      | <span data-ttu-id="705cb-176">Excel pour Windows, Mac, iPad et Online</span><span class="sxs-lookup"><span data-stu-id="705cb-176">Excel Windows, Mac, iPad and Online</span></span>           |

> [!NOTE]
> <span data-ttu-id="705cb-p116">L?attribut `Name` sp?cifie l?application h?te Office pouvant ex?cuter votre compl?ment. Les h?tes Office sont pris en charge sur diff?rentes plateformes et sont ex?cut?s sur les ordinateurs de bureau, les navigateurs web, les tablettes et les appareils mobiles. Vous ne pouvez pas indiquer quelle plateforme peut ?tre utilis?e pour ex?cuter votre compl?ment. Par exemple, si vous sp?cifiez `Mailbox`, Outlook et Outlook Web App peuvent ?tre utilis?s pour ex?cuter votre compl?ment.</span><span class="sxs-lookup"><span data-stu-id="705cb-p116">The  `Name` attribute specifies the Office host application that can run your add-in. Office hosts are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices. You can't specify which platform can be used to run your add-in. For example, if you specify `Mailbox`, both Outlook and Outlook Web App can be used to run your add-in.</span></span> 


## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="705cb-181">D?finition de l??l?ment Requirements dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="705cb-181">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="705cb-p117">L??l?ment **Requirements** indique les ensembles de conditions minimales requises ou les membres d?API qui doivent ?tre pris en charge par l?h?te Office en vue d?ex?cuter votre compl?ment. L??l?ment **Requirements** peut indiquer des ensembles de conditions requises et des m?thodes individuelles utilis?s dans votre compl?ment. Dans la version 1.1 du sch?ma de manifeste du compl?ment, l??l?ment **Requirements** est facultatif pour tous les compl?ments, sauf pour les compl?ments Outlook.</span><span class="sxs-lookup"><span data-stu-id="705cb-p117">The  **Requirements** element specifies the minimum requirement sets or API members that must be supported by the Office host to run your add-in. The **Requirements** element can specify both requirement sets and individual methods used in your add-in. In version 1.1 of the add-in manifest schema, the **Requirements** element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="705cb-p118">Utilisez uniquement l??l?ment **Requirements** pour sp?cifier des ensembles de conditions requises essentiels ou des membres d?API que votre compl?ment doit utiliser. Si la plateforme ou l?h?te Office ne prend pas en charge les ensembles de conditions requises ou les membres d?API sp?cifi?s dans l??l?ment **Requirements**, le compl?ment ne s?ex?cute pas dans cet h?te ou cette plateforme et ne s?affiche pas dans **Mes compl?ments**. Nous vous recommandons plut?t de rendre votre compl?ment disponible sur toutes les plateformes d?un h?te Office, comme Excel pour Windows, Excel Online et Excel pour iPad. Pour rendre votre compl?ment disponible sur _tous_ les h?tes et plateformes Office, utilisez des v?rifications ? l?ex?cution ? la place de l??l?ment **Requirements**.</span><span class="sxs-lookup"><span data-stu-id="705cb-p118">Only use the **Requirements** element to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the **Requirements** element, the add-in won't run in that host or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad. To make your add-in available on  _all_ Office hosts and platforms, use runtime checks instead of the **Requirements** element.</span></span>

<span data-ttu-id="705cb-188">Cet exemple de code illustre un compl?ment qui se charge dans toutes les applications h?tes Office qui prennent en charge les ?l?ments suivants :</span><span class="sxs-lookup"><span data-stu-id="705cb-188">The following code example shows an add-in that loads in all Office host applications that support the following:</span></span>

-  <span data-ttu-id="705cb-189">Un ensemble de conditions requises **TableBindings**, dont la version minimale est 1.1.</span><span class="sxs-lookup"><span data-stu-id="705cb-189">**TableBindings** requirement set, which has a minimum version of 1.1.</span></span>
    
-  <span data-ttu-id="705cb-190">Un ensemble de conditions requises **OOXML**, dont la version minimale est 1.1.</span><span class="sxs-lookup"><span data-stu-id="705cb-190">**OOXML** requirement set, which has a minimum version of 1.1.</span></span>
    
-  <span data-ttu-id="705cb-191">La m?thode **Document.getSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="705cb-191">**Document.getSelectedDataAsync** method.</span></span>

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- <span data-ttu-id="705cb-192">L??l?ment **Requirements** contient les ?l?ments enfants **Sets** et **Methods**.</span><span class="sxs-lookup"><span data-stu-id="705cb-192">The  **Requirements** element contains the **Sets** and **Methods** child elements.</span></span>
    
- <span data-ttu-id="705cb-p119">L??l?ment  **Sets** peut contenir un ou plusieurs ?l?ments  **Set**.  **DefaultMinVersion** indique la valeur **MinVersion** par d?faut de tous les ?l?ments  **Set** enfants.</span><span class="sxs-lookup"><span data-stu-id="705cb-p119">The  **Sets** element can contain one or more **Set** elements. **DefaultMinVersion** specifies the default **MinVersion** value of all child **Set** elements.</span></span>
    
- <span data-ttu-id="705cb-p120">L??l?ment **Set** sp?cifie les ensembles de conditions requises que l?h?te Office doit prendre en charge pour ex?cuter le compl?ment. L?attribut **Name** indique le nom de l?ensemble de conditions requises. L?attribut **MinVersion** sp?cifie la version minimale de l?ensemble de conditions requises. L?attribut **MinVersion** remplace la valeur de **DefaultMinVersion**. Pour plus d?informations sur les ensembles de conditions requises et les versions auxquelles les membres de votre API appartiennent, consultez [Ensemble de conditions requises pour les compl?ments Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="705cb-p120">The  **Set** element specifies requirement sets that the Office host must support to run the add-in. The **Name** attribute specifies the name of the requirement set. The **MinVersion** specifies the minimum version of the requirement set. **MinVersion** overrides the value of **DefaultMinVersion**. For more information about requirement sets and requirement set versions that your API members belong to, see [Office add-in requirement sets](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span></span>
    
- <span data-ttu-id="705cb-p121">L??l?ment **Methods** peut contenir un ou plusieurs ?l?ments **Method**. Vous ne pouvez pas utiliser l??l?ment **Methods** avec des compl?ments Outlook.</span><span class="sxs-lookup"><span data-stu-id="705cb-p121">The  **Methods** element can contain one or more **Method** elements. You can't use the **Methods** element with Outlook add-ins.</span></span>
    
- <span data-ttu-id="705cb-p122">L??l?ment  **Method** sp?cifie une m?thode individuelle qui doit ?tre prise en charge dans l?h?te Office o? votre compl?ment est ex?cut?. L?attribut **Name** est obligatoire et indique le nom de la m?thode qualifi?e avec son objet parent.</span><span class="sxs-lookup"><span data-stu-id="705cb-p122">The  **Method** element specifies an individual method that must be supported in the Office host where your add-in runs. The **Name** attribute is required and specifies the name of the method qualified with its parent object.</span></span>
    

## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="705cb-204">Utilisation des v?rifications ? l?ex?cution dans votre code JavaScript</span><span class="sxs-lookup"><span data-stu-id="705cb-204">Use runtime checks in your JavaScript code</span></span>


<span data-ttu-id="705cb-p123">Vous pouvez fournir des fonctionnalit?s suppl?mentaires dans votre compl?ment si certains ensembles de conditions requises sont pris en charge par l?h?te Office. Par exemple, vous pouvez utiliser les nouvelles interfaces API JavaScript de Word dans votre compl?ment existant si ce dernier est ex?cut? dans Word 2016. Pour ce faire, utilisez la m?thode **isSetSupported** avec le nom de l?ensemble de conditions requises. **isSetSupported** d?termine, lors de l?ex?cution, si l?h?te Office ex?cutant le compl?ment prend en charge l?ensemble des conditions requises. Si l?ensemble de conditions requises est pris en charge, **isSetSupported** renvoie **True** et ex?cute le code suppl?mentaire qui utilise les membres d?API provenant de l?ensemble de conditions requises. Si l?h?te Office ne prend pas en charge l?ensemble de conditions requises, **isSetSupported** renvoie **False** et le code suppl?mentaire n?est pas ex?cut?. Le code suivant indique la syntaxe ? utiliser avec **isSetSupported**.</span><span class="sxs-lookup"><span data-stu-id="705cb-p123">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office host. For example, you might want to use the new Word JavaScript APIs Word in your existing add-in if your add-in runs in Word 2016. To do this, you use the  **isSetSupported** method with the name of the requirement set. **isSetSupported** determines, at runtime, whether the Office host running the add-in supports the requirement set. If the requirement set is supported, **isSetSupported** returns **true** and runs the additional code that uses the API members from that requirement set. If the Office host doesn't support the requirement set, **isSetSupported** returns **false** and the additional code won't run. The following code shows the syntax to use with **isSetSupported**.</span></span>


```js
if (Office.context.requirements.isSetSupported(RequirementSetName , VersionNumber))
{
   // Code that uses API members from RequirementSetName.
}

```


-  <span data-ttu-id="705cb-p124">_RequirementSetName_ (obligatoire) est une cha?ne repr?sentant le nom de l?ensemble de conditions requises. Pour plus d?informations sur les ensembles de conditions requises disponibles, voir [Ensemble de conditions requises pour les compl?ments Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="705cb-p124">_RequirementSetName_ (required) is a string that represents the name of the requirement set. For more information about available requirement sets, see [Office add-in requirement sets](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span></span>
    
-  <span data-ttu-id="705cb-214">_VersionNumber_ (facultatif) correspond ? la version de l?ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="705cb-214">_VersionNumber_ (optional) is the version of the requirement set.</span></span>
    
<span data-ttu-id="705cb-p125">Dans Excel 2016 ou Word 2016, utilisez **isSetSupported** avec les ensembles de conditions requises  **ExcelAPI** ou **WordAPI**. La m?thode  **isSetSupported**, ainsi que les ensembles de conditions requises  **ExcelAPI** et **WordAPI**, sont disponibles dans le dernier fichier Office.js du CDN. Si vous n?utilisez pas Office.js ? partir du CDN, votre compl?ment peut g?n?rer des exceptions, car la m?thode  **isSetSupported** ne sera pas d?finie. Pour plus d?informations, voir [ Indication de l?API JavaScript la plus r?cente pour la biblioth?que Office](#specify-the-latest-javascript-api-for-office-library).</span><span class="sxs-lookup"><span data-stu-id="705cb-p125">In Excel 2016 or Word 2016, use  **isSetSupported** with the **ExcelAPI** or **WordAPI** requirement sets. The **isSetSupported** method, and the **ExcelAPI** and **WordAPI** requirement sets, are available in the latest Office.js file available from the CDN. If you don't use Office.js from the CDN, your add-in might generate exceptions because **isSetSupported** will be undefined. For more information, see [Specify the latest JavaScript API for Office library](#specify-the-latest-javascript-api-for-office-library).</span></span> 


> [!NOTE]
> <span data-ttu-id="705cb-p126">**isSetSupported** ne fonctionne pas dans Outlook ou Outlook Web App. Pour utiliser une v?rification ? l?ex?cution dans Outlook ou Outlook Web App, utilisez la technique d?crite dans la section [V?rifications ? l?ex?cution ? l?aide de m?thodes ne faisant pas partie d?un ensemble de conditions requises](#runtime-checks-using-methods-not-in-a-requirement-set).</span><span class="sxs-lookup"><span data-stu-id="705cb-p126">**isSetSupported** does not work in Outlook or Outlook Web App. To use a runtime check in Outlook or Outlook Web App, use the technique described in [Runtime checks using methods not in a requirement set](#runtime-checks-using-methods-not-in-a-requirement-set).</span></span>

<span data-ttu-id="705cb-221">L?exemple de code suivant montre comment un compl?ment peut fournir des fonctionnalit?s diff?rentes pour divers h?tes Office qui peuvent prendre en charge plusieurs ensembles de conditions requises ou membres d?API.</span><span class="sxs-lookup"><span data-stu-id="705cb-221">The following code example shows how an add-in can provide different functionality for different Office hosts that might support different requirement sets or API members.</span></span>




```js
if (Office.context.requirements.isSetSupported('WordApi', 1.1))
{
    // Run code that provides additional functionality using the JavaScript API for Word when the add-in runs in Word 2016.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
      // Run code that uses API members from the CustomXmlParts requirement set.
}
else 
{
    // Run additional code when the Office host is not Word 2016, and when the Office host does not support the CustomXmlParts requirement set.
}

```


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="705cb-222">V?rifications ? l?ex?cution ? l?aide de m?thodes ne faisant pas partie d?un ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="705cb-222">Runtime checks using methods not in a requirement set</span></span>


<span data-ttu-id="705cb-p127">Certains membres API n?appartiennent pas ? des ensembles de conditions requises. Cela s?applique uniquement aux membres d?API qui font partie de l?espace de noms de l?[interface API JavaScript pour Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) (rien sous Office), et non aux membres d?API qui appartiennent ? l?espace de noms de l?interface API JavaScript pour Word (rien dans Word) ou de la [r?f?rence de l?API JavaScript pour les compl?ments Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) (rien dans Excel). Lorsque votre compl?ment d?pend d?une m?thode qui ne fait pas partie d?un ensemble de conditions requises, vous pouvez utiliser la v?rification ? l?ex?cution pour d?terminer si la m?thode est prise en charge par l?h?te Office, comme indiqu? dans l?exemple suivant. Pour consulter la liste compl?te des m?thodes qui n?appartiennent pas ? un ensemble de conditions requises, voir [Ensemble de conditions requises pour les compl?ments Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="705cb-p127">Some API members don't belong to requirement sets. This only applies to API members that are part of the [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) namespace (anything under Office.), not API members that belong to the Word JavaScript API (anything in Word.) or [Excel add-ins JavaScript API reference](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) (anything in Excel.) namespaces. When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office host, as shown in the following code example. For a complete list of methods that don't belong to a requirement set, see [Office add-in requirement sets](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span></span>


> [!NOTE]
> <span data-ttu-id="705cb-227">Nous vous recommandons de limiter l?utilisation de ce type de v?rification ? l?ex?cution dans le code de votre compl?ment.</span><span class="sxs-lookup"><span data-stu-id="705cb-227">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="705cb-228">L?exemple de code suivant v?rifie si l?h?te prend en charge **document.setSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="705cb-228">The following code example checks whether the host supports  **document.setSelectedDataAsync**.</span></span>




```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a><span data-ttu-id="705cb-229">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="705cb-229">See also</span></span>

- [<span data-ttu-id="705cb-230">Manifeste XML des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="705cb-230">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="705cb-231">Ensembles de conditions requises pour les compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="705cb-231">Office add-in requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="705cb-232">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="705cb-232">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)