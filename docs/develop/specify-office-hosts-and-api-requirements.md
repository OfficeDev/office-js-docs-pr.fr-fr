---
title: Spécification des exigences en matière d’hôtes Office et d’API
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd517dee1faf8d3f3009a0b9ce7127f5760e730d
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437709"
---
# <a name="specify-office-hosts-and-api-requirements"></a><span data-ttu-id="bd6d8-102">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="bd6d8-102">Specify Office hosts and API requirements</span></span>

<span data-ttu-id="bd6d8-p101">Il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API pour fonctionner correctement. Par exemple, votre complément peut :</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p101">Your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API in order to work as expected. For example, your add-in might:</span></span>

- <span data-ttu-id="bd6d8-105">exécuter une ou plusieurs application Office (Word ou Excel) ;</span><span class="sxs-lookup"><span data-stu-id="bd6d8-105">Run in a single Office application (Word or Excel), or several applications.</span></span>
    
- <span data-ttu-id="bd6d8-p102">utiliser des API JavaScript disponibles uniquement dans certaines versions d’Office. Par exemple, vous pouvez utiliser les API JavaScript d’Excel dans un complément qui fonctionne dans Excel 2016 ;</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span> 
    
- <span data-ttu-id="bd6d8-108">s’exécuter uniquement dans les versions d’Office qui prennent en charge les membres d’API utilisés par votre complément.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-108">Run only in versions of Office that support API members that your add-in uses.</span></span>
    
<span data-ttu-id="bd6d8-109">Cet article vous aidera à comprendre les options que vous devez choisir afin de vous assurer que votre complément fonctionne comme prévu et atteint l’audience la plus large possible.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-109">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="bd6d8-110">Pour savoir de manière détaillée quelle version d’Office prend en charge les compléments Office, consultez la page relative à la [disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-110">For a high-level view of where Office Add-ins are currently supported, see the [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span> 

<span data-ttu-id="bd6d8-111">Le tableau suivant répertorie les concepts principaux décrits dans cet article.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-111">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="bd6d8-112">**Concept**</span><span class="sxs-lookup"><span data-stu-id="bd6d8-112">**Concept**</span></span>|<span data-ttu-id="bd6d8-113">**Description**</span><span class="sxs-lookup"><span data-stu-id="bd6d8-113">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="bd6d8-114">Application Office, application hôte Office ou hôte Office</span><span class="sxs-lookup"><span data-stu-id="bd6d8-114">Office application, Office host application, Office host, or host</span></span>|<span data-ttu-id="bd6d8-p103">Application Office utilisée pour exécuter votre complément. Par exemple, Word, Word Online ou Excel.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p103">The Office application used to run your add-in. For example, Word, Word Online, Excel, and so on.</span></span>|
|<span data-ttu-id="bd6d8-117">Plateforme</span><span class="sxs-lookup"><span data-stu-id="bd6d8-117">Platform</span></span>|<span data-ttu-id="bd6d8-118">Application sur laquelle l’hôte Office est exécuté, comme Office Online ou Office pour iPad.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-118">Where the Office host runs, such as Office Online or Office for iPad.</span></span>|
|<span data-ttu-id="bd6d8-119">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="bd6d8-119">Requirement set</span></span>|<span data-ttu-id="bd6d8-p104">Groupe nommé de membres d’API associés. Les compléments utilisent des ensembles de conditions requises pour déterminer si l’hôte Office prend en charge les membres d’API utilisés par votre complément. Il est plus facile de tester la prise en charge d’un ensemble de conditions requises, plutôt que la prise en charge de membres individuels d’API. La prise en charge de l’ensemble des conditions requises varie selon l’hôte Office et la version de ce dernier. </span><span class="sxs-lookup"><span data-stu-id="bd6d8-p104">A named group of related API members. Add-ins use requirement sets to determine whether the Office host supports API members used by your add-in. It's easier to test for the support of a requirement set than for the support of individual API members. Requirement set support varies by Office host and the version of the Office host. </span></span><br ><span data-ttu-id="bd6d8-p105">Les ensembles de conditions requises sont spécifiés dans le fichier manifeste. Quand vous définissez des ensembles de conditions requises dans le fichier manifeste, vous définissez le niveau minimal de prise en charge de l’API que l’hôte Office doit fournir pour exécuter votre complément. Les hôtes Office qui ne prennent pas en charge les ensembles de conditions requises spécifiés dans le manifeste ne peuvent pas exécuter votre complément, et votre complément ne sera pas affiché dans <span class="ui">Mes compléments</span>. Cela limite les emplacements où votre complément sera disponible. Dans le code utilisant les vérifications à l’exécution. Pour obtenir la liste complète des ensembles de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p105">Requirement sets are specified in the manifest file. When you specify requirement sets in the manifest, you set the minimum level of API support that the Office host must provide in order to run your add-in. Office hosts that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.In code using runtime checks. For the complete list of requirement sets, see [Office add-in requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>|
|<span data-ttu-id="bd6d8-128">Vérification à l’exécution</span><span class="sxs-lookup"><span data-stu-id="bd6d8-128">Runtime check</span></span>|<span data-ttu-id="bd6d8-p106">Test effectué à l’exécution pour déterminer si l’hôte Office qui exécute votre complément prend en charge les ensembles de conditions requises ou les méthodes utilisés par votre complément. Pour effectuer une vérification à l’exécution, vous pouvez utiliser une instruction **if** avec la méthode **isSetSupported**, les ensembles de conditions requises ou les noms de méthode qui ne font pas partie d’un ensemble de conditions requises. Les vérifications à l’exécution permettent de veiller à ce que votre complément atteigne le plus grand nombre possible de clients. Contrairement aux ensembles de conditions requises, les vérifications à l’exécution ne précisent pas le niveau minimal de prise en charge de l’API que l’hôte Office doit fournir pour l’exécution de votre complément. Au lieu de cela, vous devez utiliser l’instruction **if** afin de déterminer si un membre d’API est pris en charge. Si c’est le cas, vous pouvez fournir des fonctionnalités supplémentaires dans votre complément. Votre complément s’affiche toujours dans **Mes compléments** quand vous effectuez des vérifications à l’exécution.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p106">A test that is performed at runtime to determine whether the Office host running your add-in supports requirement sets or methods used by your add-in. To perform a runtime check, you use an  **if** statement with the **isSetSupported** method, the requirement sets, or the method names that aren't part of a requirement set.Use runtime checks to ensure that your add-in reaches the broadest number of customers. Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office host must provide for your add-in to run. Instead, you use the  **if** statement to determine whether an API member is supported. If it is, you can provide additional functionality in your add-in. Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="bd6d8-135">Avant de commencer</span><span class="sxs-lookup"><span data-stu-id="bd6d8-135">Before you begin</span></span>

<span data-ttu-id="bd6d8-p107">Votre complément doit utiliser la version la plus récente du schéma de manifeste de complément. Si vous utilisez les vérifications à l’exécution dans votre complément, assurez-vous que vous utilisez la dernière API JavaScript pour la bibliothèque Office (office.js).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p107">Your add-in must use the most current version of the add-in manifest schema. If you use runtime checks in your add-in, ensure that you use the latest JavaScript API for Office (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="bd6d8-138">Indication du schéma de manifeste de complément le plus récent</span><span class="sxs-lookup"><span data-stu-id="bd6d8-138">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="bd6d8-p108">Le manifeste de votre du complément doit utiliser la version 1.1 du schéma de manifeste de complément. Définissez l’élément **App_office** dans votre manifeste complément comme suit.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p108">Your add-in's manifest must use version 1.1 of the add-in manifest schema. Set the  **OfficeApp** element in your add-in manifest as follows.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-javascript-api-for-office-library"></a><span data-ttu-id="bd6d8-141">Indication de l’API JavaScript la plus récente pour la bibliothèque Office</span><span class="sxs-lookup"><span data-stu-id="bd6d8-141">Specify the latest JavaScript API for Office library</span></span>

<span data-ttu-id="bd6d8-p109">Si vous utilisez des vérifications à l’exécution, référencez la version la plus récente de l’API JavaScript pour la bibliothèque Office à partir du réseau de livraison de contenu (CDN). Pour ce faire, ajoutez la balise `script` suivante à votre code HTML. L’utilisation de `/1/` dans l’URL CDN garantit que vous référencez la version d’Office.js la plus récente.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p109">If you use runtime checks, reference the most current version of the JavaScript API for Office library from the content delivery network (CDN). To do this, add the following  `script` tag to your HTML. Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a><span data-ttu-id="bd6d8-145">Options pour spécifier des hôtes Office ou les conditions requises d’API</span><span class="sxs-lookup"><span data-stu-id="bd6d8-145">Options to specify Office hosts or API requirements</span></span>

<span data-ttu-id="bd6d8-p110">Lors de la spécification des hôtes Office ou des conditions requises d’API, vous devez tenir compte de plusieurs facteurs. Le diagramme suivant montre comment choisir la technique à utiliser dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p110">When you specify Office hosts or API requirements, there are several factors to consider. The following diagram shows how to decide which technique to use in your add-in.</span></span>

![Optez pour la meilleure solution pour votre complément lorsque vous spécifiez des hôtes Office ou des exigences d’API](../images/options-for-office-hosts.png)

- <span data-ttu-id="bd6d8-p111">Si votre complément s’exécute dans un hôte Office, définissez l’élément **Hosts** dans le manifeste. Pour plus d’informations, consultez [Définition de l’élément Hosts](#set-the-hosts-element).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p111">If your add-in runs in one Office host, set the **Hosts** element in the manifest. For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>
    
- <span data-ttu-id="bd6d8-p112">Pour définir l’ensemble minimal de conditions requises ou les membres minimaux d’API qu’un hôte Office doit prendre en charge pour exécuter votre complément, définissez l’élément **Requirements** dans le manifeste. Pour plus d’informations, consultez la section [ Définition de l’élément Requirements dans le manifeste](#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p112">To set the minimum requirement set or API members that an Office host must support to run your add-in, set the  **Requirements** element in the manifest. For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>
    
- <span data-ttu-id="bd6d8-p113">Si vous souhaitez proposer des fonctionnalités supplémentaires lorsque des ensembles de conditions requises spécifiques ou des membres d’API sont disponibles dans l’hôte Office, effectuez une vérification à l’exécution dans le code JavaScript de votre complément. Par exemple, si votre complément est exécuté dans Excel 2016, utilisez les membres d’API de la nouvelle API JavaScript pour Excel pour fournir des fonctionnalités supplémentaires. Pour plus d’informations, consultez la section [Utilisation des vérifications à l’exécution dans votre code JavaScript](#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p113">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office host, perform a runtime check in your add-in's JavaScript code. For example, if your add-in runs in Excel 2016, use API members from the new JavaScript API for Excel to provide additional functionality. For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>
    
## <a name="set-the-hosts-element"></a><span data-ttu-id="bd6d8-156">Définition de l’élément Hosts</span><span class="sxs-lookup"><span data-stu-id="bd6d8-156">Set the Hosts element</span></span>

<span data-ttu-id="bd6d8-p114">Pour exécuter votre complément dans une application hôte Office, utilisez les éléments **Hosts** et **Host** dans le manifeste. Si vous ne définissez pas l’élément **Hosts**, votre complément sera exécuté dans tous les hôtes.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p114">To make your add-in run in one Office host application, use the  **Hosts** and **Host** elements in the manifest. If you don't specify the **Hosts** element, your add-in will run in all hosts.</span></span>

<span data-ttu-id="bd6d8-159">Par exemple, les déclarations  **Hosts** et **Host** suivantes indiquent que le complément fonctionnera avec n’importe quelle version d’Excel, y compris Excel pour Windows, Excel Online et Excel pour iPad.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-159">For example, the following  **Hosts** and **Host** declaration specifies that the add-in will work with any release of Excel, which includes Excel for Windows, Excel Online, and Excel for iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="bd6d8-p115">L’élément  **Hosts** peut contenir un ou plusieurs éléments  **Host**. L’élément  **Host** indique l’hôte Office requis par votre complément. L’attribut **Name** est requis et peut être défini sur l’une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p115">The  **Hosts** element can contain one or more **Host** elements. The **Host** element specifies the Office host your add-in requires. The **Name** attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="bd6d8-163">Nom</span><span class="sxs-lookup"><span data-stu-id="bd6d8-163">Name</span></span>          | <span data-ttu-id="bd6d8-164">Applications hôtes Office</span><span class="sxs-lookup"><span data-stu-id="bd6d8-164">Office host applications</span></span>                      |
|:--------------|:----------------------------------------------|
| <span data-ttu-id="bd6d8-165">Base de données</span><span class="sxs-lookup"><span data-stu-id="bd6d8-165">Database</span></span>      | <span data-ttu-id="bd6d8-166">applications web Access</span><span class="sxs-lookup"><span data-stu-id="bd6d8-166">Access web apps</span></span>                               |
| <span data-ttu-id="bd6d8-167">Document</span><span class="sxs-lookup"><span data-stu-id="bd6d8-167">Document</span></span>      | <span data-ttu-id="bd6d8-168">Word pour Windows, Mac, iPad et Online</span><span class="sxs-lookup"><span data-stu-id="bd6d8-168">Word for Windows, Mac, iPad and Online</span></span>        |
| <span data-ttu-id="bd6d8-169">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bd6d8-169">Mailbox</span></span>       | <span data-ttu-id="bd6d8-170">Outlook pour Windows, Mac, Web et Outlook.com</span><span class="sxs-lookup"><span data-stu-id="bd6d8-170">Outlook for Windows, Mac, Web and Outlook.com</span></span> | 
| <span data-ttu-id="bd6d8-171">Présentation</span><span class="sxs-lookup"><span data-stu-id="bd6d8-171">Presentation</span></span>  | <span data-ttu-id="bd6d8-172">PowerPoint pour Windows, Mac, iPad et Online</span><span class="sxs-lookup"><span data-stu-id="bd6d8-172">PowerPoint for Windows, Mac, iPad and Online</span></span>  |
| <span data-ttu-id="bd6d8-173">Projet</span><span class="sxs-lookup"><span data-stu-id="bd6d8-173">Project</span></span>       | <span data-ttu-id="bd6d8-174">Projet</span><span class="sxs-lookup"><span data-stu-id="bd6d8-174">Project</span></span>                                       |
| <span data-ttu-id="bd6d8-175">Classeur</span><span class="sxs-lookup"><span data-stu-id="bd6d8-175">Workbook</span></span>      | <span data-ttu-id="bd6d8-176">Excel pour Windows, Mac, iPad et Online</span><span class="sxs-lookup"><span data-stu-id="bd6d8-176">Excel Windows, Mac, iPad and Online</span></span>           |

> [!NOTE]
> <span data-ttu-id="bd6d8-p116">L’attribut `Name` spécifie l’application hôte Office pouvant exécuter votre complément. Les hôtes Office sont pris en charge sur différentes plateformes et sont exécutés sur les ordinateurs de bureau, les navigateurs web, les tablettes et les appareils mobiles. Vous ne pouvez pas indiquer quelle plateforme peut être utilisée pour exécuter votre complément. Par exemple, si vous spécifiez `Mailbox`, Outlook et Outlook Web App peuvent être utilisés pour exécuter votre complément.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p116">The  `Name` attribute specifies the Office host application that can run your add-in. Office hosts are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices. You can't specify which platform can be used to run your add-in. For example, if you specify `Mailbox`, both Outlook and Outlook Web App can be used to run your add-in.</span></span> 


## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="bd6d8-181">Définition de l’élément Requirements dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="bd6d8-181">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="bd6d8-p117">L’élément **Requirements** indique les ensembles de conditions minimales requises ou les membres d’API qui doivent être pris en charge par l’hôte Office en vue d’exécuter votre complément. L’élément **Requirements** peut indiquer des ensembles de conditions requises et des méthodes individuelles utilisés dans votre complément. Dans la version 1.1 du schéma de manifeste du complément, l’élément **Requirements** est facultatif pour tous les compléments, sauf pour les compléments Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p117">The  **Requirements** element specifies the minimum requirement sets or API members that must be supported by the Office host to run your add-in. The **Requirements** element can specify both requirement sets and individual methods used in your add-in. In version 1.1 of the add-in manifest schema, the **Requirements** element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="bd6d8-p118">Utilisez uniquement l’élément **Requirements** pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément **Requirements**, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans **Mes compléments**. Nous vous recommandons plutôt de rendre votre complément disponible sur toutes les plateformes d’un hôte Office, comme Excel pour Windows, Excel Online et Excel pour iPad. Pour rendre votre complément disponible sur _tous_ les hôtes et plateformes Office, utilisez des vérifications à l’exécution à la place de l’élément **Requirements**.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p118">Only use the **Requirements** element to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the **Requirements** element, the add-in won't run in that host or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad. To make your add-in available on  _all_ Office hosts and platforms, use runtime checks instead of the **Requirements** element.</span></span>

<span data-ttu-id="bd6d8-188">Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="bd6d8-188">The following code example shows an add-in that loads in all Office host applications that support the following:</span></span>

-  <span data-ttu-id="bd6d8-189">Un ensemble de conditions requises **TableBindings**, dont la version minimale est 1.1.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-189">**TableBindings** requirement set, which has a minimum version of 1.1.</span></span>
    
-  <span data-ttu-id="bd6d8-190">Un ensemble de conditions requises **OOXML**, dont la version minimale est 1.1.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-190">**OOXML** requirement set, which has a minimum version of 1.1.</span></span>
    
-  <span data-ttu-id="bd6d8-191">La méthode **Document.getSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-191">**Document.getSelectedDataAsync** method.</span></span>

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

- <span data-ttu-id="bd6d8-192">L’élément **Requirements** contient les éléments enfants **Sets** et **Methods**.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-192">The  **Requirements** element contains the **Sets** and **Methods** child elements.</span></span>
    
- <span data-ttu-id="bd6d8-p119">L’élément  **Sets** peut contenir un ou plusieurs éléments  **Set**.  **DefaultMinVersion** indique la valeur **MinVersion** par défaut de tous les éléments  **Set** enfants.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p119">The  **Sets** element can contain one or more **Set** elements. **DefaultMinVersion** specifies the default **MinVersion** value of all child **Set** elements.</span></span>
    
- <span data-ttu-id="bd6d8-p120">L’élément **Set** spécifie les ensembles de conditions requises que l’hôte Office doit prendre en charge pour exécuter le complément. L’attribut **Name** indique le nom de l’ensemble de conditions requises. L’attribut **MinVersion** spécifie la version minimale de l’ensemble de conditions requises. L’attribut **MinVersion** remplace la valeur de **DefaultMinVersion**. Pour plus d’informations sur les ensembles de conditions requises et les versions auxquelles les membres de votre API appartiennent, consultez [Ensemble de conditions requises pour les compléments Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p120">The  **Set** element specifies requirement sets that the Office host must support to run the add-in. The **Name** attribute specifies the name of the requirement set. The **MinVersion** specifies the minimum version of the requirement set. **MinVersion** overrides the value of **DefaultMinVersion**. For more information about requirement sets and requirement set versions that your API members belong to, see [Office add-in requirement sets](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span></span>
    
- <span data-ttu-id="bd6d8-p121">L’élément **Methods** peut contenir un ou plusieurs éléments **Method**. Vous ne pouvez pas utiliser l’élément **Methods** avec des compléments Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p121">The  **Methods** element can contain one or more **Method** elements. You can't use the **Methods** element with Outlook add-ins.</span></span>
    
- <span data-ttu-id="bd6d8-p122">L’élément  **Method** spécifie une méthode individuelle qui doit être prise en charge dans l’hôte Office où votre complément est exécuté. L’attribut **Name** est obligatoire et indique le nom de la méthode qualifiée avec son objet parent.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p122">The  **Method** element specifies an individual method that must be supported in the Office host where your add-in runs. The **Name** attribute is required and specifies the name of the method qualified with its parent object.</span></span>
    

## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="bd6d8-204">Utilisation des vérifications à l’exécution dans votre code JavaScript</span><span class="sxs-lookup"><span data-stu-id="bd6d8-204">Use runtime checks in your JavaScript code</span></span>


<span data-ttu-id="bd6d8-p123">Vous pouvez fournir des fonctionnalités supplémentaires dans votre complément si certains ensembles de conditions requises sont pris en charge par l’hôte Office. Par exemple, vous pouvez utiliser les nouvelles interfaces API JavaScript de Word dans votre complément existant si ce dernier est exécuté dans Word 2016. Pour ce faire, utilisez la méthode **isSetSupported** avec le nom de l’ensemble de conditions requises. **isSetSupported** détermine, lors de l’exécution, si l’hôte Office exécutant le complément prend en charge l’ensemble des conditions requises. Si l’ensemble de conditions requises est pris en charge, **isSetSupported** renvoie **True** et exécute le code supplémentaire qui utilise les membres d’API provenant de l’ensemble de conditions requises. Si l’hôte Office ne prend pas en charge l’ensemble de conditions requises, **isSetSupported** renvoie **False** et le code supplémentaire n’est pas exécuté. Le code suivant indique la syntaxe à utiliser avec **isSetSupported**.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p123">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office host. For example, you might want to use the new Word JavaScript APIs Word in your existing add-in if your add-in runs in Word 2016. To do this, you use the  **isSetSupported** method with the name of the requirement set. **isSetSupported** determines, at runtime, whether the Office host running the add-in supports the requirement set. If the requirement set is supported, **isSetSupported** returns **true** and runs the additional code that uses the API members from that requirement set. If the Office host doesn't support the requirement set, **isSetSupported** returns **false** and the additional code won't run. The following code shows the syntax to use with **isSetSupported**.</span></span>


```js
if (Office.context.requirements.isSetSupported(RequirementSetName , VersionNumber))
{
   // Code that uses API members from RequirementSetName.
}

```


-  <span data-ttu-id="bd6d8-p124">_RequirementSetName_ (obligatoire) est une chaîne représentant le nom de l’ensemble de conditions requises. Pour plus d’informations sur les ensembles de conditions requises disponibles, voir [Ensemble de conditions requises pour les compléments Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p124">_RequirementSetName_ (required) is a string that represents the name of the requirement set. For more information about available requirement sets, see [Office add-in requirement sets](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span></span>
    
-  <span data-ttu-id="bd6d8-214">_VersionNumber_ (facultatif) correspond à la version de l’ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-214">_VersionNumber_ (optional) is the version of the requirement set.</span></span>
    
<span data-ttu-id="bd6d8-p125">Dans Excel 2016 ou Word 2016, utilisez **isSetSupported** avec les ensembles de conditions requises  **ExcelAPI** ou **WordAPI**. La méthode  **isSetSupported**, ainsi que les ensembles de conditions requises  **ExcelAPI** et **WordAPI**, sont disponibles dans le dernier fichier Office.js du CDN. Si vous n’utilisez pas Office.js à partir du CDN, votre complément peut générer des exceptions, car la méthode  **isSetSupported** ne sera pas définie. Pour plus d’informations, voir [ Indication de l’API JavaScript la plus récente pour la bibliothèque Office](#specify-the-latest-javascript-api-for-office-library).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p125">In Excel 2016 or Word 2016, use  **isSetSupported** with the **ExcelAPI** or **WordAPI** requirement sets. The **isSetSupported** method, and the **ExcelAPI** and **WordAPI** requirement sets, are available in the latest Office.js file available from the CDN. If you don't use Office.js from the CDN, your add-in might generate exceptions because **isSetSupported** will be undefined. For more information, see [Specify the latest JavaScript API for Office library](#specify-the-latest-javascript-api-for-office-library).</span></span> 


> [!NOTE]
> <span data-ttu-id="bd6d8-p126">**isSetSupported** ne fonctionne pas dans Outlook ou Outlook Web App. Pour utiliser une vérification à l’exécution dans Outlook ou Outlook Web App, utilisez la technique décrite dans la section [Vérifications à l’exécution à l’aide de méthodes ne faisant pas partie d’un ensemble de conditions requises](#runtime-checks-using-methods-not-in-a-requirement-set).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p126">**isSetSupported** does not work in Outlook or Outlook Web App. To use a runtime check in Outlook or Outlook Web App, use the technique described in [Runtime checks using methods not in a requirement set](#runtime-checks-using-methods-not-in-a-requirement-set).</span></span>

<span data-ttu-id="bd6d8-221">L’exemple de code suivant montre comment un complément peut fournir des fonctionnalités différentes pour divers hôtes Office qui peuvent prendre en charge plusieurs ensembles de conditions requises ou membres d’API.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-221">The following code example shows how an add-in can provide different functionality for different Office hosts that might support different requirement sets or API members.</span></span>




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


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="bd6d8-222">Vérifications à l’exécution à l’aide de méthodes ne faisant pas partie d’un ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="bd6d8-222">Runtime checks using methods not in a requirement set</span></span>


<span data-ttu-id="bd6d8-p127">Certains membres API n’appartiennent pas à des ensembles de conditions requises. Cela s’applique uniquement aux membres d’API qui font partie de l’espace de noms de l’[interface API JavaScript pour Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) (rien sous Office), et non aux membres d’API qui appartiennent à l’espace de noms de l’interface API JavaScript pour Word (rien dans Word) ou de la [référence de l’API JavaScript pour les compléments Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) (rien dans Excel). Lorsque votre complément dépend d’une méthode qui ne fait pas partie d’un ensemble de conditions requises, vous pouvez utiliser la vérification à l’exécution pour déterminer si la méthode est prise en charge par l’hôte Office, comme indiqué dans l’exemple suivant. Pour consulter la liste complète des méthodes qui n’appartiennent pas à un ensemble de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="bd6d8-p127">Some API members don't belong to requirement sets. This only applies to API members that are part of the [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) namespace (anything under Office.), not API members that belong to the Word JavaScript API (anything in Word.) or [Excel add-ins JavaScript API reference](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) (anything in Excel.) namespaces. When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office host, as shown in the following code example. For a complete list of methods that don't belong to a requirement set, see [Office add-in requirement sets](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).</span></span>


> [!NOTE]
> <span data-ttu-id="bd6d8-227">Nous vous recommandons de limiter l’utilisation de ce type de vérification à l’exécution dans le code de votre complément.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-227">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="bd6d8-228">L’exemple de code suivant vérifie si l’hôte prend en charge **document.setSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="bd6d8-228">The following code example checks whether the host supports  **document.setSelectedDataAsync**.</span></span>




```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a><span data-ttu-id="bd6d8-229">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bd6d8-229">See also</span></span>

- [<span data-ttu-id="bd6d8-230">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="bd6d8-230">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="bd6d8-231">Ensembles de conditions requises pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="bd6d8-231">Office add-in requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="bd6d8-232">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="bd6d8-232">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)