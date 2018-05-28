---
title: Cr?ation de commandes de compl?ment dans votre manifeste pour Excel, Word et PowerPoint
description: Utilisez VersionOverrides dans votre manifeste pour d?finir des commandes de compl?ment pour Excel, Word et PowerPoint. Utilisez les commandes de compl?ment pour cr?er des ?l?ments d?interface utilisateur, ajouter des boutons ou des listes, et effectuer des actions.
ms.date: 12/04/2017
ms.openlocfilehash: 95861fe0de6f0f56f6436b98cd7ad8dee510e82d
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-word-and-powerpoint"></a><span data-ttu-id="18220-104">Cr?ation de commandes de compl?ment dans votre manifeste pour Excel, Word et PowerPoint</span><span class="sxs-lookup"><span data-stu-id="18220-104">Create add-in commands in your manifest for Excel, Word, and PowerPoint</span></span>


<span data-ttu-id="18220-105">Utilisez **[VersionOverrides](https://dev.office.com/reference/add-ins/manifest/versionoverrides)** dans votre manifeste pour d?finir des commandes de compl?ment pour Excel, Word et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="18220-105">Use **[VersionOverrides](https://dev.office.com/reference/add-ins/manifest/versionoverrides)** in your manifest to define add-in commands for Excel, Word, and PowerPoint.</span></span> <span data-ttu-id="18220-106">Les commandes de compl?ment sont un moyen de personnaliser facilement l?interface utilisateur Office par d?faut en y ajoutant des ?l?ments d?interface de votre choix qui ex?cutent des actions.</span><span class="sxs-lookup"><span data-stu-id="18220-106">Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions.</span></span> <span data-ttu-id="18220-107">Vous pouvez utiliser les commandes de compl?ment pour :</span><span class="sxs-lookup"><span data-stu-id="18220-107">You can use add-in commands to:</span></span>
- <span data-ttu-id="18220-108">cr?er des ?l?ments d?interface utilisateur ou des points d?entr?e qui facilitent l?utilisation des fonctionnalit?s de votre compl?ment ;</span><span class="sxs-lookup"><span data-stu-id="18220-108">Create UI elements or entry points that make your add-in's functionality easier to use.</span></span>  
  
- <span data-ttu-id="18220-109">ajouter des boutons ou une liste d?roulante de boutons sur le ruban ;</span><span class="sxs-lookup"><span data-stu-id="18220-109">Add buttons or a drop-down list of buttons to the ribbon.</span></span>    
  
- <span data-ttu-id="18220-110">ajouter des options de menu individuelles (pouvant chacune contenir des sous-menus) ? des menus contextuels sp?cifiques ;</span><span class="sxs-lookup"><span data-stu-id="18220-110">Add individual menu items ? each containing optional submenus ? to specific context (shortcut) menus.</span></span>    
  
- <span data-ttu-id="18220-p103">ex?cuter des actions lorsque vous avez choisi une commande de compl?ment. Vous pouvez effectuer les op?rations suivantes :</span><span class="sxs-lookup"><span data-stu-id="18220-p103">Perform actions when your add-in command is chosen. You can:</span></span>
    
  - <span data-ttu-id="18220-p104">afficher des compl?ments de volet de t?ches avec lesquels les utilisateurs peuvent interagir. Dans votre compl?ment de volet de t?ches, vous pouvez afficher le code HTML qui utilise la structure de l?interface utilisateur Office pour cr?er une interface utilisateur personnalis?e ;</span><span class="sxs-lookup"><span data-stu-id="18220-p104">Show one or more task pane add-ins for users to interact with. Inside your task pane add-in, you can display HTML that uses Office UI Fabric to create a custom UI.</span></span>
    
     <span data-ttu-id="18220-115">*ou*</span><span class="sxs-lookup"><span data-stu-id="18220-115">*or*</span></span> 
      
  - <span data-ttu-id="18220-116">ex?cuter du code JavaScript, ce qui se fait normalement sans afficher d?interface utilisateur ;</span><span class="sxs-lookup"><span data-stu-id="18220-116">Run JavaScript code, which normally runs without displaying any UI.</span></span>
      
<span data-ttu-id="18220-p105">Cet article explique comment modifier un manifeste pour d?finir des commandes de compl?ment. Le sch?ma suivant illustre la hi?rarchie des ?l?ments utilis?s pour d?finir des commandes de compl?ment. Ces ?l?ments sont d?crits plus en d?tail dans cet article.</span><span class="sxs-lookup"><span data-stu-id="18220-p105">This article describes how to edit your manifest to define add-in commands. The following diagram shows the hierarchy of elements used to define add-in commands. These elements are described in more detail in this article.</span></span> 
      
<span data-ttu-id="18220-120">L?image ci-apr?s est une pr?sentation des ?l?ments de commandes de compl?ment dans le fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="18220-120">The following image is an overview of add-in commands elements in the manifest.</span></span> 
<span data-ttu-id="18220-121">![Pr?sentation des ?l?ments de commandes de compl?ment dans le manifeste](../images/version-overrides.png)</span><span class="sxs-lookup"><span data-stu-id="18220-121">![Overview of add-in commands elements in the manifest](../images/version-overrides.png)</span></span>
 
## <a name="step-1-start-from-a-sample"></a><span data-ttu-id="18220-122">?tape 1 : d?marrer ? partir d?un exemple</span><span class="sxs-lookup"><span data-stu-id="18220-122">Step 1: Start from a sample</span></span>

<span data-ttu-id="18220-p107">Nous vous recommandons vivement de commencer ? partir d?un des exemples que nous fournissons sur la page des [exemples de commandes de compl?ment Office](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Si vous le souhaitez, vous pouvez cr?er votre propre manifeste en suivant les ?tapes d?crites dans ce guide. Vous pouvez valider votre manifeste ? l?aide du fichier XSD sur le site des exemples de commandes de compl?ment Office. Assurez-vous que vous avez lu la rubrique [Commandes de compl?ment pour Excel, Word et PowerPoint](../design/add-in-commands.md) avant d?utiliser les commandes de compl?ment.</span><span class="sxs-lookup"><span data-stu-id="18220-p107">We strongly recommend that you start from one of the samples we provide in  [Office Add-in Commands Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Optionally, you can create your own manifest by following the steps in this guide. You can validate your manifest using the XSD file in the Office Add-in Commands Samples site. Ensure that you have read  [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands.md) before using add-in commands.</span></span>

## <a name="step-2-create-a-task-pane-add-in"></a><span data-ttu-id="18220-127">?tape 2 : cr?er un compl?ment de volet Office</span><span class="sxs-lookup"><span data-stu-id="18220-127">Step 2: Create a task pane add-in</span></span>

<span data-ttu-id="18220-p108">Pour utiliser les commandes de compl?ment, vous devez tout d?abord cr?er un compl?ment de volet Office, puis modifier le manifeste du compl?ment, comme d?crit dans cet article. Vous ne pouvez pas utiliser de commandes de compl?ment avec les compl?ments de contenu. Si vous mettez ? jour un manifeste existant, vous devez ajouter les **espaces de noms XML** appropri?s, ainsi que l??l?ment **VersionOverrides** au manifeste, comme d?crit ? l?[?tape 3 : Ajoutez l??l?ment VersionOverrides](#step-3-add-versionoverrides-element).</span><span class="sxs-lookup"><span data-stu-id="18220-p108">To start using add-in commands, you must first create a task pane add-in, and then modify the add-in's manifest as described in this article. You can't use add-in commands with content add-ins. If you're updating an existing manifest, you must add the appropiate **XML namespaces** as well as add the **VersionOverrides** element to the manifest as described in [Step 3: Add VersionOverrides element](#step-3-add-versionoverrides-element).</span></span>
   
<span data-ttu-id="18220-p109">L?exemple suivant illustre le manifeste d?un compl?ment Office 2013. Ce manifeste ne contient pas de commande de compl?ment car il n?y a pas d??l?ment **VersionOverrides**. Office 2013 ne prend pas en charge les commandes de compl?ment mais, en ajoutant **VersionOverrides** ? ce manifeste, votre compl?ment s?ex?cute dans Office 2013 et Office 2016. Dans Office 2013, votre compl?ment n?affiche pas les commandes de compl?ment et utilise la valeur **SourceLocation** pour ex?cuter votre compl?ment sous la forme d?un compl?ment de volet de t?ches unique. Dans Office 2016, si aucun ?l?ment **VersionOverrides** n?est inclus, **SourceLocation** est utilis? pour ex?cuter votre compl?ment. Cependant, si vous incluez **VersionOverrides**, votre compl?ment affiche uniquement les commandes de compl?ment et n?affiche pas votre compl?ment sous la forme d?un compl?ment de volet de t?ches unique.</span><span class="sxs-lookup"><span data-stu-id="18220-p109">The following example shows an Office 2013 add-in's manifest. There are no add-in commands in this manifest because there is no **VersionOverrides** element. Office 2013 doesn't support add-in commands, but by adding **VersionOverrides** to this manifest, your add-in will run in both Office 2013 and Office 2016. In Office 2013, your add-in won't display add-in commands, and uses the value of **SourceLocation** to run your add-in as a single task pane add-in. In Office 2016, if no **VersionOverrides** element is included, **SourceLocation** is used to run your add-in. If you include **VersionOverrides**, however, your add-in displays the add-in commands only, and doesn't display your add-in as a single task pane add-in.</span></span>
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="~remoteAppUrl/Images/Icon_32.png" />
 
  <AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
    <AppDomain>AppDomain3</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/Pages/Home.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

 <!-- The VersionOverrides element is inserted at this location in the manifest. -->

</OfficeApp>
```

## <a name="step-3-add-versionoverrides-element"></a><span data-ttu-id="18220-136">?tape 3 : ajouter un ?l?ment VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="18220-136">Step 3: Add VersionOverrides element</span></span>
<span data-ttu-id="18220-p110">L??l?ment **VersionOverrides** est l??l?ment racine qui contient la d?finition de votre commande de compl?ment. **VersionOverrides** est un ?l?ment enfant de l??l?ment **OfficeApp** dans le manifeste. Le tableau suivant r?pertorie les attributs de l??l?ment **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="18220-p110">The **VersionOverrides** element is the root element that contains the definition of your add-in command. **VersionOverrides** is a child element of the **OfficeApp** element in the manifest. The following table lists the attributes of the **VersionOverrides** element.</span></span>

|<span data-ttu-id="18220-140">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="18220-140">**Attribute**</span></span>|<span data-ttu-id="18220-141">**Description**</span><span class="sxs-lookup"><span data-stu-id="18220-141">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="18220-142">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="18220-142">**xmlns**</span></span> <br/> | <span data-ttu-id="18220-p111">Obligatoire. Emplacement du sch?ma, qui doit ?tre ? http://schemas.microsoft.com/office/taskpaneappversionoverrides ?.</span><span class="sxs-lookup"><span data-stu-id="18220-p111">Required. The schema location, which must be "http://schemas.microsoft.com/office/taskpaneappversionoverrides".</span></span> <br/> |
|<span data-ttu-id="18220-145">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="18220-145">**xsi:type**</span></span> <br/> |<span data-ttu-id="18220-p112">Obligatoire. Version du sch?ma. La version d?crite dans cet article est ? VersionOverridesV1_0 ?.</span><span class="sxs-lookup"><span data-stu-id="18220-p112">Required. The schema version. The version described in this article is "VersionOverridesV1_0".</span></span>  <br/> |
   
<span data-ttu-id="18220-149">Le tableau suivant pr?sente les ?l?ments enfants de **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="18220-149">The following table identifies the child elements of **VersionOverrides**.</span></span>
  
|<span data-ttu-id="18220-150">**?l?ment**</span><span class="sxs-lookup"><span data-stu-id="18220-150">**Element**</span></span>|<span data-ttu-id="18220-151">**Description**</span><span class="sxs-lookup"><span data-stu-id="18220-151">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="18220-152">**Description**</span><span class="sxs-lookup"><span data-stu-id="18220-152">**Description**</span></span> <br/> |<span data-ttu-id="18220-p113">Facultatif. D?crit le compl?ment. Cet ?l?ment **Description** enfant remplace un ?l?ment **Description** pr?c?dent dans la partie parent du manifeste. L?attribut **resid** pour cet ?l?ment **Description** est d?fini sur l?**id** d?un ?l?ment **Cha?ne**. L??l?ment **Cha?ne** contient le texte pour la **description**. </span><span class="sxs-lookup"><span data-stu-id="18220-p113">Optional. Describes the add-in. This child **Description** element overrides a previous **Description** element in the parent portion of the manifest. The **resid** attribute for this **Description** element is set to the **id** of a **String** element. The **String** element contains the text for **Description**. </span></span><br/> |
|<span data-ttu-id="18220-158">**Configuration requise**</span><span class="sxs-lookup"><span data-stu-id="18220-158">**Requirements**</span></span> <br/> |<span data-ttu-id="18220-p114">Facultatif. Sp?cifie l?ensemble de conditions requises minimal et la version d?Office.js qui doit ?tre activ?e par le compl?ment Office. Cet ?l?ment **Configuration requise** enfant remplace l??l?ment **Configuration requise** dans la partie parent du manifeste. Pour plus d?informations, consultez la rubrique [Sp?cifier les h?tes Office et la configuration requise d?API](../develop/specify-office-hosts-and-api-requirements.md).  </span><span class="sxs-lookup"><span data-stu-id="18220-p114">Optional. Specifies the minimum requirement set and version of Office.js that the add-in requires. This child **Requirements** element overrides the **Requirements** element in the parent portion of the manifest. For more information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md).  </span></span><br/> |
|<span data-ttu-id="18220-163">**H?tes**</span><span class="sxs-lookup"><span data-stu-id="18220-163">**Hosts**</span></span> <br/> |<span data-ttu-id="18220-p115">Obligatoire. Sp?cifie une collection d?h?tes d?Office. L??l?ment **H?tes** enfant remplace l??l?ment **H?tes** dans la partie parent du manifeste. Vous devez inclure un ensemble d?attributs **xsi:type** ? ? Classeur ? ou ? Document ?. </span><span class="sxs-lookup"><span data-stu-id="18220-p115">Required. Specifies a collection of Office hosts. The child **Hosts** element overrides the **Hosts** element in the parent portion of the manifest. You must include a **xsi:type** attribute set to "Workbook" or "Document". </span></span><br/> |
|<span data-ttu-id="18220-168">**Ressources**</span><span class="sxs-lookup"><span data-stu-id="18220-168">**Resources**</span></span> <br/> |<span data-ttu-id="18220-p116">D?finit une collection de ressources (cha?nes, URL et images) qui sont r?f?renc?es par d?autres ?l?ments de manifeste. Par exemple, la valeur de l??l?ment **Description** fait r?f?rence ? un ?l?ment enfant dans **Ressources**. L??l?ment **Ressources** est d?crit ? l?[?tape 7 : ajouter l??l?ment Ressources](#step-7-add-the-resources-element), plus loin dans cet article. </span><span class="sxs-lookup"><span data-stu-id="18220-p116">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference. For example, the **Description** element's value refers to a child element in **Resources**. The **Resources** element is described in [Step 7: Add the Resources element](#step-7-add-the-resources-element) later in this article. </span></span><br/> |
   
<span data-ttu-id="18220-172">L?exemple suivant montre comment utiliser l??l?ment **VersionOverrides** et ses ?l?ments enfants.</span><span class="sxs-lookup"><span data-stu-id="18220-172">The following example shows how to use the **VersionOverrides** element and its child elements.</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information about requirement sets -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <!-- add information about form factors -->
      </Host>
      <Host xsi:type="Document">
        <!-- add information about form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information about resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a><span data-ttu-id="18220-173">?tape 4 : ajouter des ?l?ments Hosts, Host et DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="18220-173">Step 4: Add Hosts, Host, and DesktopFormFactor elements</span></span>

<span data-ttu-id="18220-p117">L??l?ment **H?tes** contient un ou plusieurs ?l?ments **H?te**. Un ?l?ment **H?te** sp?cifie un h?te Office particulier. L??l?ment **H?te** contient des ?l?ments enfants qui sp?cifient les commandes de compl?ment ? afficher une fois que votre compl?ment est install? sur l?h?te Office. Pour afficher les m?mes commandes de compl?ment dans deux ou plusieurs h?tes Office diff?rents, vous devez dupliquer les ?l?ments enfants dans chaque **h?te**.</span><span class="sxs-lookup"><span data-stu-id="18220-p117">The **Hosts** element contains one or more **Host** elements. A **Host** element specifies a particular Office host. The **Host** element contains child elements that specify the add-in commands to display after your add-in is installed in that Office host. To show the same add-in commands in two or more different Office hosts, you must duplicate the child elements in each **Host**.</span></span>
       
<span data-ttu-id="18220-178">L??l?ment **DesktopFormFactor** sp?cifie les param?tres d?un compl?ment ex?cut? dans Office sur un bureau Windows et dans Office Online (dans un navigateur).</span><span class="sxs-lookup"><span data-stu-id="18220-178">The **DesktopFormFactor** element specifies the settings for an add-in that runs in Office on Windows desktop, and Office Online (in browser).</span></span>
      
<span data-ttu-id="18220-179">L?exemple suivant illustre l?utilisation des ?l?ments **H?tes**, **H?te** et **DesktopFormFactor**.</span><span class="sxs-lookup"><span data-stu-id="18220-179">The following is an example of **Hosts**, **Host**, and **DesktopFormFactor** elements.</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
  ...
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>

              <!-- information about FunctionFile and ExtensionPoint -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
  ...
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-5-add-the-functionfile-element"></a><span data-ttu-id="18220-180">?tape 5 : ajouter l??l?ment FunctionFile</span><span class="sxs-lookup"><span data-stu-id="18220-180">Step 5: Add the FunctionFile element</span></span>

<span data-ttu-id="18220-p118">L??l?ment **FunctionFile** d?finit un fichier qui contient du code JavaScript ? ex?cuter lorsqu?une commande de compl?ment utilise une action **ExecuteFunction** (reportez-vous ? [Contr?les de bouton](https://dev.office.com/reference/add-ins/manifest/control#Button-control) pour obtenir une description). L?attribut **resid** de l??l?ment **FunctionFile** est d?fini sur un fichier HTML qui inclut tous les fichiers JavaScript requis par vos commandes de compl?ment. Vous ne pouvez pas cr?er une liaison directe vers un fichier JavaScript. Vous pouvez uniquement cr?er une liaison vers un fichier HTML. Le nom du fichier est indiqu? en tant qu??l?ment **Url** dans l??l?ment **Resources**.</span><span class="sxs-lookup"><span data-stu-id="18220-p118">The **FunctionFile** element specifies a file that contains JavaScript code to run when an add-in command uses the **ExecuteFunction** action (see [Button controls](https://dev.office.com/reference/add-ins/manifest/control#Button-control) for a description). The **FunctionFile** element's **resid** attribute is set to a HTML file that includes all the JavaScript files your add-in commands require. You can't link directly to a JavaScript file. You can only link to an HTML file. The file name is specified as a **Url** element in the **Resources** element.</span></span>
        
<span data-ttu-id="18220-186">Vous trouverez ci-dessous un exemple de l??l?ment **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="18220-186">The following is an example of the **FunctionFile** element.</span></span>
  
```xml
<DesktopFormFactor>
    <FunctionFile resid="residDesktopFuncUrl" />
    <ExtensionPoint xsi:type="PrimaryCommandSurface">
      <!-- information about this extension point -->
    </ExtensionPoint> 

    <!-- You can define more than one ExtensionPoint element as needed -->
</DesktopFormFactor>
```

> [!IMPORTANT]
> <span data-ttu-id="18220-187">Assurez-vous que votre code JavaScript appelle `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="18220-187">Make sure your JavaScript code calls  `Office.initialize`.</span></span> 
   
<span data-ttu-id="18220-p119">Le code JavaScript dans le fichier HTML r?f?renc? par l??l?ment **FunctionFile** doit appeler `Office.initialize`. L??l?ment **FunctionName** (reportez-vous ? [Contr?les de bouton](https://dev.office.com/reference/add-ins/manifest/control#Button-control) pour obtenir une description) utilise les fonctions de **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="18220-p119">The JavaScript in the HTML file referenced by the **FunctionFile** element must call `Office.initialize`. The **FunctionName** element (see [Button controls](https://dev.office.com/reference/add-ins/manifest/control#Button-control) for a description) uses the functions in **FunctionFile**.</span></span>
     
<span data-ttu-id="18220-190">Le code suivant montre comment impl?menter la fonction utilis?e par **FunctionName**.</span><span class="sxs-lookup"><span data-stu-id="18220-190">The following code shows how to implement the function used by **FunctionName**.</span></span>

```javascript

<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here. 
        };
    })();

    // Your function must be in the global namespace.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === "failed") {
                    // Show error message. 
                }
                else {
                    // Show success message.
                }
            });
        
        // Calling event.completed is required. event.completed lets the platform know that processing has completed. 
        event.completed();
    }
</script>
```

> [!IMPORTANT]
> <span data-ttu-id="18220-p120">L?appel de l??l?ment **event.completed** indique que vous avez correctement g?r? l??v?nement. Lorsqu?une fonction est appel?e plusieurs fois (par exemple, lorsque l?utilisateur clique plusieurs fois sur une m?me commande de compl?ment), tous les ?v?nements sont automatiquement mis en file d?attente. Le premier ?v?nement s?ex?cute automatiquement, tandis que les autres ?v?nements restent dans la file d?attente. Lorsque votre fonction appelle **event.completed**, l?appel de la file d?attente suivant de cette fonction s?ex?cute. Vous devez impl?menter **event.completed** pour que votre fonction s?ex?cute correctement.</span><span class="sxs-lookup"><span data-stu-id="18220-p120">The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. You must implement **event.completed**, otherwise your function will not run.</span></span>
 
## <a name="step-6-add-extensionpoint-elements"></a><span data-ttu-id="18220-196">Etape 6 : ajouter des ?l?ments ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="18220-196">Step 6: Add ExtensionPoint elements</span></span>

<span data-ttu-id="18220-p121">L??l?ment **ExtensionPoint** d?finit o? les commandes de compl?ment doivent appara?tre dans l?interface utilisateur Office. Vous pouvez d?finir les ?l?ments **ExtensionPoint** avec ces valeurs **xsi:type** :</span><span class="sxs-lookup"><span data-stu-id="18220-p121">The **ExtensionPoint** element defines where add-in commands should appear in the Office UI. You can define **ExtensionPoint** elements with these **xsi:type** values:</span></span>
   
- <span data-ttu-id="18220-199">**PrimaryCommandSurface**, qui fait r?f?rence au ruban dans Office.</span><span class="sxs-lookup"><span data-stu-id="18220-199">**PrimaryCommandSurface**, which refers to the ribbon in Office.</span></span>
     
- <span data-ttu-id="18220-200">**ContextMenu**, qui est le menu contextuel qui appara?t lorsque vous cliquez avec le bouton droit de la souris dans l?interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="18220-200">**ContextMenu**, which is the shortcut menu that appears when you right-click in the Office UI.</span></span>
    
<span data-ttu-id="18220-201">Les exemples suivants montrent comment utiliser l??l?ment **ExtensionPoint** avec les valeurs d?attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les ?l?ments enfants qui doivent ?tre utilis?s avec chacune d?elles.</span><span class="sxs-lookup"><span data-stu-id="18220-201">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>
    
> [!IMPORTANT]
> <span data-ttu-id="18220-p122">Pour les ?l?ments qui contiennent un attribut ID, veillez ? indiquer un ID unique. Nous vous recommandons d?utiliser le nom de votre organisation, ainsi que votre ID. Par exemple, utilisez le format suivant : `<CustomTab id="mycompanyname.mygroupname">`.</span><span class="sxs-lookup"><span data-stu-id="18220-p122">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.</span></span> 
  
```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso Tab">
  <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
  <!-- <OfficeTab id="TabData"> -->
    <Label resid="residLabel4" />
    <Group id="Group1Id12">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Button1Id1">
        
        <!-- information about the control -->
      </Control>   
      <!-- other controls, as needed -->                                    
    </Group>
  </CustomTab>
</ExtensionPoint>
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="ContextMenu2">
            <!-- information about the control -->
    </Control>   
    <!-- other controls, as needed -->         
  </OfficeMenu>
</ExtensionPoint>
```

|<span data-ttu-id="18220-205">**?l?ment**</span><span class="sxs-lookup"><span data-stu-id="18220-205">**Element**</span></span>|<span data-ttu-id="18220-206">**Description**</span><span class="sxs-lookup"><span data-stu-id="18220-206">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="18220-207">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="18220-207">**CustomTab**</span></span> <br/> |<span data-ttu-id="18220-p123">Obligatoire si vous souhaitez ajouter un onglet personnalis? au ruban (? l?aide de **PrimaryCommandSurface**). Si vous utilisez l??l?ment **CustomTab**, vous ne pouvez pas utiliser l??l?ment **OfficeTab**. L?attribut **id** est obligatoire. </span><span class="sxs-lookup"><span data-stu-id="18220-p123">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required. </span></span><br/> |
|<span data-ttu-id="18220-211">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="18220-211">**OfficeTab**</span></span> <br/> |<span data-ttu-id="18220-p124">Obligatoire pour ?tendre un onglet du ruban Office par d?faut (en utilisant **PrimaryCommandSurface**). Si vous utilisez l??l?ment **OfficeTab**, vous ne pouvez pas utiliser l??l?ment **CustomTab**. </span><span class="sxs-lookup"><span data-stu-id="18220-p124">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the **OfficeTab** element, you can't use the **CustomTab** element. </span></span><br/> <span data-ttu-id="18220-214">Pour obtenir plus de valeurs d?onglet ? utiliser avec l?attribut **id**, reportez-vous ? la section [Valeurs des onglets du ruban Office par d?faut](https://dev.office.com/reference/add-ins/manifest/officetab).</span><span class="sxs-lookup"><span data-stu-id="18220-214">For more tab values to use with the **id** attribute, see [Tab values for default Office ribbon tabs](https://dev.office.com/reference/add-ins/manifest/officetab).</span></span>  <br/> |
|<span data-ttu-id="18220-215">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="18220-215">**OfficeMenu**</span></span> <br/> | <span data-ttu-id="18220-p125">Obligatoire pour ajouter des commandes de compl?ment ? un menu contextuel par d?faut (en utilisant **ContextMenu**). L?attribut **id** doit ?tre d?fini sur : </span><span class="sxs-lookup"><span data-stu-id="18220-p125">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="18220-p126">**ContextMenuText** pour Excel ou Word. Affiche l??l?ment dans le menu contextuel lorsque du texte est s?lectionn? et que l?utilisateur clique dessus avec le bouton droit de la souris. </span><span class="sxs-lookup"><span data-stu-id="18220-p126">**ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="18220-p127">**ContextMenuCell** pour Excel. Affiche l??l?ment dans le menu contextuel lorsque l?utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul. </span><span class="sxs-lookup"><span data-stu-id="18220-p127">**ContextMenuCell** for Excel. Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet. </span></span><br/> |
|<span data-ttu-id="18220-222">**Group**</span><span class="sxs-lookup"><span data-stu-id="18220-222">**Group**</span></span> <br/> |<span data-ttu-id="18220-p128">Groupe de points d?extension de l?interface utilisateur sur un onglet. Un groupe peut contenir jusqu?? six contr?les. L?attribut **id** est obligatoire. Il s?agit d?une cha?ne avec un maximum de 125 caract?res. </span><span class="sxs-lookup"><span data-stu-id="18220-p128">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="18220-226">**Label**</span><span class="sxs-lookup"><span data-stu-id="18220-226">**Label**</span></span> <br/> |<span data-ttu-id="18220-p129">Obligatoire. L??tiquette du groupe. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p129">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="18220-231">**Icon**</span><span class="sxs-lookup"><span data-stu-id="18220-231">**Icon**</span></span> <br/> |<span data-ttu-id="18220-p130">Obligatoire. Sp?cifie l?ic?ne du groupe ? utiliser sur de petits appareils, ou lorsqu?un nombre trop important de boutons est affich?. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Image**. **Image** est un enfant de l??l?ment **Images**, qui est lui-m?me un enfant de l??l?ment **Ressources**. L?attribut **size** donne la taille, en pixels, de l?image. Trois tailles d?images sont obligatoires : 16, 32 et 80. 5 tailles facultatives sont ?galement prises en charge : 20, 24, 40, 48 et 64. </span><span class="sxs-lookup"><span data-stu-id="18220-p130">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="18220-239">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="18220-239">**Tooltip**</span></span> <br/> |<span data-ttu-id="18220-p131">Facultatif. Info-bulle du groupe. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p131">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="18220-244">**Control**</span><span class="sxs-lookup"><span data-stu-id="18220-244">**Control**</span></span> <br/> |<span data-ttu-id="18220-p132">Chaque groupe exige au moins un contr?le. Un ?l?ment **Control** peut ?tre de type **Button** ou **Menu**. Utilisez **Menu** pour sp?cifier une liste d?roulante de contr?les de bouton. Actuellement, seuls les boutons et les menus sont pris en charge. Pour plus d?informations, reportez-vous aux sections [Contr?les de bouton](https://dev.office.com/reference/add-ins/manifest/control) et [Contr?les de menu](https://dev.office.com/reference/add-ins/manifest/control). </span><span class="sxs-lookup"><span data-stu-id="18220-p132">Each group requires at least one control. A **Control** element can be either a **Button** or a **Menu**. Use **Menu** to specify a drop-down list of button controls. Currently, only buttons and menus are supported. See the  [Button controls](https://dev.office.com/reference/add-ins/manifest/control) and [Menu controls](https://dev.office.com/reference/add-ins/manifest/control) sections for more information. </span></span><br/><span data-ttu-id="18220-250">**Remarque :** pour faciliter les op?rations de d?pannage, nous vous recommandons d?ajouter un ?l?ment **Control** et les ?l?ments enfants **Resources** associ?s un par un.</span><span class="sxs-lookup"><span data-stu-id="18220-250">**Note:** To make troubleshooting easier, we recommend that you add a **Control** element and the related **Resources** child elements one at a time.</span></span>          |
   

### <a name="button-controls"></a><span data-ttu-id="18220-251">Contr?les de bouton</span><span class="sxs-lookup"><span data-stu-id="18220-251">Button controls</span></span>
<span data-ttu-id="18220-p133">Un bouton effectue une action unique quand il est s?lectionn?. Il peut ex?cuter une fonction JavaScript ou afficher un volet de t?ches. L?exemple suivant montre comment d?finir deux boutons. Le premier bouton ex?cute une fonction JavaScript sans afficher d?interface utilisateur et le deuxi?me bouton affiche un volet de t?ches. Dans l??l?ment **Contr?le** :</span><span class="sxs-lookup"><span data-stu-id="18220-p133">A button performs a single action when the user selects it. It can either execute a JavaScript function or show a task pane. The following example shows how to define two buttons. The first button runs a JavaScript function without showing a UI, and the second button shows a task pane. In the **Control** element:</span></span>        

- <span data-ttu-id="18220-257">l?attribut **type** est obligatoire et doit ?tre d?fini sur **Button**.</span><span class="sxs-lookup"><span data-stu-id="18220-257">The **type** attribute is required, and must be set to **Button**.</span></span>
    
- <span data-ttu-id="18220-258">l?attribut ** id** de l??l?ment **Contr?le** est une cha?ne avec un maximum de 125 caract?res.</span><span class="sxs-lookup"><span data-stu-id="18220-258">The **id** attribute of the **Control** element is a string with a maximum of 125 characters.</span></span>
    
```xml
<!-- Define a control that calls a JavaScript function. -->
<Control xsi:type="Button" id="Button1Id1">
  <Label resid="residLabel" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getData</FunctionName>
  </Action>
</Control>

<!-- Define a control that shows a task pane. -->
<Control xsi:type="Button" id="Button2Id1">
  <Label resid="residLabel2" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon2_32x32" />
    <bt:Image size="32" resid="icon2_32x32" />
    <bt:Image size="80" resid="icon2_32x32" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="residUnitConverterUrl" />
  </Action>
</Control>
```

|<span data-ttu-id="18220-259">**?l?ments**</span><span class="sxs-lookup"><span data-stu-id="18220-259">**Elements**</span></span>|<span data-ttu-id="18220-260">**Description**</span><span class="sxs-lookup"><span data-stu-id="18220-260">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="18220-261">**Label**</span><span class="sxs-lookup"><span data-stu-id="18220-261">**Label**</span></span> <br/> |<span data-ttu-id="18220-p134">Obligatoire. Texte du bouton. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p134">Required. The text for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="18220-266">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="18220-266">**Tooltip**</span></span> <br/> |<span data-ttu-id="18220-p135">Facultatif. Info-bulle pour le bouton. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p135">Optional. The tooltip for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="18220-271">**Supertip**</span><span class="sxs-lookup"><span data-stu-id="18220-271">**Supertip**</span></span> <br/> | <span data-ttu-id="18220-p136">Obligatoire. Info-bulle multiligne associ?e ? ce bouton, qui est d?finie de la fa?on suivante : </span><span class="sxs-lookup"><span data-stu-id="18220-p136">Required. The supertip for this button, which is defined by the following: </span></span><br/> <span data-ttu-id="18220-274">**Titre**</span><span class="sxs-lookup"><span data-stu-id="18220-274">**Title**</span></span> <br/>  <span data-ttu-id="18220-p137">Obligatoire. Texte de l?info-bulle am?lior?e. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p137">Required. The text for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="18220-279">**Description**</span><span class="sxs-lookup"><span data-stu-id="18220-279">**Description**</span></span> <br/>  <span data-ttu-id="18220-p138">Obligatoire. Description de l?info-bulle. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p138">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="18220-284">**Icon**</span><span class="sxs-lookup"><span data-stu-id="18220-284">**Icon**</span></span> <br/> | <span data-ttu-id="18220-p139">Obligatoire. Contient les ?l?ments **Image** pour le bouton. Les fichiers image doivent ?tre au format .png. </span><span class="sxs-lookup"><span data-stu-id="18220-p139">Required. Contains the **Image** elements for the button. Image files must be .png format. </span></span><br/> <span data-ttu-id="18220-288">**Image**</span><span class="sxs-lookup"><span data-stu-id="18220-288">**Image**</span></span> <br/>  <span data-ttu-id="18220-p140">D?finit une image ? afficher sur le bouton. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Image**. **Image** est un enfant de l??l?ment **Images**, qui est lui-m?me un enfant de l??l?ment **Ressources**. L?attribut **size** indique la taille, en pixels, de l?image. Trois tailles d?images sont obligatoires : 16, 32 et 80. 5 tailles facultatives sont ?galement prises en charge : 20, 24, 40, 48 et 64. </span><span class="sxs-lookup"><span data-stu-id="18220-p140">Defines an image to display on the button. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="18220-295">**Action**</span><span class="sxs-lookup"><span data-stu-id="18220-295">**Action**</span></span> <br/> | <span data-ttu-id="18220-p141">Obligatoire. Indique l?action ? r?aliser lorsque l?utilisateur s?lectionne le bouton. Vous pouvez sp?cifier une des valeurs suivantes pour l?attribut **xsi:type** : </span><span class="sxs-lookup"><span data-stu-id="18220-p141">Required. Specifies the action to perform when the user selects the button. You can specify one of the following values for the **xsi:type** attribute: </span></span><br/> <span data-ttu-id="18220-p142">**ExecuteFunction**, qui ex?cute une fonction JavaScript situ?e dans le fichier r?f?renc? par **FunctionFile**. **ExecuteFunction** n?affiche pas d?interface utilisateur. L??l?ment enfant **FunctionName** sp?cifie le nom de la fonction ? ex?cuter. </span><span class="sxs-lookup"><span data-stu-id="18220-p142">**ExecuteFunction**, which runs a JavaScript function located in the file referenced by **FunctionFile**. **ExecuteFunction** does not display a UI. The **FunctionName** child element specifies the name of the function to execute. </span></span><br/> <span data-ttu-id="18220-p143">**ShowTaskPane**, qui indique un compl?ment de volet de t?ches. L??l?ment enfant **SourceLocation** indique l?emplacement du fichier source du compl?ment de volet de t?ches ? afficher. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Url** dans l??l?ment **Urls** dans l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p143">**ShowTaskPane**, which shows a task pane add-in. The **SourceLocation** child element specifies the source file location of the task pane add-in to display. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element. </span></span><br/> |
   

### <a name="menu-controls"></a><span data-ttu-id="18220-305">Contr?les de menu</span><span class="sxs-lookup"><span data-stu-id="18220-305">Menu controls</span></span>
<span data-ttu-id="18220-306">Un contr?le de type **Menu** peut ?tre utilis? avec **PrimaryCommandSurface** ou **ContextMenu**, et permet de d?finir :</span><span class="sxs-lookup"><span data-stu-id="18220-306">A **Menu** control can be used with either **PrimaryCommandSurface** or **ContextMenu**, and defines:</span></span>
  
- <span data-ttu-id="18220-307">une option de menu de niveau racine.</span><span class="sxs-lookup"><span data-stu-id="18220-307">A root-level menu item.</span></span>
   
- <span data-ttu-id="18220-308">une liste de sous-menus.</span><span class="sxs-lookup"><span data-stu-id="18220-308">A list of submenu items.</span></span>
 
<span data-ttu-id="18220-p144">Lorsqu?il est utilis? avec **PrimaryCommandSurface**, l?option de menu de niveau racine s?affiche sous la forme d?un bouton dans le ruban. Lorsque le bouton est s?lectionn?, le sous-menu s?affiche sous la forme d?une liste d?roulante. Lorsqu?il est utilis? avec **ContextMenu**, un ?l?ment de menu avec un sous-menu est ins?r? dans le menu contextuel. Dans les deux cas, les ?l?ments individuels du sous-menu peuvent ex?cuter une fonction JavaScript ou afficher un volet de t?ches. Un seul niveau de sous-menus est pris en charge pour l?instant.</span><span class="sxs-lookup"><span data-stu-id="18220-p144">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>
       
<span data-ttu-id="18220-p145">L?exemple de code ci-dessous indique comment d?finir un ?l?ment de menu comportant deux options de sous-menu. La premi?re option de sous-menu affiche un volet de t?ches et la seconde ex?cute une fonction JavaScript. Dans l??l?ment **Control** :</span><span class="sxs-lookup"><span data-stu-id="18220-p145">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function. In the **Control** element:</span></span>
    
- <span data-ttu-id="18220-317">l?attribut **xsi:type** est obligatoire et doit ?tre d?fini sur **Menu**.</span><span class="sxs-lookup"><span data-stu-id="18220-317">The **xsi:type** attribute is required, and must be set to **Menu**.</span></span>
  
- <span data-ttu-id="18220-318">L?attribut **id** est une cha?ne avec un maximum de 125 caract?res.</span><span class="sxs-lookup"><span data-stu-id="18220-318">The **id** attribute is a string with a maximum of 125 characters.</span></span>
    
```xml

<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

|<span data-ttu-id="18220-319">**?l?ments**</span><span class="sxs-lookup"><span data-stu-id="18220-319">**Elements**</span></span>|<span data-ttu-id="18220-320">**Description**</span><span class="sxs-lookup"><span data-stu-id="18220-320">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="18220-321">**Label**</span><span class="sxs-lookup"><span data-stu-id="18220-321">**Label**</span></span> <br/> |<span data-ttu-id="18220-p146">Obligatoire. Texte de l??l?ment de menu racine. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p146">Required. The text of the root menu item. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="18220-326">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="18220-326">**Tooltip**</span></span> <br/> |<span data-ttu-id="18220-p147">Facultatif. Info-bulle du menu. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p147">Optional. The tooltip for the menu. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="18220-331">**Info-bulle am?lior?e**</span><span class="sxs-lookup"><span data-stu-id="18220-331">**SuperTip**</span></span> <br/> | <span data-ttu-id="18220-p148">Obligatoire. Info-bulle multiligne associ?e au menu, qui est d?finie de la fa?on suivante : </span><span class="sxs-lookup"><span data-stu-id="18220-p148">Required. The supertip for the menu, which is defined by the following: </span></span><br/> <span data-ttu-id="18220-334">**Titre**</span><span class="sxs-lookup"><span data-stu-id="18220-334">**Title**</span></span> <br/>  <span data-ttu-id="18220-p149">Obligatoire. Texte de l?info-bulle am?lior?e. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p149">Required. The text of the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="18220-339">**Description**</span><span class="sxs-lookup"><span data-stu-id="18220-339">**Description**</span></span> <br/>  <span data-ttu-id="18220-p150">Obligatoire. Description de l?info-bulle. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="18220-p150">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="18220-344">**Icon**</span><span class="sxs-lookup"><span data-stu-id="18220-344">**Icon**</span></span> <br/> | <span data-ttu-id="18220-p151">Obligatoire. Contient les ?l?ments **Image** du menu. Les fichiers image doivent ?tre au format .png. </span><span class="sxs-lookup"><span data-stu-id="18220-p151">Required. Contains the **Image** elements for the menu. Image files must be .png format. </span></span><br/> <span data-ttu-id="18220-348">**Image**</span><span class="sxs-lookup"><span data-stu-id="18220-348">**Image**</span></span> <br/>  <span data-ttu-id="18220-p152">Image du menu. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Image**. **Image** est un enfant de l??l?ment **Images**, qui est lui-m?me un enfant de l??l?ment **Ressources**. L?attribut **size** indique la taille, en pixels, de l?image. Trois tailles d?image, en pixels, sont n?cessaires : 16, 32 et 80. 5 tailles facultatives, en pixels, sont ?galement prises en charge : 20, 24, 40, 48 et 64. </span><span class="sxs-lookup"><span data-stu-id="18220-p152">An image for the menu. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes, in pixels, are required: 16, 32, and 80. Five optional sizes, in pixels, are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="18220-355">**?l?ments**</span><span class="sxs-lookup"><span data-stu-id="18220-355">**Items**</span></span> <br/> |<span data-ttu-id="18220-p153">Obligatoire. Contient les ?l?ments **?l?ment** pour chaque ?l?ment de sous-menu. Chaque ?l?ment **?l?ment** contient les m?mes ?l?ments enfants que les [contr?les de bouton](https://dev.office.com/reference/add-ins/manifest/control).  </span><span class="sxs-lookup"><span data-stu-id="18220-p153">Required. Contains the **Item** elements for each submenu item. Each **Item** element contains the same child elements as [Button controls](https://dev.office.com/reference/add-ins/manifest/control).  </span></span><br/> |
   
## <a name="step-7-add-the-resources-element"></a><span data-ttu-id="18220-359">?tape 7 : ajouter l??l?ment Resources</span><span class="sxs-lookup"><span data-stu-id="18220-359">Step 7: Add the Resources element</span></span>

<span data-ttu-id="18220-p154">L??l?ment **Ressources** contient des ressources utilis?es par les diff?rents ?l?ments enfants de l??l?ment **VersionOverrides**. Les ressources incluent des ic?nes, des cha?nes et des URL. Un ?l?ment du manifeste peut utiliser une ressource en r?f?ren?ant l?**id** de la ressource. L?utilisation de l?**id** permet d?organiser le manifeste, en particulier lorsqu?il existe des versions diff?rentes de la ressource pour diff?rents param?tres r?gionaux. Un **id** doit comporter 32 caract?res au maximum.</span><span class="sxs-lookup"><span data-stu-id="18220-p154">The **Resources** element contains resources used by the different child elements of the **VersionOverrides** element. Resources include icons, strings, and URLs. An element in the manifest can use a resource by referencing the **id** of the resource. Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales. An **id** has a maximum of 32 characters.</span></span>
  
    
    
<span data-ttu-id="18220-p155">L?exemple suivant montre un exemple de l?utilisation de l??l?ment **Ressources**. Chaque ressource peut avoir plusieurs ?l?ments enfants **Override** afin que vous puissiez d?finir une ressource diff?rente pour un param?tre r?gional sp?cifique.</span><span class="sxs-lookup"><span data-stu-id="18220-p155">The following shows an example of how to use the **Resources** element. Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>


```xml
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp16-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp32-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp80-icon_default.png" />
    </bt:Image>        
  </bt:Images>
  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
    </bt:Url>        
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="residLabel" DefaultValue="GetData">
      <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
    </bt:String>      
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="residToolTip" DefaultValue="Get data for your document.">
      <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
    </bt:String>
  </bt:LongStrings>
</Resources>
```

|<span data-ttu-id="18220-367">**Ressource**</span><span class="sxs-lookup"><span data-stu-id="18220-367">**Resource**</span></span>|<span data-ttu-id="18220-368">**Description**</span><span class="sxs-lookup"><span data-stu-id="18220-368">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="18220-369">**Images**/ **Image**</span><span class="sxs-lookup"><span data-stu-id="18220-369">**Images**/ **Image**</span></span> <br/> | <span data-ttu-id="18220-p156">Fournit l?URL HTTPS d?un fichier image. Chaque image doit d?finir les trois tailles d?image obligatoires :</span><span class="sxs-lookup"><span data-stu-id="18220-p156">Provides the HTTPS URL to an image file. Each image must define the three required image sizes:</span></span> <br/>  <span data-ttu-id="18220-372">16 x 16</span><span class="sxs-lookup"><span data-stu-id="18220-372">16?16</span></span> <br/>  <span data-ttu-id="18220-373">32 x 32</span><span class="sxs-lookup"><span data-stu-id="18220-373">32?32</span></span> <br/>  <span data-ttu-id="18220-374">80 ? 80</span><span class="sxs-lookup"><span data-stu-id="18220-374">80?80</span></span> <br/>  <span data-ttu-id="18220-375">Les tailles d?image suivantes sont ?galement prises en charge, mais ne sont pas obligatoires :</span><span class="sxs-lookup"><span data-stu-id="18220-375">The following image sizes are also supported, but not required:</span></span> <br/>  <span data-ttu-id="18220-376">20 ? 20</span><span class="sxs-lookup"><span data-stu-id="18220-376">20?20</span></span> <br/>  <span data-ttu-id="18220-377">24 ? 24</span><span class="sxs-lookup"><span data-stu-id="18220-377">24?24</span></span> <br/>  <span data-ttu-id="18220-378">40 ? 40</span><span class="sxs-lookup"><span data-stu-id="18220-378">40?40</span></span> <br/>  <span data-ttu-id="18220-379">48 ? 48</span><span class="sxs-lookup"><span data-stu-id="18220-379">48?48</span></span> <br/>  <span data-ttu-id="18220-380">64 x 64</span><span class="sxs-lookup"><span data-stu-id="18220-380">64?64</span></span> <br/> |
|<span data-ttu-id="18220-381">**URL**/ **Url**</span><span class="sxs-lookup"><span data-stu-id="18220-381">**Urls**/ **Url**</span></span> <br/> |<span data-ttu-id="18220-p157">Indique un emplacement d?URL HTTPS. Une URL peut comporter 2 048 caract?res au maximum.</span><span class="sxs-lookup"><span data-stu-id="18220-p157">Provides an HTTPS URL location. A URL can be a maximum of 2048 characters.</span></span>  <br/> |
|<span data-ttu-id="18220-384">**ShortStrings**/ **Cha?ne**</span><span class="sxs-lookup"><span data-stu-id="18220-384">**ShortStrings**/ **String**</span></span> <br/> |<span data-ttu-id="18220-p158">Texte pour les ?l?ments **Label** et **Title**. Chaque ?l?ment **String** comporte 125 caract?res au maximum. </span><span class="sxs-lookup"><span data-stu-id="18220-p158">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="18220-387">**LongStrings**/ **Cha?ne**</span><span class="sxs-lookup"><span data-stu-id="18220-387">**LongStrings**/ **String**</span></span> <br/> |<span data-ttu-id="18220-p159">Texte des ?l?ments **Tooltip** et **Description**. Chaque ?l?ment **String** contient un maximum de 250 caract?res. </span><span class="sxs-lookup"><span data-stu-id="18220-p159">The text for **Tooltip** and **Description** elements. Each **String** contains a maximum of 250 characters. </span></span><br/> |
   
> [!NOTE] 
> <span data-ttu-id="18220-390">Vous devez utiliser le protocole SSL (Secure Sockets Layer) pour toutes les URL dans les ?l?ments **Image** et **Url**.</span><span class="sxs-lookup"><span data-stu-id="18220-390">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="tab-values-for-default-office-ribbon-tabs"></a><span data-ttu-id="18220-391">Valeurs des onglets du ruban Office par d?faut</span><span class="sxs-lookup"><span data-stu-id="18220-391">Tab values for default Office ribbon tabs</span></span>
<span data-ttu-id="18220-p160">Dans Excel et Word, vous pouvez ajouter vos commandes de compl?ment au ruban en utilisant les onglets de l?interface utilisateur Office par d?faut. Le tableau ci-dessous contient les valeurs que vous pouvez utiliser pour l?attribut **id** de l??l?ment **OfficeTab**. Les valeurs des onglets respectent la casse.</span><span class="sxs-lookup"><span data-stu-id="18220-p160">In Excel and Word, you can add your add-in commands to the ribbon by using the default Office UI tabs. The following table lists the values that you can use for the **id** attribute of the **OfficeTab** element. The tab values are case sensitive.</span></span>

|<span data-ttu-id="18220-395">**Application h?te Office**</span><span class="sxs-lookup"><span data-stu-id="18220-395">**Office host application**</span></span>|<span data-ttu-id="18220-396">**Valeurs des onglets**</span><span class="sxs-lookup"><span data-stu-id="18220-396">**Tab values**</span></span>|
|:-----|:-----|
|<span data-ttu-id="18220-397">Excel</span><span class="sxs-lookup"><span data-stu-id="18220-397">Excel</span></span>  <br/> |<span data-ttu-id="18220-398">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span><span class="sxs-lookup"><span data-stu-id="18220-398">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span></span> <br/> |
|<span data-ttu-id="18220-399">Word</span><span class="sxs-lookup"><span data-stu-id="18220-399">Word</span></span>  <br/> |<span data-ttu-id="18220-400">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span><span class="sxs-lookup"><span data-stu-id="18220-400">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span></span> <br/> |
|<span data-ttu-id="18220-401">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="18220-401">PowerPoint</span></span>  <br/> |<span data-ttu-id="18220-402">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span><span class="sxs-lookup"><span data-stu-id="18220-402">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span></span>          <br/> |
   
## <a name="see-also"></a><span data-ttu-id="18220-403">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="18220-403">See also</span></span>

-  [<span data-ttu-id="18220-404">Commandes de compl?ment pour Excel, Word et PowerPoint</span><span class="sxs-lookup"><span data-stu-id="18220-404">Add-in commands for Excel, Word and PowerPoint</span></span>](../design/add-in-commands.md)      
