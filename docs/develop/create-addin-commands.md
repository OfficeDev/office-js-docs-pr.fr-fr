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
# <a name="create-add-in-commands-in-your-manifest-for-excel-word-and-powerpoint"></a>Cr?ation de commandes de compl?ment dans votre manifeste pour Excel, Word et PowerPoint


Utilisez **[VersionOverrides](https://dev.office.com/reference/add-ins/manifest/versionoverrides)** dans votre manifeste pour d?finir des commandes de compl?ment pour Excel, Word et PowerPoint. Les commandes de compl?ment sont un moyen de personnaliser facilement l?interface utilisateur Office par d?faut en y ajoutant des ?l?ments d?interface de votre choix qui ex?cutent des actions. Vous pouvez utiliser les commandes de compl?ment pour :
- cr?er des ?l?ments d?interface utilisateur ou des points d?entr?e qui facilitent l?utilisation des fonctionnalit?s de votre compl?ment ;  
  
- ajouter des boutons ou une liste d?roulante de boutons sur le ruban ;    
  
- ajouter des options de menu individuelles (pouvant chacune contenir des sous-menus) ? des menus contextuels sp?cifiques ;    
  
- ex?cuter des actions lorsque vous avez choisi une commande de compl?ment. Vous pouvez effectuer les op?rations suivantes :
    
  - afficher des compl?ments de volet de t?ches avec lesquels les utilisateurs peuvent interagir. Dans votre compl?ment de volet de t?ches, vous pouvez afficher le code HTML qui utilise la structure de l?interface utilisateur Office pour cr?er une interface utilisateur personnalis?e ;
    
     *ou* 
      
  - ex?cuter du code JavaScript, ce qui se fait normalement sans afficher d?interface utilisateur ;
      
Cet article explique comment modifier un manifeste pour d?finir des commandes de compl?ment. Le sch?ma suivant illustre la hi?rarchie des ?l?ments utilis?s pour d?finir des commandes de compl?ment. Ces ?l?ments sont d?crits plus en d?tail dans cet article. 
      
L?image ci-apr?s est une pr?sentation des ?l?ments de commandes de compl?ment dans le fichier manifeste. 
![Pr?sentation des ?l?ments de commandes de compl?ment dans le manifeste](../images/version-overrides.png)
 
## <a name="step-1-start-from-a-sample"></a>?tape 1 : d?marrer ? partir d?un exemple

Nous vous recommandons vivement de commencer ? partir d?un des exemples que nous fournissons sur la page des [exemples de commandes de compl?ment Office](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Si vous le souhaitez, vous pouvez cr?er votre propre manifeste en suivant les ?tapes d?crites dans ce guide. Vous pouvez valider votre manifeste ? l?aide du fichier XSD sur le site des exemples de commandes de compl?ment Office. Assurez-vous que vous avez lu la rubrique [Commandes de compl?ment pour Excel, Word et PowerPoint](../design/add-in-commands.md) avant d?utiliser les commandes de compl?ment.

## <a name="step-2-create-a-task-pane-add-in"></a>?tape 2 : cr?er un compl?ment de volet Office

Pour utiliser les commandes de compl?ment, vous devez tout d?abord cr?er un compl?ment de volet Office, puis modifier le manifeste du compl?ment, comme d?crit dans cet article. Vous ne pouvez pas utiliser de commandes de compl?ment avec les compl?ments de contenu. Si vous mettez ? jour un manifeste existant, vous devez ajouter les **espaces de noms XML** appropri?s, ainsi que l??l?ment **VersionOverrides** au manifeste, comme d?crit ? l?[?tape 3 : Ajoutez l??l?ment VersionOverrides](#step-3-add-versionoverrides-element).
   
L?exemple suivant illustre le manifeste d?un compl?ment Office 2013. Ce manifeste ne contient pas de commande de compl?ment car il n?y a pas d??l?ment **VersionOverrides**. Office 2013 ne prend pas en charge les commandes de compl?ment mais, en ajoutant **VersionOverrides** ? ce manifeste, votre compl?ment s?ex?cute dans Office 2013 et Office 2016. Dans Office 2013, votre compl?ment n?affiche pas les commandes de compl?ment et utilise la valeur **SourceLocation** pour ex?cuter votre compl?ment sous la forme d?un compl?ment de volet de t?ches unique. Dans Office 2016, si aucun ?l?ment **VersionOverrides** n?est inclus, **SourceLocation** est utilis? pour ex?cuter votre compl?ment. Cependant, si vous incluez **VersionOverrides**, votre compl?ment affiche uniquement les commandes de compl?ment et n?affiche pas votre compl?ment sous la forme d?un compl?ment de volet de t?ches unique.
  
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

## <a name="step-3-add-versionoverrides-element"></a>?tape 3 : ajouter un ?l?ment VersionOverrides
L??l?ment **VersionOverrides** est l??l?ment racine qui contient la d?finition de votre commande de compl?ment. **VersionOverrides** est un ?l?ment enfant de l??l?ment **OfficeApp** dans le manifeste. Le tableau suivant r?pertorie les attributs de l??l?ment **VersionOverrides**.

|**Attribut**|**Description**|
|:-----|:-----|
|**xmlns** <br/> | Obligatoire. Emplacement du sch?ma, qui doit ?tre ? http://schemas.microsoft.com/office/taskpaneappversionoverrides ?. <br/> |
|**xsi:type** <br/> |Obligatoire. Version du sch?ma. La version d?crite dans cet article est ? VersionOverridesV1_0 ?.  <br/> |
   
Le tableau suivant pr?sente les ?l?ments enfants de **VersionOverrides**.
  
|**?l?ment**|**Description**|
|:-----|:-----|
|**Description** <br/> |Facultatif. D?crit le compl?ment. Cet ?l?ment **Description** enfant remplace un ?l?ment **Description** pr?c?dent dans la partie parent du manifeste. L?attribut **resid** pour cet ?l?ment **Description** est d?fini sur l?**id** d?un ?l?ment **Cha?ne**. L??l?ment **Cha?ne** contient le texte pour la **description**. <br/> |
|**Configuration requise** <br/> |Facultatif. Sp?cifie l?ensemble de conditions requises minimal et la version d?Office.js qui doit ?tre activ?e par le compl?ment Office. Cet ?l?ment **Configuration requise** enfant remplace l??l?ment **Configuration requise** dans la partie parent du manifeste. Pour plus d?informations, consultez la rubrique [Sp?cifier les h?tes Office et la configuration requise d?API](../develop/specify-office-hosts-and-api-requirements.md).  <br/> |
|**H?tes** <br/> |Obligatoire. Sp?cifie une collection d?h?tes d?Office. L??l?ment **H?tes** enfant remplace l??l?ment **H?tes** dans la partie parent du manifeste. Vous devez inclure un ensemble d?attributs **xsi:type** ? ? Classeur ? ou ? Document ?. <br/> |
|**Ressources** <br/> |D?finit une collection de ressources (cha?nes, URL et images) qui sont r?f?renc?es par d?autres ?l?ments de manifeste. Par exemple, la valeur de l??l?ment **Description** fait r?f?rence ? un ?l?ment enfant dans **Ressources**. L??l?ment **Ressources** est d?crit ? l?[?tape 7 : ajouter l??l?ment Ressources](#step-7-add-the-resources-element), plus loin dans cet article. <br/> |
   
L?exemple suivant montre comment utiliser l??l?ment **VersionOverrides** et ses ?l?ments enfants.

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

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a>?tape 4 : ajouter des ?l?ments Hosts, Host et DesktopFormFactor

L??l?ment **H?tes** contient un ou plusieurs ?l?ments **H?te**. Un ?l?ment **H?te** sp?cifie un h?te Office particulier. L??l?ment **H?te** contient des ?l?ments enfants qui sp?cifient les commandes de compl?ment ? afficher une fois que votre compl?ment est install? sur l?h?te Office. Pour afficher les m?mes commandes de compl?ment dans deux ou plusieurs h?tes Office diff?rents, vous devez dupliquer les ?l?ments enfants dans chaque **h?te**.
       
L??l?ment **DesktopFormFactor** sp?cifie les param?tres d?un compl?ment ex?cut? dans Office sur un bureau Windows et dans Office Online (dans un navigateur).
      
L?exemple suivant illustre l?utilisation des ?l?ments **H?tes**, **H?te** et **DesktopFormFactor**.

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

## <a name="step-5-add-the-functionfile-element"></a>?tape 5 : ajouter l??l?ment FunctionFile

L??l?ment **FunctionFile** d?finit un fichier qui contient du code JavaScript ? ex?cuter lorsqu?une commande de compl?ment utilise une action **ExecuteFunction** (reportez-vous ? [Contr?les de bouton](https://dev.office.com/reference/add-ins/manifest/control#Button-control) pour obtenir une description). L?attribut **resid** de l??l?ment **FunctionFile** est d?fini sur un fichier HTML qui inclut tous les fichiers JavaScript requis par vos commandes de compl?ment. Vous ne pouvez pas cr?er une liaison directe vers un fichier JavaScript. Vous pouvez uniquement cr?er une liaison vers un fichier HTML. Le nom du fichier est indiqu? en tant qu??l?ment **Url** dans l??l?ment **Resources**.
        
Vous trouverez ci-dessous un exemple de l??l?ment **FunctionFile**.
  
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
> Assurez-vous que votre code JavaScript appelle `Office.initialize`. 
   
Le code JavaScript dans le fichier HTML r?f?renc? par l??l?ment **FunctionFile** doit appeler `Office.initialize`. L??l?ment **FunctionName** (reportez-vous ? [Contr?les de bouton](https://dev.office.com/reference/add-ins/manifest/control#Button-control) pour obtenir une description) utilise les fonctions de **FunctionFile**.
     
Le code suivant montre comment impl?menter la fonction utilis?e par **FunctionName**.

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
> L?appel de l??l?ment **event.completed** indique que vous avez correctement g?r? l??v?nement. Lorsqu?une fonction est appel?e plusieurs fois (par exemple, lorsque l?utilisateur clique plusieurs fois sur une m?me commande de compl?ment), tous les ?v?nements sont automatiquement mis en file d?attente. Le premier ?v?nement s?ex?cute automatiquement, tandis que les autres ?v?nements restent dans la file d?attente. Lorsque votre fonction appelle **event.completed**, l?appel de la file d?attente suivant de cette fonction s?ex?cute. Vous devez impl?menter **event.completed** pour que votre fonction s?ex?cute correctement.
 
## <a name="step-6-add-extensionpoint-elements"></a>Etape 6 : ajouter des ?l?ments ExtensionPoint

L??l?ment **ExtensionPoint** d?finit o? les commandes de compl?ment doivent appara?tre dans l?interface utilisateur Office. Vous pouvez d?finir les ?l?ments **ExtensionPoint** avec ces valeurs **xsi:type** :
   
- **PrimaryCommandSurface**, qui fait r?f?rence au ruban dans Office.
     
- **ContextMenu**, qui est le menu contextuel qui appara?t lorsque vous cliquez avec le bouton droit de la souris dans l?interface utilisateur Office.
    
Les exemples suivants montrent comment utiliser l??l?ment **ExtensionPoint** avec les valeurs d?attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les ?l?ments enfants qui doivent ?tre utilis?s avec chacune d?elles.
    
> [!IMPORTANT]
> Pour les ?l?ments qui contiennent un attribut ID, veillez ? indiquer un ID unique. Nous vous recommandons d?utiliser le nom de votre organisation, ainsi que votre ID. Par exemple, utilisez le format suivant : `<CustomTab id="mycompanyname.mygroupname">`. 
  
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

|**?l?ment**|**Description**|
|:-----|:-----|
|**CustomTab** <br/> |Obligatoire si vous souhaitez ajouter un onglet personnalis? au ruban (? l?aide de **PrimaryCommandSurface**). Si vous utilisez l??l?ment **CustomTab**, vous ne pouvez pas utiliser l??l?ment **OfficeTab**. L?attribut **id** est obligatoire. <br/> |
|**OfficeTab** <br/> |Obligatoire pour ?tendre un onglet du ruban Office par d?faut (en utilisant **PrimaryCommandSurface**). Si vous utilisez l??l?ment **OfficeTab**, vous ne pouvez pas utiliser l??l?ment **CustomTab**. <br/> Pour obtenir plus de valeurs d?onglet ? utiliser avec l?attribut **id**, reportez-vous ? la section [Valeurs des onglets du ruban Office par d?faut](https://dev.office.com/reference/add-ins/manifest/officetab).  <br/> |
|**OfficeMenu** <br/> | Obligatoire pour ajouter des commandes de compl?ment ? un menu contextuel par d?faut (en utilisant **ContextMenu**). L?attribut **id** doit ?tre d?fini sur : <br/> **ContextMenuText** pour Excel ou Word. Affiche l??l?ment dans le menu contextuel lorsque du texte est s?lectionn? et que l?utilisateur clique dessus avec le bouton droit de la souris. <br/> **ContextMenuCell** pour Excel. Affiche l??l?ment dans le menu contextuel lorsque l?utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul. <br/> |
|**Group** <br/> |Groupe de points d?extension de l?interface utilisateur sur un onglet. Un groupe peut contenir jusqu?? six contr?les. L?attribut **id** est obligatoire. Il s?agit d?une cha?ne avec un maximum de 125 caract?res. <br/> |
|**Label** <br/> |Obligatoire. L??tiquette du groupe. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> |
|**Icon** <br/> |Obligatoire. Sp?cifie l?ic?ne du groupe ? utiliser sur de petits appareils, ou lorsqu?un nombre trop important de boutons est affich?. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Image**. **Image** est un enfant de l??l?ment **Images**, qui est lui-m?me un enfant de l??l?ment **Ressources**. L?attribut **size** donne la taille, en pixels, de l?image. Trois tailles d?images sont obligatoires : 16, 32 et 80. 5 tailles facultatives sont ?galement prises en charge : 20, 24, 40, 48 et 64. <br/> |
|**Tooltip** <br/> |Facultatif. Info-bulle du groupe. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> |
|**Control** <br/> |Chaque groupe exige au moins un contr?le. Un ?l?ment **Control** peut ?tre de type **Button** ou **Menu**. Utilisez **Menu** pour sp?cifier une liste d?roulante de contr?les de bouton. Actuellement, seuls les boutons et les menus sont pris en charge. Pour plus d?informations, reportez-vous aux sections [Contr?les de bouton](https://dev.office.com/reference/add-ins/manifest/control) et [Contr?les de menu](https://dev.office.com/reference/add-ins/manifest/control). <br/>**Remarque :** pour faciliter les op?rations de d?pannage, nous vous recommandons d?ajouter un ?l?ment **Control** et les ?l?ments enfants **Resources** associ?s un par un.          |
   

### <a name="button-controls"></a>Contr?les de bouton
Un bouton effectue une action unique quand il est s?lectionn?. Il peut ex?cuter une fonction JavaScript ou afficher un volet de t?ches. L?exemple suivant montre comment d?finir deux boutons. Le premier bouton ex?cute une fonction JavaScript sans afficher d?interface utilisateur et le deuxi?me bouton affiche un volet de t?ches. Dans l??l?ment **Contr?le** :        

- l?attribut **type** est obligatoire et doit ?tre d?fini sur **Button**.
    
- l?attribut ** id** de l??l?ment **Contr?le** est une cha?ne avec un maximum de 125 caract?res.
    
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

|**?l?ments**|**Description**|
|:-----|:-----|
|**Label** <br/> |Obligatoire. Texte du bouton. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> |
|**Tooltip** <br/> |Facultatif. Info-bulle pour le bouton. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> |
|**Supertip** <br/> | Obligatoire. Info-bulle multiligne associ?e ? ce bouton, qui est d?finie de la fa?on suivante : <br/> **Titre** <br/>  Obligatoire. Texte de l?info-bulle am?lior?e. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> **Description** <br/>  Obligatoire. Description de l?info-bulle. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> |
|**Icon** <br/> | Obligatoire. Contient les ?l?ments **Image** pour le bouton. Les fichiers image doivent ?tre au format .png. <br/> **Image** <br/>  D?finit une image ? afficher sur le bouton. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Image**. **Image** est un enfant de l??l?ment **Images**, qui est lui-m?me un enfant de l??l?ment **Ressources**. L?attribut **size** indique la taille, en pixels, de l?image. Trois tailles d?images sont obligatoires : 16, 32 et 80. 5 tailles facultatives sont ?galement prises en charge : 20, 24, 40, 48 et 64. <br/> |
|**Action** <br/> | Obligatoire. Indique l?action ? r?aliser lorsque l?utilisateur s?lectionne le bouton. Vous pouvez sp?cifier une des valeurs suivantes pour l?attribut **xsi:type** : <br/> **ExecuteFunction**, qui ex?cute une fonction JavaScript situ?e dans le fichier r?f?renc? par **FunctionFile**. **ExecuteFunction** n?affiche pas d?interface utilisateur. L??l?ment enfant **FunctionName** sp?cifie le nom de la fonction ? ex?cuter. <br/> **ShowTaskPane**, qui indique un compl?ment de volet de t?ches. L??l?ment enfant **SourceLocation** indique l?emplacement du fichier source du compl?ment de volet de t?ches ? afficher. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Url** dans l??l?ment **Urls** dans l??l?ment **Ressources**. <br/> |
   

### <a name="menu-controls"></a>Contr?les de menu
Un contr?le de type **Menu** peut ?tre utilis? avec **PrimaryCommandSurface** ou **ContextMenu**, et permet de d?finir :
  
- une option de menu de niveau racine.
   
- une liste de sous-menus.
 
Lorsqu?il est utilis? avec **PrimaryCommandSurface**, l?option de menu de niveau racine s?affiche sous la forme d?un bouton dans le ruban. Lorsque le bouton est s?lectionn?, le sous-menu s?affiche sous la forme d?une liste d?roulante. Lorsqu?il est utilis? avec **ContextMenu**, un ?l?ment de menu avec un sous-menu est ins?r? dans le menu contextuel. Dans les deux cas, les ?l?ments individuels du sous-menu peuvent ex?cuter une fonction JavaScript ou afficher un volet de t?ches. Un seul niveau de sous-menus est pris en charge pour l?instant.
       
L?exemple de code ci-dessous indique comment d?finir un ?l?ment de menu comportant deux options de sous-menu. La premi?re option de sous-menu affiche un volet de t?ches et la seconde ex?cute une fonction JavaScript. Dans l??l?ment **Control** :
    
- l?attribut **xsi:type** est obligatoire et doit ?tre d?fini sur **Menu**.
  
- L?attribut **id** est une cha?ne avec un maximum de 125 caract?res.
    
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

|**?l?ments**|**Description**|
|:-----|:-----|
|**Label** <br/> |Obligatoire. Texte de l??l?ment de menu racine. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> |
|**Tooltip** <br/> |Facultatif. Info-bulle du menu. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> |
|**Info-bulle am?lior?e** <br/> | Obligatoire. Info-bulle multiligne associ?e au menu, qui est d?finie de la fa?on suivante : <br/> **Titre** <br/>  Obligatoire. Texte de l?info-bulle am?lior?e. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **ShortStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> **Description** <br/>  Obligatoire. Description de l?info-bulle. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Cha?ne**. **Cha?ne** est un enfant de l??l?ment **LongStrings**, qui est lui-m?me un enfant de l??l?ment **Ressources**. <br/> |
|**Icon** <br/> | Obligatoire. Contient les ?l?ments **Image** du menu. Les fichiers image doivent ?tre au format .png. <br/> **Image** <br/>  Image du menu. L?attribut **resid** doit ?tre d?fini sur la valeur de l?attribut **id** d?un ?l?ment **Image**. **Image** est un enfant de l??l?ment **Images**, qui est lui-m?me un enfant de l??l?ment **Ressources**. L?attribut **size** indique la taille, en pixels, de l?image. Trois tailles d?image, en pixels, sont n?cessaires : 16, 32 et 80. 5 tailles facultatives, en pixels, sont ?galement prises en charge : 20, 24, 40, 48 et 64. <br/> |
|**?l?ments** <br/> |Obligatoire. Contient les ?l?ments **?l?ment** pour chaque ?l?ment de sous-menu. Chaque ?l?ment **?l?ment** contient les m?mes ?l?ments enfants que les [contr?les de bouton](https://dev.office.com/reference/add-ins/manifest/control).  <br/> |
   
## <a name="step-7-add-the-resources-element"></a>?tape 7 : ajouter l??l?ment Resources

L??l?ment **Ressources** contient des ressources utilis?es par les diff?rents ?l?ments enfants de l??l?ment **VersionOverrides**. Les ressources incluent des ic?nes, des cha?nes et des URL. Un ?l?ment du manifeste peut utiliser une ressource en r?f?ren?ant l?**id** de la ressource. L?utilisation de l?**id** permet d?organiser le manifeste, en particulier lorsqu?il existe des versions diff?rentes de la ressource pour diff?rents param?tres r?gionaux. Un **id** doit comporter 32 caract?res au maximum.
  
    
    
L?exemple suivant montre un exemple de l?utilisation de l??l?ment **Ressources**. Chaque ressource peut avoir plusieurs ?l?ments enfants **Override** afin que vous puissiez d?finir une ressource diff?rente pour un param?tre r?gional sp?cifique.


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

|**Ressource**|**Description**|
|:-----|:-----|
|**Images**/ **Image** <br/> | Fournit l?URL HTTPS d?un fichier image. Chaque image doit d?finir les trois tailles d?image obligatoires : <br/>  16 x 16 <br/>  32 x 32 <br/>  80 ? 80 <br/>  Les tailles d?image suivantes sont ?galement prises en charge, mais ne sont pas obligatoires : <br/>  20 ? 20 <br/>  24 ? 24 <br/>  40 ? 40 <br/>  48 ? 48 <br/>  64 x 64 <br/> |
|**URL**/ **Url** <br/> |Indique un emplacement d?URL HTTPS. Une URL peut comporter 2 048 caract?res au maximum.  <br/> |
|**ShortStrings**/ **Cha?ne** <br/> |Texte pour les ?l?ments **Label** et **Title**. Chaque ?l?ment **String** comporte 125 caract?res au maximum. <br/> |
|**LongStrings**/ **Cha?ne** <br/> |Texte des ?l?ments **Tooltip** et **Description**. Chaque ?l?ment **String** contient un maximum de 250 caract?res. <br/> |
   
> [!NOTE] 
> Vous devez utiliser le protocole SSL (Secure Sockets Layer) pour toutes les URL dans les ?l?ments **Image** et **Url**.

### <a name="tab-values-for-default-office-ribbon-tabs"></a>Valeurs des onglets du ruban Office par d?faut
Dans Excel et Word, vous pouvez ajouter vos commandes de compl?ment au ruban en utilisant les onglets de l?interface utilisateur Office par d?faut. Le tableau ci-dessous contient les valeurs que vous pouvez utiliser pour l?attribut **id** de l??l?ment **OfficeTab**. Les valeurs des onglets respectent la casse.

|**Application h?te Office**|**Valeurs des onglets**|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |
   
## <a name="see-also"></a>Voir aussi

-  [Commandes de compl?ment pour Excel, Word et PowerPoint](../design/add-in-commands.md)      
