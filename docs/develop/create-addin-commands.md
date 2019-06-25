---
title: Création de commandes de complément dans votre manifeste pour Excel, Word et PowerPoint
description: Utilisez VersionOverrides dans votre manifeste pour définir des commandes de complément pour Excel, Word et PowerPoint. Utilisez les commandes de complément pour créer des éléments d’interface utilisateur, ajouter des boutons ou des listes, et effectuer des actions.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 9917eaa7b28ea843703a1de566b41277517b20fa
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128177"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-word-and-powerpoint"></a>Création de commandes de complément dans votre manifeste pour Excel, Word et PowerPoint


Utilisez **[VersionOverrides](/office/dev/add-ins/reference/manifest/versionoverrides)** dans votre manifeste pour définir des commandes de complément pour Excel, Word et PowerPoint. Les commandes de complément sont un moyen de personnaliser facilement l’interface utilisateur Office par défaut en y ajoutant des éléments d’interface de votre choix qui exécutent des actions. Vous pouvez utiliser les commandes de complément pour :
- créer des éléments d’interface utilisateur ou des points d’entrée qui facilitent l’utilisation des fonctionnalités de votre complément ;  
  
- ajouter des boutons ou une liste déroulante de boutons sur le ruban ;
  
- ajouter des options de menu individuelles (pouvant chacune contenir des sous-menus) à des menus contextuels spécifiques ;
  
- exécuter des actions lorsque vous avez choisi une commande de complément. Vous pouvez effectuer les opérations suivantes :

  - afficher des compléments de volet de tâches avec lesquels les utilisateurs peuvent interagir. Dans votre complément de volet de tâches, vous pouvez afficher le code HTML qui utilise la structure de l’interface utilisateur Office pour créer une interface utilisateur personnalisée ;

     *ou*

  - exécuter du code JavaScript, ce qui se fait normalement sans afficher d’interface utilisateur ;

Cet article explique comment modifier un manifeste pour définir des commandes de complément. Le schéma suivant illustre la hiérarchie des éléments utilisés pour définir des commandes de complément. Ces éléments sont décrits plus en détail dans cet article. 

L’image ci-après est une présentation des éléments de commandes de complément dans le fichier manifeste. ![Présentation des éléments de commandes de complément dans le manifeste](../images/version-overrides.png)

## <a name="step-1-start-from-a-sample"></a>Étape 1 : démarrer à partir d’un exemple

Nous vous recommandons vivement de commencer à partir d’un des exemples que nous fournissons sur la page des [exemples de commandes de complément Office](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Si vous le souhaitez, vous pouvez créer votre propre manifeste en suivant les étapes décrites dans ce guide. Vous pouvez valider votre manifeste à l’aide du fichier XSD sur le site des exemples de commandes de complément Office. Assurez-vous que vous avez lu la rubrique [Commandes de complément pour Excel, Word et PowerPoint](../design/add-in-commands.md) avant d’utiliser les commandes de complément.

## <a name="step-2-create-a-task-pane-add-in"></a>Étape 2 : créer un complément de volet Office

Pour utiliser les commandes de complément, vous devez tout d’abord créer un complément de volet Office, puis modifier le manifeste du complément, comme décrit dans cet article. Vous ne pouvez pas utiliser de commandes de complément avec les compléments de contenu. Si vous mettez à jour un manifeste existant, vous devez ajouter les **espaces de noms XML** appropriés, ainsi que l’élément **VersionOverrides** au manifeste, comme décrit à l’[étape 3 : Ajoutez l’élément VersionOverrides](#step-3-add-versionoverrides-element).

L’exemple suivant illustre le manifeste d’un complément Office 2013. Ce manifeste ne contient pas de commande de complément car il n’y a pas d’élément **VersionOverrides**. Office 2013 ne prend pas en charge les commandes de complément mais, en ajoutant **VersionOverrides** à ce manifeste, votre complément s’exécute dans Office 2013 et Office 2016. Dans Office 2013, votre complément n’affiche pas les commandes de complément et utilise la valeur **SourceLocation** pour exécuter votre complément sous la forme d’un complément de volet de tâches unique. Dans Office 2016, si aucun élément **VersionOverrides** n’est inclus, **SourceLocation** est utilisé pour exécuter votre complément. Cependant, si vous incluez **VersionOverrides**, votre complément affiche uniquement les commandes de complément et n’affiche pas votre complément sous la forme d’un complément de volet de tâches unique.
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
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

## <a name="step-3-add-versionoverrides-element"></a>Étape 3 : ajouter un élément VersionOverrides

L’élément **VersionOverrides** est l’élément racine qui contient la définition de votre commande de complément. **VersionOverrides** est un élément enfant de l’élément **OfficeApp** dans le manifeste. Le tableau suivant répertorie les attributs de l’élément **VersionOverrides**.

|**Attribut**|**Description**|
|:-----|:-----|
|**xmlns** <br/> | Obligatoire. Emplacement du schéma, qui doit être `http://schemas.microsoft.com/office/taskpaneappversionoverrides`. <br/> |
|**xsi:type** <br/> |Obligatoire. Version du schéma. La version décrite dans cet article est « VersionOverridesV1_0 ».  <br/> |

Le tableau suivant présente les éléments enfants de **VersionOverrides**.
  
|**Élément**|**Description**|
|:-----|:-----|
|**Description** <br/> |Facultatif. Décrit le complément. Cet élément **Description** enfant remplace un élément **Description** précédent dans la partie parent du manifeste. L’attribut **resid** pour cet élément **Description** est défini sur l’**id** d’un élément **Chaîne**. L’élément **Chaîne** contient le texte pour la **description**. <br/> |
|**Configuration requise** <br/> |Facultatif. Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cet élément **Configuration requise** enfant remplace l’élément **Configuration requise** dans la partie parent du manifeste. Pour plus d’informations, consultez la rubrique [Spécifier les hôtes Office et la configuration requise d’API](../develop/specify-office-hosts-and-api-requirements.md).  <br/> |
|**Hôtes** <br/> |Obligatoire. Spécifie une collection d’hôtes d’Office. L’élément **Hôtes** enfant remplace l’élément **Hôtes** dans la partie parent du manifeste. Vous devez inclure un ensemble d’attributs **xsi:type** à « Classeur » ou « Document ». <br/> |
|**Ressources** <br/> |Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste. Par exemple, la valeur de l’élément **Description** fait référence à un élément enfant dans **Ressources**. L’élément **Ressources** est décrit à l’[étape 7 : ajouter l’élément Ressources](#step-7-add-the-resources-element), plus loin dans cet article. <br/> |

L’exemple suivant montre comment utiliser l’élément **VersionOverrides** et ses éléments enfants.

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

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a>Étape 4 : ajouter des éléments Hosts, Host et DesktopFormFactor

L’élément **Hôtes** contient un ou plusieurs éléments **Hôte**. Un élément **Hôte** spécifie un hôte Office particulier. L’élément **Hôte** contient des éléments enfants qui spécifient les commandes de complément à afficher une fois que votre complément est installé sur l’hôte Office. Pour afficher les mêmes commandes de complément dans deux ou plusieurs hôtes Office différents, vous devez dupliquer les éléments enfants dans chaque **hôte**.

L’élément **DesktopFormFactor** spécifie les paramètres d’un complément exécuté dans Office sur le web (dans un navigateur) et Windows.

L’exemple suivant illustre l’utilisation des éléments **Hôtes**, **Hôte** et **DesktopFormFactor**.

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

## <a name="step-5-add-the-functionfile-element"></a>Étape 5 : ajouter l’élément FunctionFile

L’élément **FunctionFile** définit un fichier qui contient du code JavaScript à exécuter lorsqu’une commande de complément utilise une action **ExecuteFunction** (reportez-vous à [Contrôles de bouton](/office/dev/add-ins/reference/manifest/control#button-control) pour obtenir une description). L’attribut **resid** de l’élément **FunctionFile** est défini sur un fichier HTML qui inclut tous les fichiers JavaScript requis par vos commandes de complément. Vous ne pouvez pas créer une liaison directe vers un fichier JavaScript. Vous pouvez uniquement créer une liaison vers un fichier HTML. Le nom du fichier est indiqué en tant qu’élément **Url** dans l’élément **Resources**.

Vous trouverez ci-dessous un exemple de l’élément **FunctionFile**.
  
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

Le code JavaScript dans le fichier HTML référencé par l’élément **FunctionFile** doit appeler `Office.initialize`. L’élément **FunctionName** (reportez-vous à [Contrôles de bouton](/office/dev/add-ins/reference/manifest/control#button-control) pour obtenir une description) utilise les fonctions de **FunctionFile**.

Le code suivant montre comment implémenter la fonction utilisée par **FunctionName**.

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
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
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
> L’appel de l’élément **event.completed** indique que vous avez correctement géré l’événement. Lorsqu’une fonction est appelée plusieurs fois (par exemple, lorsque l’utilisateur clique plusieurs fois sur une même commande de complément), tous les événements sont automatiquement mis en file d’attente. Le premier événement s’exécute automatiquement, tandis que les autres événements restent dans la file d’attente. Lorsque votre fonction appelle **event.completed**, l’appel de la file d’attente suivant de cette fonction s’exécute. Vous devez implémenter **event.completed** pour que votre fonction s’exécute correctement.

## <a name="step-6-add-extensionpoint-elements"></a>Etape 6 : ajouter des éléments ExtensionPoint

L’élément **ExtensionPoint** définit où les commandes de complément doivent apparaître dans l’interface utilisateur Office. Vous pouvez définir les éléments **ExtensionPoint** avec ces valeurs **xsi:type** :

- **PrimaryCommandSurface**, qui fait référence au ruban dans Office.

- **ContextMenu**, qui est le menu contextuel qui apparaît lorsque vous cliquez avec le bouton droit de la souris dans l’interface utilisateur Office.

Les exemples suivants montrent comment utiliser l’élément **ExtensionPoint** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.

> [!IMPORTANT]
> Pour les éléments qui contiennent un attribut ID, veillez à indiquer un ID unique. Nous vous recommandons d’utiliser le nom de votre organisation, ainsi que votre ID. Par exemple, utilisez le format suivant : `<CustomTab id="mycompanyname.mygroupname">`. 
  
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

|**Élément**|**Description**|
|:-----|:-----|
|**CustomTab** <br/> |Obligatoire si vous souhaitez ajouter un onglet personnalisé au ruban (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **CustomTab**, vous ne pouvez pas utiliser l’élément **OfficeTab**. L’attribut **id** est obligatoire. <br/> |
|**OfficeTab** <br/> |Obligatoire pour étendre un onglet du ruban Office par défaut (en utilisant **PrimaryCommandSurface**). Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**. <br/> Pour obtenir plus de valeurs d’onglet à utiliser avec l’attribut **id**, reportez-vous à la section [Valeurs des onglets du ruban Office par défaut](/office/dev/add-ins/reference/manifest/officetab).  <br/> |
|**OfficeMenu** <br/> | Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : <br/> **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. <br/> **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul. <br/> |
|**Group** <br/> |Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est obligatoire. Il s’agit d’une chaîne avec un maximum de 125 caractères. <br/> |
|**Label** <br/> |Obligatoire. L’étiquette du groupe. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> |
|**Icon** <br/> |Obligatoire. Spécifie l’icône du groupe à utiliser sur de petits appareils, ou lorsqu’un nombre trop important de boutons est affiché. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image**. **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément **Ressources**. L’attribut **size** donne la taille, en pixels, de l’image. Trois tailles d’images sont obligatoires : 16, 32 et 80. 5 tailles facultatives sont également prises en charge : 20, 24, 40, 48 et 64. <br/> |
|**Tooltip** <br/> |Facultatif. Info-bulle du groupe. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> |
|**Control** <br/> |Chaque groupe exige au moins un contrôle. Un élément **Control** peut être de type **Button** ou **Menu**. Utilisez **Menu** pour spécifier une liste déroulante de contrôles de bouton. Actuellement, seuls les boutons et les menus sont pris en charge. Pour plus d’informations, reportez-vous aux sections [Contrôles de bouton](/office/dev/add-ins/reference/manifest/control#button-control) et [Contrôles de menu](/office/dev/add-ins/reference/manifest/control#menu-dropdown-button-controls). <br/>**Remarque :** pour faciliter les opérations de dépannage, nous vous recommandons d’ajouter un élément **Control** et les éléments enfants **Resources** associés un par un.          |


### <a name="button-controls"></a>Contrôles de bouton

Un bouton effectue une action unique quand il est sélectionné. Il peut exécuter une fonction JavaScript ou afficher un volet de tâches. L’exemple suivant montre comment définir deux boutons. Le premier bouton exécute une fonction JavaScript sans afficher d’interface utilisateur et le deuxième bouton affiche un volet de tâches. Dans l’élément **Contrôle** :

- l’attribut **type** est obligatoire et doit être défini sur **Button**.

- l’attribut ** id** de l’élément **Contrôle** est une chaîne avec un maximum de 125 caractères.

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

|**Éléments**|**Description**|
|:-----|:-----|
|**Label** <br/> |Obligatoire. Texte du bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> |
|**Tooltip** <br/> |Facultatif. Info-bulle pour le bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> |
|**Supertip** <br/> | Obligatoire. Info-bulle multiligne associée à ce bouton, qui est définie de la façon suivante : <br/> **Titre** <br/>  Obligatoire. Texte de l’info-bulle améliorée. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> **Description** <br/>  Obligatoire. Description de l’info-bulle. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> |
|**Icon** <br/> | Obligatoire. Contient les éléments **Image** pour le bouton. Les fichiers image doivent être au format .png. <br/> **Image** <br/>  Définit une image à afficher sur le bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image**. **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément **Ressources**. L’attribut **size** indique la taille, en pixels, de l’image. Trois tailles d’images sont obligatoires : 16, 32 et 80. 5 tailles facultatives sont également prises en charge : 20, 24, 40, 48 et 64. <br/> |
|**Action** <br/> | Obligatoire. Indique l’action à réaliser lorsque l’utilisateur sélectionne le bouton. Vous pouvez spécifier une des valeurs suivantes pour l’attribut **xsi:type** : <br/> **ExecuteFunction**, qui exécute une fonction JavaScript située dans le fichier référencé par **FunctionFile**. **ExecuteFunction** n’affiche pas d’interface utilisateur. L’élément enfant **FunctionName** spécifie le nom de la fonction à exécuter.<br/> **ShowTaskPane**, qui indique un complément de volet de tâches. L’élément enfant **SourceLocation** indique l’emplacement du fichier source du complément de volet de tâches à afficher. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément **Urls** dans l’élément **Ressources**. <br/> |


### <a name="menu-controls"></a>Contrôles de menu
Un contrôle de type **Menu** peut être utilisé avec **PrimaryCommandSurface** ou **ContextMenu**, et permet de définir :
  
- une option de menu de niveau racine.

- une liste de sous-menus.
 
Lorsqu’il est utilisé avec **PrimaryCommandSurface**, l’option de menu de niveau racine s’affiche sous la forme d’un bouton dans le ruban. Lorsque le bouton est sélectionné, le sous-menu s’affiche sous la forme d’une liste déroulante. Lorsqu’il est utilisé avec **ContextMenu**, un élément de menu avec un sous-menu est inséré dans le menu contextuel. Dans les deux cas, les éléments individuels du sous-menu peuvent exécuter une fonction JavaScript ou afficher un volet de tâches. Un seul niveau de sous-menus est pris en charge pour l’instant.

L’exemple de code ci-dessous indique comment définir un élément de menu comportant deux options de sous-menu. La première option de sous-menu affiche un volet de tâches et la seconde exécute une fonction JavaScript. Dans l’élément **Control** :

- l’attribut **xsi:type** est obligatoire et doit être défini sur **Menu**.
  
- L’attribut **id** est une chaîne avec un maximum de 125 caractères.

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

|**Éléments**|**Description**|
|:-----|:-----|
|**Label** <br/> |Obligatoire. Texte de l’élément de menu racine. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> |
|**Tooltip** <br/> |Facultatif. Info-bulle du menu. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> |
|**Info-bulle améliorée** <br/> | Obligatoire. Info-bulle multiligne associée au menu, qui est définie de la façon suivante : <br/> **Titre** <br/>  Obligatoire. Texte de l’info-bulle améliorée. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> **Description** <br/>  Obligatoire. Description de l’info-bulle. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément **Ressources**. <br/> |
|**Icon** <br/> | Obligatoire. Contient les éléments **Image** du menu. Les fichiers image doivent être au format .png. <br/> **Image** <br/>  Image du menu. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image**. **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément **Ressources**. L’attribut **size** indique la taille, en pixels, de l’image. Trois tailles d’image, en pixels, sont nécessaires : 16, 32 et 80. 5 tailles facultatives, en pixels, sont également prises en charge : 20, 24, 40, 48 et 64. <br/> |
|**Éléments** <br/> |Obligatoire. Contient les éléments **Élément** pour chaque élément de sous-menu. Chaque élément **Élément** contient les mêmes éléments enfants que les [contrôles de bouton](/office/dev/add-ins/reference/manifest/control#button-control).  <br/> |
   
## <a name="step-7-add-the-resources-element"></a>Étape 7 : ajouter l’élément Resources

L’élément **Ressources** contient des ressources utilisées par les différents éléments enfants de l’élément **VersionOverrides**. Les ressources incluent des icônes, des chaînes et des URL. Un élément du manifeste peut utiliser une ressource en référençant l’**id** de la ressource. L’utilisation de l’**id** permet d’organiser le manifeste, en particulier lorsqu’il existe des versions différentes de la ressource pour différents paramètres régionaux. Un **id** doit comporter 32 caractères au maximum.
  
L’exemple suivant montre un exemple de l’utilisation de l’élément **Ressources**. Chaque ressource peut avoir plusieurs éléments enfants **Override** afin que vous puissiez définir une ressource différente pour un paramètre régional spécifique.


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
|**Images**/ **Image** <br/> | Fournit l’URL HTTPS d’un fichier image. Chaque image doit définir les trois tailles d’image obligatoires : <br/>  16 x 16 <br/>  32 x 32 <br/>  80 × 80 <br/>  Les tailles d’image suivantes sont également prises en charge, mais ne sont pas obligatoires : <br/>  20 × 20 <br/>  24 × 24 <br/>  40 × 40 <br/>  48 × 48 <br/>  64 x 64 <br/> |
|**URL**/ **Url** <br/> |Indique un emplacement d’URL HTTPS. Une URL peut comporter 2 048 caractères au maximum.  <br/> |
|**ShortStrings**/ **Chaîne** <br/> |Texte pour les éléments **Label** et **Title**. Chaque élément **String** comporte 125 caractères au maximum. <br/> |
|**LongStrings**/ **Chaîne** <br/> |Texte des éléments **Tooltip** et **Description**. Chaque élément **String** contient un maximum de 250 caractères. <br/> |
   > [!NOTE]
> Vous devez utiliser le protocole SSL (Secure Sockets Layer) pour toutes les URL dans les éléments **Image** et **Url**.

### <a name="tab-values-for-default-office-ribbon-tabs"></a>Valeurs des onglets du ruban Office par défaut

Dans Excel et Word, vous pouvez ajouter vos commandes de complément au ruban en utilisant les onglets de l’interface utilisateur Office par défaut. Le tableau ci-dessous contient les valeurs que vous pouvez utiliser pour l’attribut **id** de l’élément **OfficeTab**. Les valeurs des onglets respectent la casse.

|**Application hôte Office**|**Valeurs des onglets**|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |

## <a name="see-also"></a>Voir aussi

-  [Commandes de complément pour Excel, Word et PowerPoint](../design/add-in-commands.md)
