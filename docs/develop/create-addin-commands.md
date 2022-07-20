---
title: Créer des commandes complémentaires dans votre formulaire pour Excel, PowerPoint et Word
description: Utilisez VersionOverrides dans votre manifeste pour définir des commandes de complément pour Excel, PowerPoint et Word. Utilisez les commandes de complément pour créer des éléments d’interface utilisateur, ajouter des boutons ou des listes, et effectuer des actions.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 44cd5818879af6788ef58050b5ca475b5f4d3dbd
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889508"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-powerpoint-and-word"></a>Créer des commandes complémentaires dans votre formulaire pour Excel, PowerPoint et Word

> [!NOTE]
> Les commandes de complément sont actuellement prises en charge dans Outlook. Pour plus d’informations, consultez [les commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md)

Utilisez **[VersionOverrides](/javascript/api/manifest/versionoverrides)** dans votre manifeste pour définir des commandes de complément pour Excel, PowerPoint et Word. Les commandes de complément sont un moyen de personnaliser facilement l’interface utilisateur Office par défaut en y ajoutant des éléments d’interface de votre choix qui exécutent des actions. Pour une présentation des commandes de complément, consultez [les commandes de complément pour Excel, PowerPoint et Word](../design/add-in-commands.md).

Cet article explique comment modifier votre manifeste pour définir des commandes de complément et comment créer le code pour [les commandes de fonction](../design/add-in-commands.md#types-of-add-in-commands). Le schéma suivant illustre la hiérarchie des éléments utilisés pour définir des commandes de complément. Ces éléments sont décrits plus en détail dans cet article.

![Vue d’ensemble des éléments de commandes de complément dans le manifeste. Le nœud supérieur ici est VersionOverrides avec les hôtes enfants et les ressources. Sous Hosts se trouvent Host, puis DesktopFormFactor. Sous DesktopFormFactor se trouvent FunctionFile et ExtensionPoint. Sous ExtensionPoint, vous pouvez trouver CustomTab ou OfficeTab et le menu Office. Sous CustomTab ou l’onglet Office, il y a Groupe, puis Contrôle, puis Action. Sous Menu Office, contrôle puis action. Sous Ressources (enfant de VersionOverrides) se trouvent images, URL, ShortStrings et LongStrings.](../images/version-overrides.png)

## <a name="step-1-create-the-project"></a>Étape 1 : Créer le projet

Nous vous recommandons de créer un projet en suivant l’un des démarrages rapides tels que [Créer un complément du volet Office Excel](../quickstarts/excel-quickstart-jquery.md). Chaque démarrage rapide pour Excel, Word et PowerPoint génère un projet qui contient déjà une commande de complément (bouton) pour afficher le volet Office. Vérifiez que vous avez lu [des commandes de complément pour Excel, Word et PowerPoint](../design/add-in-commands.md) avant d’utiliser des commandes de complément.

## <a name="step-2-create-a-task-pane-add-in"></a>Étape 2 : créer un complément de volet Office

Pour commencer à utiliser des commandes de complément, vous devez d’abord créer un complément du volet Office, puis modifier le manifeste du complément comme décrit dans cet article. Vous ne pouvez pas utiliser de commandes de complément avec des compléments de contenu. Si vous mettez à jour un manifeste existant, vous devez ajouter les **espaces de noms XML appropriés** et ajouter l’élément **\<VersionOverrides\>** au manifeste comme décrit à l’étape [3 : Ajouter l’élément VersionOverrides](#step-3-add-versionoverrides-element).

L’exemple suivant illustre le manifeste d’un complément Office 2013. Il n’existe aucune commande de complément dans ce manifeste, car il n’y a pas **\<VersionOverrides\>** d’élément. Office 2013 ne prend pas en charge les commandes de complément, mais en ajoutant **\<VersionOverrides\>** à ce manifeste, votre complément s’exécutera dans Office 2013 et Office 2016. Dans Office 2013, votre complément n’affiche pas les commandes de complément et utilise la valeur de **\<SourceLocation\>** celle-ci pour exécuter votre complément en tant que complément de volet Office unique. Dans Office 2016, si aucun élément n’est **\<VersionOverrides\>** inclus, le volet Office de votre complément s’ouvre automatiquement sur l’URL spécifiée dans **\<SourceLocation\>**. Si vous incluez **\<VersionOverrides\>**, toutefois, votre complément affiche uniquement les commandes de complément et n’affiche pas initialement le volet Office de votre complément.
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="https://www.contoso.com/Images/Icon_32.png" />
  <SupportUrl DefaultValue="https://www.contoso.com/contact" />
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

L’élément **\<VersionOverrides\>** est l’élément racine qui contient la définition de votre commande de complément. **\<VersionOverrides\>** est un élément enfant de l’élément **\<OfficeApp\>** dans le manifeste. Le tableau suivant répertorie les attributs de l’élément **\<VersionOverrides\>** .

|Attribut|Description|
|:-----|:-----|
|**xmlns** <br/> | Obligatoire. Emplacement du schéma, qui doit être `http://schemas.microsoft.com/office/taskpaneappversionoverrides`. <br/> |
|**xsi:type** <br/> |Obligatoire. Version du schéma. La version décrite dans cet article est « VersionOverridesV1_0 ».  <br/> |

Le tableau suivant identifie les éléments enfants de **\<VersionOverrides\>**.
  
|Élément|Description|
|:-----|:-----|
|**\<Description\>** <br/> |Facultatif. Décrit le complément. Cet élément enfant **\<Description\>** remplace un élément précédent **\<Description\>** dans la partie parente du manifeste. L’attribut **de résident** de cet **\<Description\>** élément est défini sur l’ID d’un  **\<String\>** élément. L’élément **\<String\>** contient le texte pour **\<Description\>**. <br/> |
|**\<Requirements\>** <br/> |Facultatif. Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cet élément enfant **\<Requirements\>** remplace l’élément **\<Requirements\>** dans la partie parente du manifeste. Pour plus d’informations, consultez [Spécifier les applications Office et les exigences de l’API](../develop/specify-office-hosts-and-api-requirements.md).  <br/> |
|**\<Hosts\>** <br/> |Obligatoire. Spécifie une collection d’applications Office. L’élément enfant **\<Hosts\>** remplace l’élément **\<Hosts\>** dans la partie parente du manifeste. Vous devez inclure un ensemble d’attributs **xsi:type** à « Classeur » ou « Document ». <br/> |
|**\<Resources\>** <br/> |Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste. Par exemple, la valeur de l’élément **\<Description\>** fait référence à un élément enfant dans **\<Resources\>**. L’élément **\<Resources\>** est décrit à [l’étape 7 : Ajouter l’élément Resources](#step-7-add-the-resources-element) plus loin dans cet article. <br/> |

L’exemple suivant montre comment utiliser l’élément **\<VersionOverrides\>** et ses éléments enfants.

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

L’élément **\<Hosts\>** contient un ou plusieurs **\<Host\>** éléments. Un **\<Host\>** élément spécifie une application Office particulière. L’élément **\<Host\>** contient des éléments enfants qui spécifient les commandes de complément à afficher après l’installation de votre complément dans cette application Office. Pour afficher les mêmes commandes de complément dans deux ou plusieurs applications Office différentes, vous devez dupliquer les éléments enfants dans chacune **\<Host\>** d’elles.

L’élément **\<DesktopFormFactor\>** spécifie les paramètres d’un complément qui s’exécute dans Office sur le Web (dans un navigateur) et Windows.

Voici un exemple de , **\<Host\>** et **\<DesktopFormFactor\>** d’éléments **\<Hosts\>**.

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

L’élément **\<FunctionFile\>** spécifie un fichier qui contient du code JavaScript à exécuter lorsqu’une commande de complément utilise l’action **ExecuteFunction** (consultez [les contrôles Button](/javascript/api/manifest/control-button) pour obtenir une description). L’attribut **\<FunctionFile\>** **de résident** de l’élément est défini sur un fichier HTML qui inclut tous les fichiers JavaScript dont vos commandes de complément ont besoin. Vous ne pouvez pas créer une liaison directe vers un fichier JavaScript. Vous pouvez uniquement créer une liaison vers un fichier HTML. Le nom de fichier est spécifié en tant qu’élément **\<Url\>** dans l’élément **\<Resources\>** .

Voici un exemple de l’élément **\<FunctionFile\>** .
  
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

Le Code JavaScript dans le fichier HTML référencé par l’élément **\<FunctionFile\>** doit appeler `Office.initialize`. L’élément **\<FunctionName\>** (voir [contrôles Button](/javascript/api/manifest/control-button) pour obtenir une description) utilise les fonctions dans **\<FunctionFile\>**.

Le code suivant montre comment implémenter la fonction utilisée par **\<FunctionName\>**.

```js
<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here.
        };
    })();

    // Define the function.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("Function command works. Button ID=" + event.source.id,
            function (asyncResult) {
                const error = asyncResult.error;
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
    
    // You must register the function with the following line.
    Office.actions.associate("writeText", writeText);
</script>
```

> [!IMPORTANT]
> L’appel de l’élément **event.completed** indique que vous avez correctement géré l’événement. Lorsqu’une fonction est appelée plusieurs fois (par exemple, lorsque l’utilisateur clique plusieurs fois sur une même commande de complément), tous les événements sont automatiquement mis en file d’attente. Le premier événement s’exécute automatiquement, tandis que les autres événements restent dans la file d’attente. Lorsque votre fonction appelle **event.completed**, l’appel de la file d’attente suivant de cette fonction s’exécute. Vous devez implémenter **event.completed** pour que votre fonction s’exécute correctement.

## <a name="step-6-add-extensionpoint-elements"></a>Etape 6 : ajouter des éléments ExtensionPoint

L’élément **\<ExtensionPoint\>** définit où les commandes de complément doivent apparaître dans l’interface utilisateur d’Office. Vous pouvez définir des **\<ExtensionPoint\>** éléments avec ces valeurs **xsi:type** .

- **PrimaryCommandSurface**, qui fait référence au ruban dans Office.

- **ContextMenu**, qui est le menu contextuel qui apparaît lorsque vous cliquez avec le bouton droit de la souris dans l’interface utilisateur Office.

Les exemples suivants montrent comment utiliser l’élément **\<ExtensionPoint\>** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu** , ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.

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

|Élément|Description|
|:-----|:-----|
|**\<CustomTab\>** <br/> |Obligatoire si vous souhaitez ajouter un onglet personnalisé au ruban (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **\<CustomTab\>** , vous ne pouvez pas l’utiliser **\<OfficeTab\>** . L’attribut  **id** est requis. <br/> |
|**\<OfficeTab\>** <br/> |Obligatoire si vous souhaitez étendre un onglet de ruban d’application Office par défaut (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **\<OfficeTab\>** , vous ne pouvez pas l’utiliser **\<CustomTab\>** . <br/> Pour plus de valeurs d’onglet à utiliser avec l’attribut **d’ID** , consultez [les valeurs Tab pour les onglets du ruban d’application Office par défaut](/javascript/api/manifest/officetab).  <br/> |
|**\<OfficeMenu\>** <br/> | Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : <br/> **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. <br/> **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul. <br/> |
|**\<Group\>** <br/> |Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est obligatoire. Il s’agit d’une chaîne avec un maximum de 125 caractères. <br/> |
|**\<Label\>** <br/> |Obligatoire. Libellé du groupe. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<ShortStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> |
|**\<Icon\>** <br/> |Obligatoire. Indique l’icône du groupe qui doit être utilisée sur les périphériques de petit facteur de forme ou lorsque les boutons sont affichés en trop grand nombre. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<Image\>** élément. L’élément **\<Image\>** est un élément enfant de l’élément **\<Images\>** , qui est un élément enfant de l’élément **\<Resources\>** . L’attribut **size** donne la taille, en pixels, de l’image. Trois tailles d’image, en pixels, sont obligatoires : 16, 32 et 80. Cinq tailles facultatives, en pixels, sont également prises en charge : 20, 24, 40, 48 et 64. <br/> |
|**\<Tooltip\>** <br/> |Facultatif. Info-bulle du groupe. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<LongStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> |
|**\<Control\>** <br/> |Chaque groupe requiert au moins un contrôle. Un **\<Control\>** élément peut être un **bouton** ou un **menu**. Utilisez **Menu** pour spécifier une liste déroulante de contrôles de bouton. Actuellement, seuls les boutons et les menus sont pris en charge. Pour plus d’informations, consultez [Contrôles bouton](/javascript/api/manifest/control-button) et [contrôles de menu](/javascript/api/manifest/control-menu) . <br/>**Note:** Pour faciliter la résolution des problèmes, nous vous recommandons d’ajouter un **\<Control\>** élément et les éléments enfants associés **\<Resources\>** un par un.          |

### <a name="button-controls"></a>Contrôles de bouton

Un bouton effectue une seule action lorsque l’utilisateur le sélectionne. Il peut exécuter une fonction JavaScript ou afficher un volet de tâches. L’exemple de code suivant montre comment définir deux boutons. Le premier exécute une fonction JavaScript sans afficher d’interface utilisateur et le second affiche un volet de tâches. Dans l’élément **\<Control\>** :

- l’attribut **type** est obligatoire et doit être défini sur **Button**.

- **L’attribut id** de l’élément **\<Control\>** est une chaîne avec un maximum de 125 caractères.

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

|Éléments|Description|
|:-----|:-----|
|**\<Label\>** <br/> |Obligatoire. Texte du bouton. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<ShortStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> |
|**\<Tooltip\>** <br/> |Facultatif. Info-bulle pour le bouton. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<LongStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> |
|**\<Supertip\>** <br/> | Obligatoire. Info-bulle multiligne associée à ce bouton, qui est définie de la façon suivante : <br/> **Titre** <br/>  Obligatoire. Texte de la propriété SuperTip. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<ShortStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> **\<Description\>** <br/>  Obligatoire. Description de l’info-bulle. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<LongStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> |
|**\<Icon\>** <br/> | Obligatoire. Contient les **\<Image\>** éléments du bouton. Les fichiers image doivent être au format .png. <br/> **\<Image\>** <br/>  Définit une image à afficher sur le bouton. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<Image\>** élément. L’élément **\<Image\>** est un élément enfant de l’élément **\<Images\>** , qui est un élément enfant de l’élément **\<Resources\>** . L’attribut **size** indique la taille, en pixels, de l’image. Trois tailles d’image, en pixels, sont obligatoires : 16, 32 et 80. Cinq tailles facultatives, en pixels, sont également prises en charge : 20, 24, 40, 48 et 64. <br/> |
|**\<Action\>** <br/> | Obligatoire. Indique l’action à réaliser lorsque l’utilisateur sélectionne le bouton. Vous pouvez spécifier une des valeurs suivantes pour l’attribut **xsi:type** : <br/> **ExecuteFunction**, qui exécute une fonction JavaScript située dans le fichier référencé par **\<FunctionFile\>**. L’élément **\<FunctionName\>** enfant spécifie le nom de la fonction à exécuter. <br/> **ShowTaskPane**, qui affiche le volet Office du complément. L’élément **\<SourceLocation\>** enfant spécifie l’emplacement du fichier source de la page à afficher. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<Url\>** élément dans l’élément **\<Urls\>** de l’élément **\<Resources\>** . <br/> |

### <a name="menu-controls"></a>Contrôles de menu

Un contrôle de type **Menu** peut être utilisé avec **PrimaryCommandSurface** ou **ContextMenu**, et permet de définir :
  
- une option de menu de niveau racine.
- une liste de sous-menus.

Lorsqu’il est utilisé avec **PrimaryCommandSurface**, l’option de menu de niveau racine s’affiche sous la forme d’un bouton dans le ruban. Lorsque le bouton est sélectionné, le sous-menu s’affiche sous la forme d’une liste déroulante. Lorsqu’il est utilisé avec **ContextMenu**, un élément de menu avec un sous-menu est inséré dans le menu contextuel. Dans les deux cas, les éléments individuels du sous-menu peuvent exécuter une fonction JavaScript ou afficher un volet de tâches. Un seul niveau de sous-menus est pris en charge pour l’instant.

L’exemple suivant montre comment définir un élément de menu avec deux éléments de sous-menu. Le premier élément de sous-menu affiche un volet Office et le deuxième élément de sous-menu exécute une fonction JavaScript. Dans l’élément **\<Control\>** :

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

|Éléments|Description|
|:-----|:-----|
|**\<Label\>** <br/> |Obligatoire. Texte de l’élément de menu racine. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<ShortStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> |
|**\<Tooltip\>** <br/> |Facultatif. Info-bulle du menu. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<LongStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> |
|**\<SuperTip\>** <br/> | Obligatoire. Info-bulle multiligne associée au menu, qui est définie de la façon suivante : <br/> **\<Title\>** <br/>  Obligatoire. Texte de l’info-bulle améliorée. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<ShortStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> **\<Description\>** <br/>  Obligatoire. Description de l’info-bulle. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<String\>** élément. L’élément **\<String\>** est un élément enfant de l’élément **\<LongStrings\>** , qui est un élément enfant de l’élément **\<Resources\>** . <br/> |
|**\<Icon\>** <br/> | Obligatoire. Contient les **\<Image\>** éléments du menu. Les fichiers image doivent être au format .png. <br/> **\<Image\>** <br/>  Image du menu. L’attribut **de résident** doit être défini sur la valeur de l’attribut **id** d’un **\<Image\>** élément. L’élément **\<Image\>** est un élément enfant de l’élément **\<Images\>** , qui est un élément enfant de l’élément **\<Resources\>** . L’attribut **size** indique la taille, en pixels, de l’image. Trois tailles d’image, en pixels, sont nécessaires : 16, 32 et 80. 5 tailles facultatives, en pixels, sont également prises en charge : 20, 24, 40, 48 et 64. <br/> |
|**\<Items\>** <br/> |Obligatoire. Contient les **\<Item\>** éléments de chaque élément de sous-menu. Chaque **\<Item\>** élément contient les mêmes éléments enfants que [les contrôles Button](/javascript/api/manifest/control-button).  <br/> |

## <a name="step-7-add-the-resources-element"></a>Étape 7 : ajouter l’élément Resources

L’élément **\<Resources\>** contient des ressources utilisées par les différents éléments enfants de l’élément **\<VersionOverrides\>** . Les ressources incluent des icônes, des chaînes et des URL. Un élément du manifeste peut utiliser une ressource en référençant l’**id** de la ressource. L’utilisation de l’**id** permet d’organiser le manifeste, en particulier lorsqu’il existe des versions différentes de la ressource pour différents paramètres régionaux. Un **id** doit comporter 32 caractères au maximum.
  
L’exemple suivant montre comment utiliser l’élément **\<Resources\>** . Chaque ressource peut avoir un ou plusieurs **\<Override\>** éléments enfants pour définir une ressource différente pour des paramètres régionaux spécifiques.

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

|Resource|Description|
|:-----|:-----|
|**\<Images\>**/ **\<Image\>** <br/> | Fournit l’URL HTTPS d’un fichier image. Chaque image doit définir les trois tailles d’image obligatoires : <br/>  16 x 16 <br/>  32 x 32 <br/>  80 × 80 <br/>  Les tailles d’image suivantes sont également prises en charge, mais ne sont pas obligatoires : <br/>  20 × 20 <br/>  24 × 24 <br/>  40 × 40 <br/>  48 × 48 <br/>  64 x 64 <br/> |
|**\<Urls\>**/ **\<Url\>** <br/> |Indique un emplacement d’URL HTTPS. Une URL peut comporter 2 048 caractères au maximum.  <br/> |
|**\<ShortStrings\>**/ **\<String\>** <br/> |Texte pour **\<Label\>** et **\<Title\>** éléments. Chacun **\<String\>** contient un maximum de 125 caractères. <br/> |
|**\<LongStrings\>**/ **\<String\>** <br/> |Texte pour **\<Tooltip\>** et **\<Description\>** éléments. Chacun **\<String\>** contient un maximum de 250 caractères. <br/> |

> [!NOTE]
> Vous devez utiliser SSL (Secure Sockets Layer) pour toutes les URL des éléments et **\<Url\>** des **\<Image\>** URL.

### <a name="tab-values-for-default-office-app-ribbon-tabs"></a>Valeurs d’onglet pour les onglets du ruban de l’application Office par défaut

Dans Excel et Word, vous pouvez ajouter vos commandes de complément au ruban en utilisant les onglets de l’interface utilisateur Office par défaut. Le tableau suivant répertorie les valeurs que vous pouvez utiliser pour l’attribut **id** de l’élément **\<OfficeTab\>** . Les valeurs des onglets respectent la casse.

|Application cliente Office|Valeurs des onglets|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |

## <a name="see-also"></a>Voir aussi

- [Commandes de complément pour Excel, PowerPoint et Word](../design/add-in-commands.md)
- [Exemple : Créer un complément Excel avec des boutons de commande](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/excel)
- [Exemple : Créer un complément Word avec des boutons de commande](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/word)
- [Exemple : Créer un complément PowerPoint avec des boutons de commande](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/powerpoint)
