---
title: Manifeste XML des compléments Office
description: Obtenez une vue d’ensemble du manifeste de Complément Office et de ses applications.
ms.date: 03/18/2020
localization_priority: Priority
ms.openlocfilehash: 4d2fa054cc268b68eb1c05ba82f9cd7745bc8685
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093748"
---
# <a name="office-add-ins-xml-manifest"></a>Manifeste XML des compléments Office

Le fichier manifeste XML d’un complément Office la manière dont votre complément doit être activé lorsqu’un utilisateur final l’installe et l’utilise avec des documents et des applications Office.

Un fichier de manifeste XML basé sur ce schéma permet à un Complément Office d’effectuer les opérations suivantes :

* Se décrire en fournissant un ID, une version, une description, un nom d’affichage et un paramètre régional par défaut.

* Précisez les images utilisées pour l'image de marque du complément et l'iconographie utilisée pour [commandes complémentaires][] dans le ruban d'application de l'Office.

* Spécifier comment le complément s’intègre à Office, y compris les interfaces utilisateur personnalisées, telles que les boutons du ruban créés par le complément.

* Spécifier les dimensions par défaut demandées pour des compléments de contenu, et la hauteur demandée pour des compléments Outlook.

* Déclarer les autorisations que le Complément Office nécessite, par exemple la lecture du document ou l’écriture dans celui-ci.

* Pour des compléments Outlook, définir la ou les règles qui spécifient le contexte dans lequel ils seront activés et seront en interaction avec un message, un rendez-vous ou un élément de demande de réunion.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="required-elements"></a>Éléments requis

Le tableau suivant spécifie les éléments qui sont requis pour les trois types de compléments Office.

> [!NOTE]
> Il existe également un ordre obligatoire d’apparition des éléments au sein de leur élément parent. Pour plus d’informations, reportez-vous à la rubrique [Comment trouver l’ordre approprié d’éléments manifeste](manifest-element-ordering.md).


### <a name="required-elements-by-office-add-in-type"></a>Éléments requis par type de complément Office

| Élément                                                                                      | Contenu | Volet de tâches | Outlook |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| [OfficeApp][]                                                                                |    X    |     X     |    X    |
| 
  [Id][]                                                                                       |    X    |     X     |    X    |
| [Version][]                                                                                  |    X    |     X     |    X    |
| [ProviderName][]                                                                             |    X    |     X     |    X    |
| [DefaultLocale][]                                                                            |    X    |     X     |    X    |
| [DisplayName][]                                                                              |    X    |     X     |    X    |
| [Description][]                                                                              |    X    |     X     |    X    |
| [IconUrl][]                                                                                  |    X    |     X     |    X    |
| [SupportUrl][]\*\*                                                                           |    X    |     X     |    X    |
| [DefaultSettings (ContentApp)][]<br/>[DefaultSettings (TaskPaneApp)][]                       |    X    |     X     |         |
| [SourceLocation (ContentApp)][]<br/>[SourceLocation (TaskPaneApp)][]                         |    X    |     X     |         |
| [DesktopSettings][]                                                                          |         |           |    X    |
| [SourceLocation (MailApp)][]                                                                 |         |           |    X    |
| 
  [Permissions (ContentApp)][]<br/>
  [Permissions (TaskPaneApp)][]<br/>
  [Permissions (MailApp)][] |    X    |     X     |    X    |
| 
  [Rule (RuleCollection)][]<br/>
  [Rule (MailApp)][]                                             |         |           |    X    |
| [Requirements (MailApp)][]                                                                  |         |           |    X    |
| [Set*][]<br/>[Sets (MailAppRequirements)*][]                                                 |         |           |    X    |
| [Form*][]<br/>[FormSettings*][]                                                              |         |           |    X    |
| [Sets (Requirements)*][]                                                                     |    X    |     X     |         |
| [Hosts*][]                                                                                   |    X    |     X     |         |

_\*Ajouté dans le schéma de manifeste du complément Office version 1.1._

_\*\* SupportUrl n’est obligatoire que pour les compléments distribués via AppSource._

<!-- Links for above table -->

[officeapp]: ../reference/manifest/officeapp.md
[id]: ../reference/manifest/id.md
[version]: ../reference/manifest/version.md
[providername]: ../reference/manifest/providername.md
[defaultlocale]: ../reference/manifest/defaultlocale.md
[displayname]: ../reference/manifest/displayname.md
[description]: ../reference/manifest/description.md
[iconurl]: ../reference/manifest/iconurl.md
[supporturl]: ../reference/manifest/supporturl.md
[defaultsettings (contentapp)]: ../reference/manifest/defaultsettings.md
[defaultsettings (taskpaneapp)]: ../reference/manifest/defaultsettings.md
[sourcelocation (contentapp)]: ../reference/manifest/sourcelocation.md
[sourcelocation (taskpaneapp)]: ../reference/manifest/sourcelocation.md
[desktopsettings]: /previous-versions/office/fp179684%28v=office.15%29
[sourcelocation (mailapp)]: /previous-versions/office/fp123668%28v=office.15%29
[permissions (contentapp)]: ../reference/manifest/permissions.md
[permissions (taskpaneapp)]: ../reference/manifest/permissions.md
[permissions (mailapp)]: ../reference/manifest/permissions.md
[rule (rulecollection)]: ../reference/manifest/rule.md
[rule (mailapp)]: ../reference/manifest/rule.md
[requirements (mailapp)*]: ../reference/manifest/requirements.md
[set*]: ../reference/manifest/set.md
[sets (mailapprequirements)*]: ../reference/manifest/sets.md
[form*]: ../reference/manifest/form.md
[formsettings*]: ../reference/manifest/formsettings.md
[sets (requirements)*]: ../reference/manifest/sets.md
[hosts*]: ../reference/manifest/hosts.md

## <a name="hosting-requirements"></a>Configuration requise pour l’hébergement

Tous les URI des images, tels que ceux utilisés pour les [commandes de complément][], doivent prendre en charge la mise en cache. Le serveur qui héberge l’image ne doit pas renvoyer d’en-tête `Cache-Control` spécifiant `no-cache`, `no-store` ou des options similaires dans la réponse HTTP.

Toutes les URL, telles que les emplacements des fichiers source spécifiés dans l’élément [SourceLocation](../reference/manifest/sourcelocation.md), doivent être **sécurisées par une protection SSL (HTTPS)**. [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a>Bonnes pratiques pour l’envoi dans AppSource

Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.

Les compléments envoyés à AppSource doivent également inclure l’élément [SupportUrl](../reference/manifest/supporturl.md). Pour plus d’informations, reportez-vous à [Stratégies de validation pour les applications et les compléments envoyés à AppSource](/legal/marketplace/certification-policies).

Utilisez uniquement l’élément [AppDomains](../reference/manifest/appdomains.md) pour spécifier des domaines différents de celui spécifié dans l’élément [SourceLocation](../reference/manifest/sourcelocation.md) pour les scénarios d’authentification.

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a>Spécifier les domaines que vous souhaitez ouvrir dans la fenêtre de complément

Quand vous exécutez Office sur le web, votre volet Office peut accéder à n’importe quelle URL. Cependant, sur les plateformes de bureau, si votre complément tente d’accéder à une URL située dans un autre domaine que celui qui héberge la page de démarrage (comme indiqué dans l’élément [SourceLocation](../reference/manifest/sourcelocation.md) du fichier manifeste), cette URL s’ouvre dans une nouvelle fenêtre de navigateur en dehors du volet de complément de l’application hôte Office.

Pour remplacer ce comportement (version de bureau d’Office), spécifiez chaque domaine à ouvrir dans la fenêtre de complément dans la liste des domaines spécifiés dans l’élément [AppDomains](../reference/manifest/appdomains.md) du fichier manifeste. Si le complément tente d’accéder à une URL située dans un domaine figurant dans cette liste, il s’ouvre dans le volet Office d’Office sur le web et de la version de bureau d’Office. S’il tente d’accéder à une URL qui ne figure pas dans la liste, dans la version de bureau d’Office, cette URL s’ouvre dans une nouvelle fenêtre de navigateur (en dehors du volet de complément).

> [!NOTE]
> Il existe deux exceptions à ce comportement :
>
> - Il s’applique uniquement au volet racine du complément. S’il existe un iframe incorporé dans la page de complément, l’iframe peut être dirigé vers n’importe quelle URL, qu’elle figure dans la liste des **AppDomains** ou non, y compris dans la version de bureau d’Office.
> - Lorsqu’une boîte de dialogue est ouverte avec l’API [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-), l’URL transmise à la méthode doit se trouver dans le même domaine que le complément, mais la boîte de dialogue peut ensuite être redirigée vers n’importe quelle URL, même si elle est répertoriée dans **AppDomains**, y compris dans la version de bureau d’Office.

L’exemple de manifeste XML suivant héberge sa page de complément principale dans le domaine `https://www.contoso.com` comme indiqué dans l’élément **SourceLocation**. Il indique également le domaine `https://www.northwindtraders.com` dans un élément [AppDomain](../reference/manifest/appdomain.md) au sein de la liste d’éléments **AppDomains**. Si le complément ouvre une page dans le domaine `www.northwindtraders.com`, cette page s’ouvre dans le volet de complément, y compris dans le bureau Office.

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a>Spécifier les domaines à partir desquels les appels d’API Office.js sont effectués

Votre complément peut effectuer des appels d’API Office.js à partir du domaine référencé dans l’élément [SourceLocation](../reference/manifest/sourcelocation.md) du fichier manifeste. Si votre complément comporte d’autres IFrames qui nécessitent un accès aux API Office.js, ajoutez le domaine de cette URL source à la liste spécifiée dans l’élément [AppDomains](../reference/manifest/appdomains.md) du fichier manifeste. Si un IFrame associé à une source qui ne figure pas dans la liste `AppDomains` tente d’effectuer un appel d’API Office.js, le complément reçoit une [erreur d’autorisation refusée](../reference/javascript-api-for-office-error-codes.md).

## <a name="manifest-v11-xml-file-examples-and-schemas"></a>Exemples et schémas de fichier XML manifeste version 1.1

Les sections suivantes présentent des exemples de fichiers manifeste XML version 1.1 pour des compléments de contenu, de volet Office et Outlook.

# <a name="task-pane"></a>[Volet Office](#tab/tabid-1)

[Schémas de manifeste de compléments](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office app ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
                    <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane1.Url" />
                      </Action>
                    </Item>

                    <Item id="Contoso.Menu.Item2">
                      <Label resid="Contoso.Item2.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item2.Label" />
                        <Description resid="Contoso.Item2.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="content"></a>[Content](#tab/tabid-2)

[Schémas de manifeste de compléments](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mail"></a>[Application de messagerie](#tab/tabid-3)

[Schémas de manifeste de compléments](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-an-office-add-ins-manifest"></a>Valider un manifeste de complément Office

Pour plus d’informations sur la validation d’un manifeste par rapport à la [XSD (XML Schema Definition)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), voir [Valider le manifeste d’un complément Office](../testing/troubleshoot-manifest.md).

## <a name="see-also"></a>Voir aussi

* [Comment trouver l’ordre approprié d’éléments manifeste](manifest-element-ordering.md)
* [Création de commandes de complément dans votre manifeste][commandes de complément]
* [Spécification des exigences en matière d’hôtes Office et d’API](specify-office-hosts-and-api-requirements.md)
* [Localisation des compléments Office](localization.md)
* [Référence de schéma pour les manifestes des compléments Office](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
* [Mettre à jour la version du manifeste et de l’API](update-your-javascript-api-for-office-and-manifest-schema-version.md)
* [Identifier un complément COM équivalent](make-office-add-in-compatible-with-existing-com-add-in.md)
* [Demande d’autorisations d’utilisation de l’API dans des compléments](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
* [Valider le manifeste d’un complément Office](../testing/troubleshoot-manifest.md)

[commandes de complément]: create-addin-commands.md
