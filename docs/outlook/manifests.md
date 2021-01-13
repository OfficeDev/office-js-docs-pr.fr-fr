---
title: Manifestes des compléments Outlook
description: Le manifeste décrit l’intégration d’un complément Outlook avec les clients Outlook et comprend un exemple.
ms.date: 05/27/2020
localization_priority: Priority
ms.openlocfilehash: 0135db8b6ff2b9fbcb3b6370979d8013aa21155a
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839823"
---
# <a name="outlook-add-in-manifests"></a><span data-ttu-id="a9d93-103">Manifestes des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="a9d93-103">Outlook add-in manifests</span></span>

<span data-ttu-id="a9d93-p101">Un complément Outlook contient deux composants : le manifeste du complément XML et une page web, pris en charge par la bibliothèque JavaScript pour les compléments Office (office.js). Le manifeste décrit l’intégration du complément avec les clients Outlook. Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p101">An Outlook add-in consists of two components: the XML add-in manifest and a web page supported by the JavaScript library for Office Add-ins (office.js). The manifest describes how the add-in integrates across Outlook clients. The following is an example.</span></span>

 > [!NOTE]
 > <span data-ttu-id="a9d93-p102">Dans l’exemple suivant, toutes les valeurs d’URL commencent par «https://appdemo.contoso.com». Cette valeur est un espace réservé. Dans un manifeste valide réel, ces valeurs contiendraient des URL web HTTPS valides.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p102">All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8" ?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="https://appdemo.contoso.com/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="https://appdemo.contoso.com/images/blue-128.png" />
  <SupportUrl DefaultValue="https://appdemo.contoso.com"/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read task pane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="functionFile" />

          <!-- Message read form -->
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadDemoGroup">
                <Label resid="groupLabel" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="msgReadFunctionButton">
                  <Label resid="funcReadButtonLabel" />
                  <Supertip>
                    <Title resid="funcReadSuperTipTitle" />
                    <Description resid="funcReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getSubject</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="msgReadMenuButton">
                  <Label resid="menuReadButtonLabel" />
                  <Supertip>
                    <Title resid="menuReadSuperTipTitle" />
                    <Description resid="menuReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="msgReadMenuItem1">
                      <Label resid="menuItem1ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem1ReadLabel" />
                        <Description resid="menuItem1ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemClass</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem2">
                      <Label resid="menuItem2ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem2ReadLabel" />
                        <Description resid="menuItem2ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getDateTimeCreated</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem3">
                      <Label resid="menuItem3ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem3ReadLabel" />
                        <Description resid="menuItem3ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemID</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="paneReadButtonLabel" />
                  <Supertip>
                    <Title resid="paneReadSuperTipTitle" />
                    <Description resid="paneReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="readTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <!-- Blue icon -->
        <bt:Image id="blue-icon-16" DefaultValue="https://appdemo.contoso.com/images/blue-16.png" />
        <bt:Image id="blue-icon-32" DefaultValue="https://appdemo.contoso.com/images/blue-32.png" />
        <bt:Image id="blue-icon-80" DefaultValue="https://appdemo.contoso.com/images/blue-80.png" />
        <!-- Red icon -->
        <bt:Image id="red-icon-16" DefaultValue="https://appdemo.contoso.com/images/red-16.png" />
        <bt:Image id="red-icon-32" DefaultValue="https://appdemo.contoso.com/images/red-32.png" />
        <bt:Image id="red-icon-80" DefaultValue="https://appdemo.contoso.com/images/red-80.png" />
        <!-- Green icon -->
        <bt:Image id="green-icon-16" DefaultValue="https://appdemo.contoso.com/images/green-16.png" />
        <bt:Image id="green-icon-32" DefaultValue="https://appdemo.contoso.com/images/green-32.png" />
        <bt:Image id="green-icon-80" DefaultValue="https://appdemo.contoso.com/images/green-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="https://appdemo.contoso.com/FunctionFile/Functions.html" />
        <bt:Url id="readTaskPaneUrl" DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Add-in Demo" />
        <bt:String id="funcReadButtonLabel" DefaultValue="Get subject" />
        <bt:String id="menuReadButtonLabel" DefaultValue="Get property" />
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties" />

        <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment" />
        <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get" />
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties" />

        <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class" />
        <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created" />
        <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button." />
        <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button." />
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane." />

        <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
        <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
        <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

## <a name="schema-versions"></a><span data-ttu-id="a9d93-110">Versions de schéma</span><span class="sxs-lookup"><span data-stu-id="a9d93-110">Schema versions</span></span>

<span data-ttu-id="a9d93-p103">Tous les clients Outlook ne prennent pas en charge les fonctionnalités les plus récentes, et certains utilisateurs Outlook disposeront d’une version antérieure d’Outlook. Le fait de disposer de versions de schéma permet aux développeurs de créer des compléments à compatibilité descendante, en utilisant les fonctionnalités les plus récentes lorsqu’elles sont disponibles mais qui fonctionnent toujours sur les versions antérieures.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p103">Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.</span></span>

<span data-ttu-id="a9d93-p104">L’élément **VersionOverrides** dans le manifeste en est un exemple. Tous les éléments définis dans **VersionOverrides** remplaceront le même élément dans l’autre partie du manifeste. Cela signifie que, dès que possible, Outlook utilisera les éléments de la section **VersionOverrides** pour configurer le complément. Toutefois, si la version d’Outlook ne prend pas en charge une version de **VersionOverrides**, Outlook l’ignorera et se référera aux informations contenues dans le reste du manifeste.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p104">The **VersionOverrides** element in the manifest is an example of this. All elements defined inside **VersionOverrides** will override the same element in the other part of the manifest. This means that, whenever possible, Outlook will use what is in the **VersionOverrides** section to set up the add-in. However, if the version of Outlook doesn't support a certain version of **VersionOverrides**, Outlook will ignore it and depend on the information in the rest of the manifest.</span></span> 

<span data-ttu-id="a9d93-117">Cette approche signifie que les développeurs ne doivent pas créer plusieurs manifestes individuels, mais plutôt conserver tous les éléments définis dans un fichier.</span><span class="sxs-lookup"><span data-stu-id="a9d93-117">This approach means that developers don't have to create multiple individual manifests, but rather keep everything defined in one file.</span></span>

<span data-ttu-id="a9d93-118">Les versions actuelles du schéma sont les suivantes :</span><span class="sxs-lookup"><span data-stu-id="a9d93-118">The current versions of the schema are:</span></span>


|<span data-ttu-id="a9d93-119">Version</span><span class="sxs-lookup"><span data-stu-id="a9d93-119">Version</span></span>|<span data-ttu-id="a9d93-120">Description</span><span class="sxs-lookup"><span data-stu-id="a9d93-120">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="a9d93-121">v1.0</span><span class="sxs-lookup"><span data-stu-id="a9d93-121">v1.0</span></span>|<span data-ttu-id="a9d93-p105">Prend en charge la version 1.0 de l’API Office JavaScript. Pour les compléments Outlook, la prise en charge des formulaires de lecture est également incluse.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p105">Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form.</span></span> |
|<span data-ttu-id="a9d93-124">v1.1</span><span class="sxs-lookup"><span data-stu-id="a9d93-124">v1.1</span></span>|<span data-ttu-id="a9d93-p106">Prend en charge la version 1.1 de l’interface API Office JavaScript et **VersionOverrides**. Pour les compléments Outlook, la prise en charge des formulaires de composition est incluse.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p106">Supports version 1.1 of the Office JavaScript API and **VersionOverrides**. For Outlook add-ins, this adds support for compose form.</span></span>|
|<span data-ttu-id="a9d93-127">**VersionOverrides** 1.0</span><span class="sxs-lookup"><span data-stu-id="a9d93-127">**VersionOverrides** 1.0</span></span>|<span data-ttu-id="a9d93-p107">Prend en charge les versions ultérieures de l’API Office JavaScript. La prise en charge des commandes de complément est incluse.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p107">Supports later versions of the Office JavaScript API. This supports add-in commands.</span></span>|
|<span data-ttu-id="a9d93-130">**VersionOverrides** 1.1</span><span class="sxs-lookup"><span data-stu-id="a9d93-130">**VersionOverrides** 1.1</span></span>|<span data-ttu-id="a9d93-p108">Prend en charge les versions ultérieures de l’interface API Office JavaScript. Les commandes de complément sont prises en charge, ainsi que de nouvelles fonctionnalités, telles que les [volets Office à épingler](pinnable-taskpane.md) et les compléments mobiles.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p108">Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.</span></span>|

<span data-ttu-id="a9d93-p109">Cet article porte sur les conditions requises pour la version 1.1 du manifeste. Même si le manifeste de votre complément utilise l’élément **VersionOverrides**, il est important d’inclure les éléments de la version 1.1 du manifeste afin que votre complément fonctionne avec des clients plus anciens qui ne prennent pas en charge **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p109">This article will cover the requirements for a v1.1 manifest. Even if your add-in manifest uses the **VersionOverrides** element, it is still important to include the v1.1 manifest elements to allow your add-in to work with older clients that do not support **VersionOverrides**.</span></span>

> [!NOTE]
> <span data-ttu-id="a9d93-p110">Outlook utilise un schéma pour valider les manifestes. Ce schéma requiert que les éléments du manifeste apparaissent dans un ordre spécifique. Si vous incluez des éléments dans un ordre autre que celui demandé, vous pouvez obtenir des erreurs lors du chargement de votre complément. Vous pouvez télécharger le [schéma de définition XML (XSD, XML Schema Definition)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) pour créer votre manifeste avec les éléments dans l’ordre requis.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p110">Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.</span></span>

## <a name="root-element"></a><span data-ttu-id="a9d93-139">Élément racine</span><span class="sxs-lookup"><span data-stu-id="a9d93-139">Root element</span></span>

<span data-ttu-id="a9d93-p111">L’élément racine du manifeste de complément Outlook est **OfficeApp**. Cet élément indique également l’espace de noms, la version de schéma et le type de complément par défaut. Placez tous les autres éléments du manifeste entre ses balises d’ouverture et de fermeture. Vous trouverez ci-dessous un exemple d’élément racine :</span><span class="sxs-lookup"><span data-stu-id="a9d93-p111">The root element for the Outlook add-in manifest is **OfficeApp**. This element also declares the default namespace, schema version and the type of add-in. Place all other elements in the manifest within its open and close tags. The following is an example of the root element:</span></span>


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest -->

</OfficeApp>
```

## <a name="version"></a><span data-ttu-id="a9d93-144">Version</span><span class="sxs-lookup"><span data-stu-id="a9d93-144">Version</span></span>

<span data-ttu-id="a9d93-p112">Il s’agit de la version du complément spécifique. Si un développeur met à jour un élément du manifeste, la version doit être incrémentée. Ainsi, lorsque le nouveau manifeste sera installé, il remplacera l’existant et l’utilisateur recevra les nouvelles fonctionnalités. Si ce complément a été soumis dans le magasin, le nouveau manifeste devra être soumis une deuxième fois et validé à nouveau. Ensuite, les utilisateurs de ce complément recevront le nouveau manifeste mis à jour automatiquement dans quelques heures, une fois approuvé.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p112">This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.</span></span>

<span data-ttu-id="a9d93-p113">If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p113">If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.</span></span>

## <a name="versionoverrides"></a><span data-ttu-id="a9d93-153">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="a9d93-153">VersionOverrides</span></span>

<span data-ttu-id="a9d93-154">L’élément **VersionOverrides** représente l’emplacement des informations pour les [commandes de complément](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="a9d93-154">The **VersionOverrides** element is the location of information for [add-in commands](add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="a9d93-155">Cet élément est également l’endroit où les compléments définissent la prise en charge des [compléments mobiles](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="a9d93-155">This element is also where add-ins define support for [mobile add-ins](add-mobile-support.md).</span></span>

<span data-ttu-id="a9d93-156">Pour plus d’informations sur cet élément, consultez [Créer des commandes complémentaires dans votre formulaire pour Excel, PowerPoint et Word](../develop/create-addin-commands.md).</span><span class="sxs-lookup"><span data-stu-id="a9d93-156">For a discussion on this element, see [Create add-in commands in your manifest for Excel, PowerPoint, and Word](../develop/create-addin-commands.md).</span></span>

## <a name="localization"></a><span data-ttu-id="a9d93-157">Localisation</span><span class="sxs-lookup"><span data-stu-id="a9d93-157">Localization</span></span>

<span data-ttu-id="a9d93-p114">Certains aspects du complément doivent être localisés pour les différents paramètres régionaux, tels que le nom, la description et l’URL qui est chargée. Ces éléments peuvent être facilement localisés en spécifiant la valeur par défaut et les valeurs de remplacement locales dans l’élément **Resources** au sein de l’élément **VersionOverrides**. Pour remplacer une image, une URL et une chaîne, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="a9d93-p114">Some aspects of the add-in need to be localized for different locales, such as the name, description and the URL that's loaded. These elements can easily be localized by specifying the default value and then locale overrides in the **Resources** element within the **VersionOverrides** element. The following shows how to override an image, a URL, and a string:</span></span>


```XML
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
      <!-- add information for other locales -->
    </bt:Image>
  </bt:Images>

  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
      <!-- add information for other locales -->
    </bt:Url>
  </bt:Urls>

  <bt:ShortStrings> 
    <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
      <bt:Override Locale="ar-sa" Value="<add localized value here>" />
      <!-- add information for other locales -->
    </bt:String>
  </bt:ShortStrings>
</Resources>
```

<span data-ttu-id="a9d93-161">La référence de schéma contient des informations complètes sur les éléments pouvant être localisés.</span><span class="sxs-lookup"><span data-stu-id="a9d93-161">The schema reference contains full information on which elements can be localized.</span></span>

## <a name="hosts"></a><span data-ttu-id="a9d93-162">Hôtes</span><span class="sxs-lookup"><span data-stu-id="a9d93-162">Hosts</span></span>

<span data-ttu-id="a9d93-163">Les compléments Outlook spécifient l’élément **Hosts** comme ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="a9d93-163">Outlook add-ins specify the **Hosts** element like the following.</span></span>

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

<span data-ttu-id="a9d93-164">Il existe une différence avec l’élément **Hosts** au sein de l’élément **VersionOverrides**, qui est abordée dans [Créer des commandes de complément dans votre manifeste pour Excel, PowerPoint et Word](../develop/create-addin-commands.md).</span><span class="sxs-lookup"><span data-stu-id="a9d93-164">This is separate from the **Hosts** element inside the **VersionOverrides** element, which is discussed in [Create add-in commands in your manifest for Excel, PowerPoint, and Word](../develop/create-addin-commands.md).</span></span>

## <a name="requirements"></a><span data-ttu-id="a9d93-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a9d93-165">Requirements</span></span>

<span data-ttu-id="a9d93-p115">L’élément **Requirements** spécifie l’ensemble d’API disponible pour le complément. Pour un complément Outlook, l’ensemble de conditions requises doit être Mailbox et avoir la valeur 1.1 ou supérieure. Reportez-vous à la référence d’API pour connaître la dernière version de condition requise. Pour plus d’informations sur les ensembles de conditions requises, reportez-vous à la rubrique [API de complément Outlook](apis.md). </span><span class="sxs-lookup"><span data-stu-id="a9d93-p115">The **Requirements** element specifies the set of APIs available to the add-in. For an Outlook add-in, the requirement set must be Mailbox and a value of 1.1 or above. Please refer to the API reference for the latest requirement set version. Refer to the [Outlook add-in APIs](apis.md) for more information on requirement sets.</span></span>

<span data-ttu-id="a9d93-170">L’élément **Requirements** peut également apparaître dans l’élément **VersionOverrides**, ce qui permet au complément de spécifier d’autres conditions requises lorsqu’il est chargé dans des clients qui prennent en charge **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="a9d93-170">The **Requirements** element can also appear in the **VersionOverrides** element, allowing the add-in to specify a different requirement when loaded in clients that support **VersionOverrides**.</span></span>

<span data-ttu-id="a9d93-171">L’exemple suivant utilise l’attribut **DefaultMinVersion** de l’élément **Sets** pour exiger office.js version 1.1 ou ultérieure, et l’attribut **MinVersion** de l’élément **Set** pour exiger l’ensemble de conditions requises Mail box version 1.1.</span><span class="sxs-lookup"><span data-stu-id="a9d93-171">The following example uses the **DefaultMinVersion** attribute of the **Sets** element to require office.js version 1.1 or higher, and the **MinVersion** attribute of the **Set** element to require the Mailbox requirement set version 1.1.</span></span>

```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```

## <a name="form-settings"></a><span data-ttu-id="a9d93-172">Paramètres de formulaire</span><span class="sxs-lookup"><span data-stu-id="a9d93-172">Form settings</span></span>

<span data-ttu-id="a9d93-p116">L’élément **FormSettings** est utilisé par les clients Outlook plus anciens, qui prennent en charge uniquement le schéma version 1.1 et non **VersionOverrides**. À l’aide de cet élément, les développeurs définissent la façon dont le complément s’affiche dans ces clients. Il existe deux parties : **ItemRead** et **ItemEdit**.**ItemRead** est utilisé pour spécifier la manière dont le complément apparaît lorsque l’utilisateur lit les messages et les rendez-vous. **ItemEdit** décrit comment le complément s’affiche lorsque l’utilisateur compose une réponse, un nouveau message, un nouveau rendez-vous ou modifie un rendez-vous dont il est l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p116">The **FormSettings** element is used by older Outlook clients, which only support schema 1.1 and not **VersionOverrides**. Using this element, developers define how the add-in will appear in such clients. There are two parts - **ItemRead** and **ItemEdit**. **ItemRead** is used to specify how the add-in appears when the user reads messages and appointments. **ItemEdit** describes how the add-in appears while the user is composing a reply, new message, new appointment or editing an appointment where they are the organizer.</span></span>

<span data-ttu-id="a9d93-p117">Ces paramètres sont directement liés aux règles d’activation dans l’élément **Rule**. Par exemple, si un complément spécifie qu’il doit apparaître sur un message lors de sa composition, un formulaire **ItemEdit** doit être spécifié.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p117">These settings are directly related to the activation rules in the **Rule** element. For example, if an add-in specifies that it should appear on a message in compose mode, an **ItemEdit** form must be specified.</span></span>

<span data-ttu-id="a9d93-180">Pour plus d’informations, voir Schema reference for Office Add-ins manifests (v1.1).</span><span class="sxs-lookup"><span data-stu-id="a9d93-180">For more details, please refer to the [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>

## <a name="app-domains"></a><span data-ttu-id="a9d93-181">Domaines d’application</span><span class="sxs-lookup"><span data-stu-id="a9d93-181">App domains</span></span>

<span data-ttu-id="a9d93-p118">Le domaine de la page de démarrage du complément que vous spécifiez dans l’élément **SourceLocation** est le domaine par défaut pour le complément. Si vous n’utilisez pas les éléments **AppDomains** et **AppDomain** et que votre complément tente d’accéder à un autre domaine, le navigateur ouvre une nouvelle fenêtre en dehors du panneau de complément. Afin que le complément puisse accéder à un autre domaine dans le volet de complément, ajoutez un élément **AppDomains** et incluez chaque domaine supplémentaire dans son propre sous-élément **AppDomain** dans le manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p118">The domain of the add-in start page that you specify in the **SourceLocation** element is the default domain for the add-in. Without using the **AppDomains** and **AppDomain** elements, if your add-in attempts to navigate to another domain, the browser will open a new window outside of the add-in pane. In order to allow the add-in to navigate to another domain within the add-in pane, add an **AppDomains** element and include each additional domain in its own **AppDomain** sub-element in the add-in manifest.</span></span>

<span data-ttu-id="a9d93-185">L’exemple suivant spécifie le domaine  `https://www.contoso2.com` comme second domaine auquel le complément peut accéder à l’intérieur du volet du complément :</span><span class="sxs-lookup"><span data-stu-id="a9d93-185">The following example specifies a domain  `https://www.contoso2.com` as a second domain that the add-in can navigate to within the add-in pane:</span></span>

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

<span data-ttu-id="a9d93-186">Les domaines d’application sont également nécessaires pour activer le partage entre la fenêtre contextuelle et le complément en cours d’exécution dans le client riche.</span><span class="sxs-lookup"><span data-stu-id="a9d93-186">App domains are also necessary to enable cookie sharing between the pop-out window and the add-in running in the rich client.</span></span>

<span data-ttu-id="a9d93-187">Le tableau suivant décrit le comportement du navigateur lorsque votre complément tente d’accéder à une URL en dehors du domaine par défaut du complément.</span><span class="sxs-lookup"><span data-stu-id="a9d93-187">The following table describes browser behavior when your add-in attempts to navigate to a URL outside of the add-in's default domain.</span></span>

|<span data-ttu-id="a9d93-188">Client Outlook</span><span class="sxs-lookup"><span data-stu-id="a9d93-188">Outlook client</span></span>|<span data-ttu-id="a9d93-189">Domaine défini</span><span class="sxs-lookup"><span data-stu-id="a9d93-189">Domain defined</span></span><br><span data-ttu-id="a9d93-190">dans AppDomains</span><span class="sxs-lookup"><span data-stu-id="a9d93-190">in AppDomains?</span></span>|<span data-ttu-id="a9d93-191">Comportement du navigateur</span><span class="sxs-lookup"><span data-stu-id="a9d93-191">Browser behavior</span></span>|
|---|---|---|
|<span data-ttu-id="a9d93-192">Tous les clients</span><span class="sxs-lookup"><span data-stu-id="a9d93-192">All clients</span></span>|<span data-ttu-id="a9d93-193">Oui</span><span class="sxs-lookup"><span data-stu-id="a9d93-193">Yes</span></span>|<span data-ttu-id="a9d93-194">Le lien s’ouvre dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="a9d93-194">Link opens in add-in task pane.</span></span>|
|<span data-ttu-id="a9d93-195">Outlook 2016 pour Windows (achat unique)</span><span class="sxs-lookup"><span data-stu-id="a9d93-195">Outlook 2016 on Windows (one-time purchase)</span></span><br><span data-ttu-id="a9d93-196">Outlook 2013 sous Windows</span><span class="sxs-lookup"><span data-stu-id="a9d93-196">Outlook 2013 on Windows</span></span>|<span data-ttu-id="a9d93-197">Non</span><span class="sxs-lookup"><span data-stu-id="a9d93-197">No</span></span>|<span data-ttu-id="a9d93-198">Le lien s’ouvre dans Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="a9d93-198">Link opens in Internet Explorer 11.</span></span>|
|<span data-ttu-id="a9d93-199">Autres clients</span><span class="sxs-lookup"><span data-stu-id="a9d93-199">Other clients</span></span>|<span data-ttu-id="a9d93-200">Non</span><span class="sxs-lookup"><span data-stu-id="a9d93-200">No</span></span>|<span data-ttu-id="a9d93-201">Le lien s’ouvre dans le navigateur par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a9d93-201">Link opens in user's default browser.</span></span>|

<span data-ttu-id="a9d93-202">Pour plus d’informations, voir [Spécifier les domaines que vous souhaitez ouvrir dans la fenêtre de complément](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).</span><span class="sxs-lookup"><span data-stu-id="a9d93-202">For more details, see the [Specify domains you want to open in the add-in window](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).</span></span>

## <a name="permissions"></a><span data-ttu-id="a9d93-203">Autorisations</span><span class="sxs-lookup"><span data-stu-id="a9d93-203">Permissions</span></span>

<span data-ttu-id="a9d93-p119">L’élément **Permissions** contient les autorisations requises pour le complément. Généralement, vous devez spécifier l’autorisation nécessaire minimale dont votre complément a besoin selon la méthode exacte que vous prévoyez d’utiliser. Par exemple, un complément de messagerie qui s’active dans les formulaires de composition et qui lit uniquement mais n’écrit pas dans les propriétés de l’élément comme [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), et qui n’appelle pas [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour accéder aux opérations liées aux services web Exchange doit spécifier l’autorisation **ReadItem**. Pour plus de détails sur les autorisations disponibles, reportez-vous à l’article [Présentation des autorisations de complément Outlook](understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="a9d93-p119">The **Permissions** element contains the required permissions for the add-in. In general, you should specify the minimum necessary permission that your add-in needs, depending on the exact methods that you plan to use. For example, a mail add-in that activates in compose forms and only reads but does not write to item properties like [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and does not call [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to access any Exchange Web Services operations should specify **ReadItem** permission. For details on the available permissions, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

<span data-ttu-id="a9d93-208">**Modèle d’autorisations à 4 niveaux pour les compléments de messagerie**</span><span class="sxs-lookup"><span data-stu-id="a9d93-208">**Four-tier permissions model for mail add-ins**</span></span>

![Modèle d’autorisations à 4 niveaux pour le schéma d’applications de messagerie v1.1](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a><span data-ttu-id="a9d93-210">Règles d’activation</span><span class="sxs-lookup"><span data-stu-id="a9d93-210">Activation rules</span></span>

<span data-ttu-id="a9d93-p120">Les règles d’activation sont spécifiées dans l’élément **Rule**. L’élément **Rule** peut apparaître comme un enfant de l’élément **OfficeApp** dans les manifestes 1.1.</span><span class="sxs-lookup"><span data-stu-id="a9d93-p120">Activation rules are specified in the **Rule** element. The **Rule** element can appear as a child of the **OfficeApp** element in 1.1 manifests.</span></span>

<span data-ttu-id="a9d93-213">Les règles d’activation peuvent être utilisées pour activer un complément basé sur une ou plusieurs des conditions suivantes sur l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a9d93-213">Activation rules can be used to activate an add-in based on one or more of the following conditions on the currently selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="a9d93-214">Les règles d’activation s’appliquent uniquement aux clients qui ne prennent pas en charge l’élément **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="a9d93-214">Activation rules only apply to clients that do not support the **VersionOverrides** element.</span></span>

- <span data-ttu-id="a9d93-215">Le type d’élément et/ou la classe de message</span><span class="sxs-lookup"><span data-stu-id="a9d93-215">The item type and/or message class</span></span>

- <span data-ttu-id="a9d93-216">La présence d’un type spécifique d’entité connue, comme une adresse ou un numéro de téléphone</span><span class="sxs-lookup"><span data-stu-id="a9d93-216">The presence of a specific type of known entity, such as an address or phone number</span></span>

- <span data-ttu-id="a9d93-217">Une correspondance d’expression régulière dans le corps, l’objet ou l’adresse e-mail de l’expéditeur</span><span class="sxs-lookup"><span data-stu-id="a9d93-217">A regular expression match in the body, subject, or sender email address</span></span>

- <span data-ttu-id="a9d93-218">L’existence d’une pièce jointe</span><span class="sxs-lookup"><span data-stu-id="a9d93-218">The presence of an attachment</span></span>

<span data-ttu-id="a9d93-219">Pour plus de détails et des exemples de règles d’activation, voir [Règles d’activation pour les compléments Outlook](activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="a9d93-219">For details and samples of activation rules, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>


## <a name="next-steps-add-in-commands"></a><span data-ttu-id="a9d93-220">Prochaines étapes : commandes de complément</span><span class="sxs-lookup"><span data-stu-id="a9d93-220">Next steps: Add-in commands</span></span>

<span data-ttu-id="a9d93-221">Une fois que vous avez défini un manifeste de base, définissez les commandes de complément pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="a9d93-221">After defining a basic manifest, define add-in commands for your add-in.</span></span> <span data-ttu-id="a9d93-222">Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way.</span><span class="sxs-lookup"><span data-stu-id="a9d93-222">Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way.</span></span> <span data-ttu-id="a9d93-223">For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="a9d93-223">For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="a9d93-224">Pour un exemple de complément qui définit les commandes de complément, voir [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span><span class="sxs-lookup"><span data-stu-id="a9d93-224">For an example add-in that defines add-in commands, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span></span>

## <a name="next-steps-add-mobile-support"></a><span data-ttu-id="a9d93-225">Étapes suivantes : Ajouter la prise en charge mobile</span><span class="sxs-lookup"><span data-stu-id="a9d93-225">Next steps: Add mobile support</span></span>

<span data-ttu-id="a9d93-p122">Les compléments peuvent éventuellement ajouter la prise en charge d’Outlook Mobile. Outlook Mobile prend en charge les commandes de complément de la même manière qu’Outlook sous Windows et Mac. Pour plus d’informations, voir la section [Ajouter la prise en charge des commandes de complément pour Outlook Mobile](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="a9d93-p122">Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and Mac. For more information, see [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a9d93-229">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a9d93-229">See also</span></span>

- [<span data-ttu-id="a9d93-230">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="a9d93-230">Localization for Office Add-ins</span></span>](../develop/localization.md)
- [<span data-ttu-id="a9d93-231">Confidentialité, autorisations et sécurité pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="a9d93-231">Privacy, permissions, and security for Outlook add-ins</span></span>](privacy-and-security.md)
- [<span data-ttu-id="a9d93-232">API de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="a9d93-232">Outlook add-in APIs</span></span>](apis.md)
- [<span data-ttu-id="a9d93-233">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="a9d93-233">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="a9d93-234">Référence de schéma pour les manifestes des compléments Office (version 1.1)</span><span class="sxs-lookup"><span data-stu-id="a9d93-234">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="a9d93-235">Concevoir vos compléments Office</span><span class="sxs-lookup"><span data-stu-id="a9d93-235">Design your Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="a9d93-236">Présentation des autorisations de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="a9d93-236">Understanding Outlook add-in permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="a9d93-237">Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="a9d93-237">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="a9d93-238">Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues</span><span class="sxs-lookup"><span data-stu-id="a9d93-238">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)