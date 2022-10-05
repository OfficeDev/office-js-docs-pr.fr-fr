---
title: Manifestes des compléments Outlook
description: Le manifeste décrit l’intégration d’un complément Outlook avec les clients Outlook et comprend un exemple.
ms.date: 05/27/2020
ms.localizationpriority: high
ms.openlocfilehash: c09c483519e4d5cd0dce7dda840130698820b6ee
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68466977"
---
# <a name="outlook-add-in-manifests"></a>Manifestes des compléments Outlook

An Outlook add-in consists of two components: the XML add-in manifest and a web page supported by the JavaScript library for Office Add-ins (office.js). The manifest describes how the add-in integrates across Outlook clients. The following is an example.

 > [!NOTE]
 > All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.

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

## <a name="schema-versions"></a>Versions de schéma

Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.

L’élément **\<VersionOverrides\>** dans le manifeste en est un exemple. Tous les éléments définis à l’intérieur de **\<VersionOverrides\>** remplacent le même élément dans l’autre partie du manifeste. Cela signifie que, dans la mesure du possible, Outlook utilise ce qui se trouve dans la section **\<VersionOverrides\>** pour configurer le complément. Toutefois, si la version d’Outlook ne prend pas en charge une certaine version de **\<VersionOverrides\>**, Outlook l’ignore et dépend des informations contenues dans le reste du manifeste. 

Cette approche signifie que les développeurs ne doivent pas créer plusieurs manifestes individuels, mais plutôt conserver tous les éléments définis dans un fichier.

Les versions actuelles du schéma sont les suivantes :


|Version|Description|
|:-----|:-----|
|v1.0|Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form. |
|v1.1|Prend en charge la version 1.1 de l’API JavaScript Office et **\<VersionOverrides\>**. Pour les compléments Outlook, la prise en charge des formulaires de composition est incluse.|
|**\<VersionOverrides\>** 1.0|Prend en charge les versions ultérieures de l’API JavaScript Office. La prise en charge des commandes de complément est incluse.|
|**\<VersionOverrides\>** 1.1|Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.|

Cet article porte sur les conditions requises pour la version 1.1 du manifeste. Même si votre manifeste de complément utilise l’élément **\<VersionOverrides\>** , il est toujours important d’inclure les éléments de manifeste v1.1 pour permettre à votre complément d’utiliser des clients plus anciens qui ne prennent pas en charge **\<VersionOverrides\>**.

> [!NOTE]
> Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.

## <a name="root-element"></a>Élément racine

L’élément racine du manifeste du complément Outlook est **\<OfficeApp\>**. Cet élément indique également l’espace de noms, la version de schéma et le type de complément par défaut. Placez tous les autres éléments du manifeste entre ses balises d’ouverture et de fermeture. Voici un exemple de l’élément racine.


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

## <a name="version"></a>Version

This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.

If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.

## <a name="versionoverrides"></a>VersionOverrides

 **L’élément\<VersionOverrides\>** est l’emplacement des informations pour [les commandes de complément](add-in-commands-for-outlook.md).

Cet élément est également l’endroit où les compléments définissent la prise en charge des [compléments mobiles](add-mobile-support.md).

Pour plus d’informations sur cet élément, consultez [Créer des commandes complémentaires dans votre formulaire pour Excel, PowerPoint et Word](../develop/create-addin-commands.md).

## <a name="localization"></a>Localisation

Certains aspects du complément doivent être localisés pour les différents paramètres régionaux, tels que le nom, la description et l’URL qui est chargée. Ces éléments peuvent facilement être localisés en spécifiant la valeur par défaut, puis en remplaçant les paramètres régionaux dans l’élément **\<Resources\>** dans l’élément **\<VersionOverrides\>** . L’exemple suivant montre comment remplacer une image, une URL et une chaîne.


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

La référence de schéma contient des informations complètes sur les éléments pouvant être localisés.

## <a name="hosts"></a>Hôtes

Les compléments Outlook spécifient l’élément **\<Hosts\>** comme suit :

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

Il est séparé de l’élément **\<Hosts\>** à l’intérieur de l’élément **\<VersionOverrides\>** , qui est abordé dans [Créer des commandes de complément dans votre manifeste pour Excel, PowerPoint et Word](../develop/create-addin-commands.md).

## <a name="requirements"></a>Conditions requises

 **L’élément\<Requirements\>** spécifie l’ensemble d’API disponibles pour le complément. Pour un complément Outlook, l’ensemble de conditions requises doit être Mailbox et avoir la valeur 1.1 ou supérieure. Reportez-vous à la référence d’API pour connaître la dernière version de condition requise. Pour plus d’informations sur les ensembles de conditions requises, voir [API de complément Outlook](apis.md).

 **L’élément\<Requirements\>** peut également apparaître dans l’élément **\<VersionOverrides\>** , ce qui permet au complément de spécifier une exigence différente lors du chargement dans les clients qui prennent en charge **\<VersionOverrides\>**.

L’exemple suivant utilise l’attribut **DefaultMinVersion** de l’élément **\<Sets\>** pour exiger office.js version 1.1 ou ultérieure, et l’attribut **MinVersion** de l’élément **\<Set\>** pour exiger l’ensemble de conditions requises de boîte aux lettres version 1.1.

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

## <a name="form-settings"></a>Paramètres de formulaire

 **L’élément\<FormSettings\>** est utilisé par les anciens clients Outlook, qui prennent uniquement en charge le schéma 1.1 et non **\<VersionOverrides\>**. À l’aide de cet élément, les développeurs définissent la façon dont le complément s’affiche dans ces clients. Il existe deux parties : **ItemRead** et **ItemEdit**. **ItemRead** est utilisé pour spécifier la manière dont le complément apparaît lorsque l’utilisateur lit les messages et les rendez-vous. **ItemEdit** décrit comment le complément s’affiche lorsque l’utilisateur compose une réponse, un nouveau message, un nouveau rendez-vous ou modifie un rendez-vous dont il est l’organisateur.

Ces paramètres sont directement liés aux règles d’activation dans l’élément **\<Rule\>** . Par exemple, si un complément spécifie qu’il doit apparaître sur un message lors de sa composition, un formulaire  **ItemEdit** doit être spécifié.

Pour plus d’informations, voir Schema reference for Office Add-ins manifests (v1.1).

## <a name="app-domains"></a>Domaines d’application

Le domaine de la page de démarrage du complément que vous spécifiez dans l’élément **\<SourceLocation\>** est le domaine par défaut du complément. Sans utiliser les **éléments\<AppDomains\>** et **\<AppDomain\>** , si votre complément tente d’accéder à un autre domaine, le navigateur ouvre une nouvelle fenêtre en dehors du volet du complément. Pour permettre au complément d’accéder à un autre domaine dans le volet du complément, ajoutez un **élément\<AppDomains\>** et incluez chaque domaine supplémentaire dans son propre **sous-élément\<AppDomain\>** dans le manifeste du complément.

L’exemple suivant spécifie le domaine  `https://www.contoso2.com` comme second domaine auquel le complément peut accéder à l’intérieur du volet du complément.

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

Les domaines d’application sont également nécessaires pour activer le partage entre la fenêtre contextuelle et le complément en cours d’exécution dans le client riche.

Le tableau suivant décrit le comportement du navigateur lorsque votre complément tente d’accéder à une URL en dehors du domaine par défaut du complément.

|Client Outlook|Domaine défini<br>dans AppDomains|Comportement du navigateur|
|---|---|---|
|Tous les clients|Oui|Le lien s’ouvre dans le volet Office du complément.|
|- Outlook 2016 sur Windows (perpétuel avec licence en volume)<br>- Outlook 2013 sur Windows (perpétuel)|Non|Le lien s’ouvre dans Internet Explorer 11.|
|Autres clients|Non|Le lien s’ouvre dans le navigateur par défaut de l’utilisateur.|

Pour plus d’informations, voir [Spécifier les domaines que vous souhaitez ouvrir dans la fenêtre de complément](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).

## <a name="permissions"></a>Autorisations

 **L’élément\<Permissions\>** contient les autorisations requises pour le complément. Généralement, vous devez spécifier l’autorisation nécessaire minimale dont votre complément a besoin selon la méthode exacte que vous prévoyez d’utiliser. Par exemple, un complément de messagerie qui s’active dans les formulaires de composition et qui lit uniquement mais n’écrit pas dans les propriétés de l’élément comme [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), et qui n’appelle pas [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) pour accéder aux opérations liées aux services web doit spécifier l’autorisation **ReadItem**. Pour plus de détails sur les autorisations disponibles, voir [Spécifier les autorisations pour l’accès du complément Outlook à la boîte aux lettres de l’utilisateur](understanding-outlook-add-in-permissions.md).

**Modèle d’autorisations à 4 niveaux pour les compléments de messagerie**

![Modèle d’autorisations à 4 niveaux pour le schéma d’applications de messagerie v1.1.](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a>Règles d’activation

Les règles d’activation sont spécifiées dans l’élément **\<Rule\>** .  **L’élément\<Rule\>** peut apparaître en tant qu’enfant de l’élément **\<OfficeApp\>** dans les manifestes 1.1.

Les règles d’activation peuvent être utilisées pour activer un complément basé sur une ou plusieurs des conditions suivantes sur l’élément sélectionné.

> [!NOTE]
> Les règles d’activation s’appliquent uniquement aux clients qui ne prennent pas en charge l’élément **\<VersionOverrides\>** .

- Le type d’élément et/ou la classe de message

- La présence d’un type spécifique d’entité connue, comme une adresse ou un numéro de téléphone

- Une correspondance d’expression régulière dans le corps, l’objet ou l’adresse e-mail de l’expéditeur

- L’existence d’une pièce jointe

Pour plus de détails et des exemples de règles d’activation, voir [Règles d’activation pour les compléments Outlook](activation-rules.md).


## <a name="next-steps-add-in-commands"></a>Prochaines étapes : commandes de complément

After defining a basic manifest, define add-in commands for your add-in. Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

Pour un exemple de complément qui définit les commandes de complément, voir [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).

## <a name="next-steps-add-mobile-support"></a>Étapes suivantes : Ajouter la prise en charge mobile

Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and Mac. For more information, see [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).

## <a name="see-also"></a>Voir aussi

- [Localisation des compléments Office](../develop/localization.md)
- [Confidentialité, autorisations et sécurité pour les compléments Outlook](privacy-and-security.md)
- [API de complément Outlook](apis.md)
- [Manifeste XML des compléments Office](../develop/add-in-manifests.md)
- [Référence de schéma pour les manifestes des compléments Office (version 1.1)](../develop/add-in-manifests.md)
- [Concevoir vos compléments Office](../design/add-in-design.md)
- [Présentation des autorisations de complément Outlook](understanding-outlook-add-in-permissions.md)
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md)