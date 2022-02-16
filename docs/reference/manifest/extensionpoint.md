---
title: Élément Extension dans le fichier manifeste
description: Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8ccc08a9c0d42edf89c904b8809a530239be4c
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855631"
---
# <a name="extensionpoint-element"></a>Élément ExtensionPoint

 Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office. L’élément **ExtensionPoint** est un élément enfant de [AllFormFactors](allformfactors.md) ou [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Oui  | Type de point d’extension défini. Les valeurs possibles dépendent de l Office’application hôte définie dans la valeur de **l’élément Host** très grande.|

## <a name="extension-points-for-excel-onenote-powerpoint-and-word-add-in-commands"></a>Points d’extension pour Excel, OneNote, PowerPoint et les commandes de modules supplémentaires Word

Il existe trois types de points d’extension disponibles dans tout ou partie de ces hôtes.

- [PrimaryCommandSurface](#primarycommandsurface) (valide pour Word, Excel, PowerPoint et OneNote) : ruban dans Office.
- [ContextMenu](#contextmenu) (valide pour Word, Excel, PowerPoint et OneNote) : menu contextiqué qui s’affiche lorsque vous sélectionnez et maintenez (ou cliquez avec le bouton droit) dans l’interface utilisateur Office.
- [CustomFunctions](#customfunctions) (valide uniquement pour Excel) : fonction personnalisée écrite en JavaScript pour Excel.

Consultez les sous-sections suivantes pour les éléments enfants et des exemples de ces types de points d’extension.

### <a name="primarycommandsurface"></a>PrimaryCommandSurface

La surface de commande principale dans Word, Excel, PowerPoint et OneNote est le ruban.

#### <a name="child-elements"></a>Éléments enfants

|Élément|Description|
|:-----|:-----|
|[CustomTab] (customtab.md|Obligatoire si vous souhaitez ajouter un onglet personnalisé au ruban (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **CustomTab**, vous ne pouvez pas utiliser l’élément **OfficeTab**. L’attribut **id** est obligatoire. |
|[OfficeTab](officetab.md)|Obligatoire si vous souhaitez étendre un onglet application Office ruban par défaut (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**.|

#### <a name="example"></a>Exemple

L’exemple suivant montre comment utiliser **l’élément ExtensionPoint** avec **PrimaryCommandSurface**. Il ajoute un onglet personnalisé au ruban.

> [!IMPORTANT]
> Pour les éléments qui contiennent un attribut ID, veillez à fournir un ID unique.

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.MyTab1">
    <Label resid="residLabel4" />
    <Group id="Contoso.Group1">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Contoso.Button1">
          <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
```

### <a name="contextmenu"></a>ContextMenu

Un menu contextiqué est un menu contextiqué qui s’affiche lorsque vous cliquez avec le bouton droit dans Office’interface utilisateur.

#### <a name="child-elements"></a>Éléments enfants
 
|Élément|Description|
|:-----|:-----|
|[OfficeMenu](officemenu.md)|Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). **L’attribut id** doit être définie sur l’une des chaînes suivantes : <br/> - **ContextMenuText** si le menu contextuel doit s’ouvrir lorsqu’un utilisateur clique avec le bouton droit sur le texte sélectionné. <br/> - **ContextMenuCell** si le menu contextiqué doit s’ouvrir lorsque l’utilisateur clique avec le bouton droit sur une cellule d’Excel feuille de calcul.|

#### <a name="example"></a>Exemple

L’exemple suivant ajoute un menu contexté personnalisé aux cellules d’une feuille Excel feuille de calcul.

```xml
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="Contoso.ContextMenu2">
            <!-- information about the control -->
    </Control>
    <!-- other controls, as needed -->
  </OfficeMenu>
</ExtensionPoint>
```

### <a name="customfunctions"></a>CustomFunctions

Fonction personnalisée écrite en JavaScript ou TypeScript pour Excel.

#### <a name="child-elements"></a>Éléments enfants

|Élément|Description|
|:-----|:-----|
|[Script](script.md)|Obligatoire. Liens vers le fichier JavaScript avec la définition et le code d’inscription de la fonction personnalisée.|
|[Page](page.md)|Obligatoire. Liens vers la page HTML de vos fonctions personnalisées.|
|[MetaData](metadata.md)|Obligatoire. Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.|
|[Namespace](namespace.md)|Facultatif. Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.|

#### <a name="example"></a>Exemple

```xml
<ExtensionPoint xsi:type="CustomFunctions">
  <Script>
    <SourceLocation resid="Functions.Script.Url"/>
  </Script>
  <Page>
    <SourceLocation resid="Shared.Url"/>
  </Page>
  <Metadata>
    <SourceLocation resid="Functions.Metadata.Url"/>
  </Metadata>
  <Namespace resid="Functions.Namespace"/>
</ExtensionPoint>
```

## <a name="extension-points-for-outlook"></a>Points d’extension pour Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (peut uniquement être utilisé dans [DesktopFormFactor](desktopformfactor.md).)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface)
- [LaunchEvent](#launchevent)
- [Événements](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface

Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique. Dans l’application de bureau Outlook, cela apparaît dans le ruban.

#### <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

#### <a name="officetab-example"></a>Exemple OfficeTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Exemple CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="Contoso.TabCustom2">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface

Ce point d’extension place des boutons sur le ruban pour les compléments à l’aide du formulaire de composition de messagerie. 

#### <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

#### <a name="officetab-example"></a>Exemple OfficeTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Exemple CustomTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="Contoso.TabCustom3">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention de l’organisateur de la réunion. 

#### <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

#### <a name="officetab-example"></a>Exemple OfficeTab

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Exemple CustomTab

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="Contoso.TabCustom4">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention du participant à la réunion. 

#### <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

#### <a name="officetab-example"></a>Exemple OfficeTab

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Exemple CustomTab

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="Contoso.TabCustom5">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

Ce point d’extension place des boutons sur le ruban pour l’extension de module.

> [!IMPORTANT]
> L’inscription des [événements de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) [boîte aux](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) lettres et d’élément n’est pas disponible avec ce point d’extension.

#### <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

### <a name="mobilemessagereadcommandsurface"></a>MobileMessageReadCommandSurface

Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique dans le facteur de forme pour environnement mobile.

#### <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [Group](group.md) |  Ajoute un groupe de boutons à la surface de commande.  |

Les éléments **ExtensionPoint** de ce type peuvent uniquement avoir un élément enfant, à savoir un élément **Group**.

Pour les éléments **Control** contenus dans ce point d’extension, l’attribut **xsi:type** doit avoir la valeur `MobileButton`.

#### <a name="example"></a>Exemple

```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="Contoso.mobileGroup1">
    <Label resid="residAppName"/>
      <Control  xsi:type="MobileButton id="Contoso.mobileButton1"">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="mobileonlinemeetingcommandsurface"></a>MobileOnlineMeetingCommandSurface

Ce point d’extension place un basculement adapté au mode dans l’surface de commande d’un rendez-vous dans le facteur de forme mobile. Un organisateur de réunion peut créer une réunion en ligne. Un participant peut ensuite participer à la réunion en ligne. Pour en savoir plus sur ce scénario, consultez l’article [Créer un Outlook mobile pour](../../outlook/online-meeting.md) un fournisseur de réunion en ligne.

> [!NOTE]
> Ce point d’extension est uniquement pris en charge sur Android et iOS avec Microsoft 365 abonnement.
>
> L’inscription des [événements de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) [boîte aux](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) lettres et d’élément n’est pas disponible avec ce point d’extension.

#### <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [Control](control.md) |  Ajoute un bouton à la surface de commande.  |

`ExtensionPoint` les éléments de ce type ne peuvent avoir qu’un seul élément enfant : un `Control` élément.

L’attribut `Control` doit être attribué à l’élément contenu dans ce point `xsi:type` d’extension `MobileButton`.

Les `Icon` images doivent être en échelles de gris à l’aide de code hex ou `#919191` de son équivalent dans [d’autres formats de couleur.](https://convertingcolors.com/hex-color-919191.html)

#### <a name="example"></a>Exemple

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="Contoso.onlineMeetingFunctionButton1">
    <Label resid="residUILessButton0Name" />
    <Icon>
      <bt:Image resid="UiLessIcon" size="25" scale="1" />
      <bt:Image resid="UiLessIcon" size="25" scale="2" />
      <bt:Image resid="UiLessIcon" size="25" scale="3" />
      <bt:Image resid="UiLessIcon" size="32" scale="1" />
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
      <bt:Image resid="UiLessIcon" size="32" scale="3" />
      <bt:Image resid="UiLessIcon" size="48" scale="1" />
      <bt:Image resid="UiLessIcon" size="48" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>insertContosoMeeting</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### <a name="launchevent"></a>LaunchEvent

Ce point d’extension permet à un application de s’activer en fonction des événements pris en charge dans le facteur de forme de bureau. Pour en savoir plus sur ce scénario et pour obtenir la liste complète des événements pris en charge, consultez l’article Configurer votre complément [Outlook pour l’activation basée sur des événements](../../outlook/autolaunch.md).

> [!IMPORTANT]
> L’inscription des [événements de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) [boîte aux](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) lettres et d’élément n’est pas disponible avec ce point d’extension.

#### <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  Liste de [LaunchEvent pour](launchevent.md) l’activation basée sur des événements.  |
| [SourceLocation](sourcelocation.md) |  Emplacement du fichier JavaScript source.  |

#### <a name="example"></a>Exemple

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

### <a name="events"></a>Événements

Ce point d’extension ajoute un gestionnaire d’événements pour un événement spécifié. Pour plus d’informations sur l’utilisation de ce point d’extension, consultez la fonctionnalité d’envoi [Outlook des modules complémentaires](../../outlook/outlook-on-send-addins.md).

> [!IMPORTANT]
> L’inscription des [événements de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) [boîte aux](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) lettres et d’élément n’est pas disponible avec ce point d’extension.

| Élément | Description  |
|:-----|:-----|
|  [Event](event.md) |  Indique l’événement et la fonction gestionnaire d’événements.  |

#### <a name="itemsend-event-example"></a>Exemple d’événement ItemSend

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a>DetectedEntity

Ce point d’extension ajoute une activation de complément contextuel sur un type d’entité spécifié.

> [!IMPORTANT]
> L’inscription des [événements de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) [boîte aux](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) lettres et d’élément n’est pas disponible avec ce point d’extension.

Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, `xsi:type`l’attribut doit avoir la valeur `VersionOverridesV1_1`.

> [!NOTE]
> Ce type d’élément est disponible pour [les clients Outlook qui prennent en charge les ensembles de conditions requises 1.6 ou version ultérieure](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).

|  Élément |  Description  |
|:-----|:-----|
|  [Label](#label) |  Spécifie l’étiquette pour le complément dans la fenêtre contextuelle.  |
|  [SourceLocation](sourcelocation.md) |  Spécifie l’URL de la fenêtre contextuelle.  |
|  [Règle](rule.md) |  Spécifie la ou les règles qui déterminent lorsqu’un complément s’active.  |

#### <a name="label"></a>Étiquette

Obligatoire. Libellé du groupe. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md).

#### <a name="highlight-requirements"></a>Exigences relatives à la mise en surbrillance

Le seul moyen pour qu’un utilisateur puisse activer un complément contextuel consiste à interagir avec une entité en surbrillance. Les développeurs peuvent contrôler les entités qui sont mises en surbrillance à l’aide de l’attribut `Highlight` de l’élément `Rule` pour les types de règles `ItemHasKnownEntity` et `ItemHasRegularExpressionMatch`.

Toutefois, il existe certaines limitations à connaître. Ces limitations sont en place pour vous assurer qu’il y aura toujours une entité en surbrillance dans les messages ou rendez-vous applicables pour permettre à l’utilisateur d’activer le complément.

- Les types d’entité `EmailAddress` et `Url` ne peuvent pas être mis en surbrillance et par conséquent ne peuvent pas être utilisés pour activer un complément.
- Si vous utilisez une seule règle, la valeur `Highlight` DOIT être définie sur `all`.
- Si vous utilisez un type de règle `RuleCollection` avec `Mode="AND"` pour combiner plusieurs règles, au moins l’une des règles DOIT définir `Highlight` sur la valeur `all`.
- Si vous utilisez un type de règle `RuleCollection` avec `Mode="OR"` pour combiner plusieurs règles, toutes les règles DOIVENT définir `Highlight` sur la valeur `all`.

#### <a name="detectedentity-event-example"></a>Exemple d’événement DetectedEntity

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint>
```
