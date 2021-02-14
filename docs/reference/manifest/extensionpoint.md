---
title: Élément Extension dans le fichier manifeste
description: Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 20e1f58070d61b02a1c2c2fcefc4ce2b0ad94979
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237706"
---
# <a name="extensionpoint-element"></a>Élément ExtensionPoint

 Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office. L’élément **ExtensionPoint** est un élément enfant de [AllFormFactors](allformfactors.md) ou [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Oui  | Type de point d’extension défini.|

## <a name="extension-points-for-excel-only"></a>Points d’extension pour Excel uniquement

- **CustomFunctions** – fonction personnalisée écrite en JavaScript pour Excel.

[L’exemple de code XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) montre comment utiliser l’élément **ExtensionPoint** avec la valeur d’attribut **CustomFunctions** et les éléments enfants à utiliser.

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>Points d’extension pour les commandes de complément Word, Excel, PowerPoint et OneNote

- **PrimaryCommandSurface** : ruban dans Office.
- **ContextMenu** : menu contextuel qui apparaît lorsque vous cliquez avec le bouton droit de la souris dans l’interface utilisateur Office.

Les exemples suivants montrent comment utiliser l’élément **ExtensionPoint** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.

> [!IMPORTANT]
> Pour les éléments qui contiennent un attribut ID, assurez-vous que vous indiquez un ID unique. Nous vous recommandons d’utiliser le nom de votre organisation, ainsi que votre ID. Par exemple, utilisez le format suivant. <CustomTab id="mycompanyname.mygroupname">

```XML
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

#### <a name="child-elements"></a>Éléments enfants
 
|Élément|Description|
|:-----|:-----|
|**CustomTab**|Obligatoire si vous souhaitez ajouter un onglet personnalisé au ruban (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **CustomTab**, vous ne pouvez pas utiliser l’élément **OfficeTab**. L’attribut **id** est obligatoire. |
|**OfficeTab**|Obligatoire si vous souhaitez étendre un onglet de ruban d’application Office par défaut (à l’aide **de PrimaryCommandSurface).** Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**. Pour plus d’informations, voir [OfficeTab](officetab.md).|
|**OfficeMenu**|Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : <br/> - **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. <br/> - **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul.|
|**Group**|Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est obligatoire. Il s’agit d’une chaîne avec un maximum de 125 caractères. |
|**Label**|Obligatoire. Libellé du groupe. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un **élément String.** L’élément  **String** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément  **Resources**.|
|**Icon**|Obligatoire. Indique l’icône du groupe qui doit être utilisée sur les périphériques de petit facteur de forme ou lorsque les boutons sont affichés en trop grand nombre. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un **élément Image.** L’élément  **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément  **Resources**. L’attribut **size** donne la taille, en pixels, de l’image. Trois tailles d’image, en pixels, sont obligatoires : 16, 32 et 80. Cinq tailles facultatives, en pixels, sont également prises en charge : 20, 24, 40, 48 et 64.|
|**Tooltip**|Facultatif. Info-bulle du groupe. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un **élément String.** L’élément  **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément  **Resources**.|
|**Control**|Chaque groupe requiert au moins un contrôle. Un élément **Control** peut être de type **Button** ou **Menu**. Utilisez **Menu** pour spécifier une liste déroulante de contrôles de bouton. Actuellement, seuls les boutons et les menus sont pris en charge. Pour plus d’informations, reportez-vous aux sections [Contrôles de bouton](control.md#button-control) et [Contrôles de menu](control.md#menu-dropdown-button-controls).<br/>**Remarque :**  Pour faciliter la résolution des problèmes, nous vous recommandons d’ajouter un élément **Control** et les éléments enfants **Resources** associés un par un.|
|**Script**|Liens vers le fichier JavaScript avec la définition de la fonction personnalisée et le code d’inscription. Cet élément n’est pas utilisé dans l’aperçu pour les développeurs. À la place, la page HTML est responsable du chargement de tous les fichiers JavaScript.|
|**Page**|Liens vers la page HTML de vos fonctions personnalisées.|

## <a name="extension-points-for-outlook"></a>Points d’extension pour Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (peut uniquement être utilisé dans [DesktopFormFactor](desktopformfactor.md).)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface)
- [LaunchEvent](#launchevent-preview)
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
  <CustomTab id="TabCustom1">
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
  <CustomTab id="TabCustom1">
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
  <CustomTab id="TabCustom1">
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
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

Ce point d’extension place des boutons sur le ruban pour l’extension de module.

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
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="mobileonlinemeetingcommandsurface"></a>MobileOnlineMeetingCommandSurface

Ce point d’extension place un basculement approprié en mode dans l’surface de commande d’un rendez-vous dans le facteur de forme mobile. Un organisateur de réunion peut créer une réunion en ligne. Un participant peut ensuite participer à la réunion en ligne. Pour en savoir plus sur ce scénario, consultez l’article Créer un application mobile Outlook pour un fournisseur de réunion [en ligne.](../../outlook/online-meeting.md)

> [!NOTE]
> Ce point d’extension est uniquement pris en charge sur Android avec un abonnement Microsoft 365.

#### <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [Control](control.md) |  Ajoute un bouton à la surface de commande.  |

`ExtensionPoint` les éléments de ce type ne peuvent avoir qu’un seul élément enfant : un `Control` élément.

L’attribut doit être attribué à l’élément contenu dans ce `Control` point `xsi:type` d’extension. `MobileButton`

Les images doivent être en échelles de gris à l’aide de code hex ou de son équivalent `Icon` `#919191` dans [d’autres formats de couleur.](https://convertingcolors.com/hex-color-919191.html)

#### <a name="example"></a>Exemple

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
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

### <a name="launchevent-preview"></a>LaunchEvent (aperçu)

> [!NOTE]
> Ce point d’extension est uniquement pris en charge en [prévisualisation](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) dans Outlook sur le web et sur Windows avec un abonnement Microsoft 365.

Ce point d’extension permet à un application de s’activer en fonction des événements pris en charge dans le facteur de forme de bureau. Actuellement, les seuls événements pris en charge sont `OnNewMessageCompose` et `OnNewAppointmentOrganizer` . Pour en savoir plus sur ce scénario, consultez l’article Configurer votre complément Outlook pour [l’activation](../../outlook/autolaunch.md) basée sur un événement.

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

Ce point d’extension ajoute un gestionnaire d’événements pour un événement spécifié. Pour plus d’informations sur l’utilisation de ce point d’extension, voir La fonctionnalité d’envoi pour [les modules complémentaires Outlook.](../../outlook/outlook-on-send-addins.md)

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

Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, `xsi:type`l’attribut doit avoir la valeur `VersionOverridesV1_1`.

> [!NOTE]
> Ce type d’élément est disponible pour [les clients Outlook qui prennent en charge les ensembles de conditions requises 1.6 ou version ultérieure](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).

|  Élément |  Description  |
|:-----|:-----|
|  [Label](#label) |  Spécifie l’étiquette pour le complément dans la fenêtre contextuelle.  |
|  [SourceLocation](sourcelocation.md) |  Spécifie l’URL de la fenêtre contextuelle.  |
|  [Règle](rule.md) |  Spécifie la ou les règles qui déterminent lorsqu’un complément s’active.  |

#### <a name="label"></a>Étiquette

Obligatoire. Libellé du groupe. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources.](resources.md)

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
