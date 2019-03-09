---
title: Élément Extension dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Priority
ms.openlocfilehash: 7555c5e4245da55fd4d01761e0484912fe54eccc
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512873"
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

Les exemples suivants montrent comment utiliser l’élément  **ExtensionPoint** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.

> [!IMPORTANT] 
> Pour les éléments qui contiennent un attribut ID, veillez à fournir un ID unique. Nous recommandons d’utiliser le nom de votre société, ainsi que votre identifiant. Par exemple, utilisez la syntaxe suivante. <CustomTab id="mycompanyname.mygroupname">

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
 
|**Élément**|**Description**|
|:-----|:-----|
|**CustomTab**|Obligatoire pour ajouter un onglet personnalisé au ruban (en utilisant  **PrimaryCommandSurface**). Si vous utilisez l’élément  **CustomTab**, vous ne pouvez pas utiliser l’élément  **OfficeTab**. L’attribut  **id** est requis.|
|**OfficeTab**|Obligatoire pour étendre un onglet du ruban Office par défaut (en utilisant **PrimaryCommandSurface**). Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**. Pour plus d’informations, voir [OfficeTab](officetab.md).|
|**OfficeMenu**|Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : <br/> - **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. <br/> - **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul.|
|**Group**|Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut comporter jusqu’à six contrôles. L’attribut  **id** est requis. Il s’agit d’une chaîne contenant un maximum de 125 caractères.|
|**Label**|Obligatoire. Libellé du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. L’élément  **String** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément  **Resources**.|
|**Icon**|Obligatoire. Indique l’icône du groupe qui doit être utilisée sur les périphériques de petit facteur de forme ou lorsque les boutons sont affichés en trop grand nombre. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image**. L’élément  **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément  **Resources**. L’attribut **size** donne la taille, en pixels, de l’image. Trois tailles d’image, en pixels, sont obligatoires : 16, 32 et 80. Cinq tailles facultatives, en pixels, sont également prises en charge : 20, 24, 40, 48 et 64.|
|**Tooltip**|Facultatif. Info-bulle du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. L’élément  **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément  **Resources**.|
|**Control**|Chaque groupe requiert au moins un contrôle. Un élément **Control** peut être un **bouton** ou un **menu**. Utilisez un **menu** pour spécifier une liste déroulante de contrôles de bouton. Actuellement, seuls les boutons et les menus sont pris en charge. Pour plus d’informations, reportez-vous aux sections [Contrôles de bouton](control.md#button-control) et [Contrôles de menu](control.md#menu-dropdown-button-controls).<br/>**Remarque :**  pour faciliter les opérations de dépannage, nous vous recommandons d’ajouter un élément **Control** et les éléments enfants **Resources** associés un par un.|
|**Script**|Liens vers le fichier JavaScript avec la définition de la fonction personnalisée et le code d’inscription. Cet élément n’est pas utilisé dans l’aperçu pour les développeurs. À la place, la page HTML est responsable du chargement de tous les fichiers JavaScript.|
|**Page**|Liens vers la page HTML de vos fonctions personnalisées.|

## <a name="extension-points-for-outlook"></a>Points d’extension pour Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (peut uniquement être utilisé dans [DesktopFormFactor](desktopformfactor.md).)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
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

### <a name="events"></a>Événements

Ce point d’extension ajoute un gestionnaire d’événements pour un événement spécifié.

> [!NOTE]
> Ce type d’élément est uniquement pris en charge par Outlook sur le web dans Office 365.

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
> Ce type d’élément est uniquement pris en charge par Outlook sur le web dans Office 365.

|  Élément |  Description  |
|:-----|:-----|
|  [Label](#label) |  Spécifie l’étiquette pour le complément dans la fenêtre contextuelle.  |
|  [SourceLocation](sourcelocation.md) |  Spécifie l’URL de la fenêtre contextuelle.  |
|  [Règle](rule.md) |  Spécifie la ou les règles qui déterminent lorsqu’un complément s’active.  |

#### <a name="label"></a>Étiquette

Obligatoire. Libellé du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).

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
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```
