---
title: Élément Control de type Menu dans le fichier manifeste
description: Définit un menu dont les éléments peuvent exécuter des actions ou lancer des volets Des tâches.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7287b8e2cdf2378140ef50a41306820a0fd4002f
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467894"
---
# <a name="control-element-of-type-menu"></a>Élément Control de type Menu

Un menu définit une liste d’options. Chaque option de menu exécute une fonction ou affiche un volet Office.

> [!NOTE]
> Cet article suppose une connaissance de [l’article](control.md) de référence du contrôle de base qui contient des informations importantes sur les attributs de l’élément.

Le contrôle de menu définit :

- Contrôle de menu de niveau racine.
- Liste d’éléments de menu.

Lorsqu’il est utilisé avec le [point d’extension](extensionpoint.md) **PrimaryCommandSurface**, l’élément de menu racine s’affiche sous la forme d’un bouton sur le ruban. Lorsque le bouton est sélectionné, le menu s’affiche en tant que liste déroulante. Les sous-menus ne sont pas pris en charge.

Lorsqu’il est utilisé avec le [point d’extension](extensionpoint.md) **ContextMenu**, un élément de menu racine s’affiche dans le menu contextu. Lorsque l’élément racine est sélectionné, les éléments de menu s’affichent en tant que sous-menu. Aucun des éléments ne peut lui-même être un sous-menu, car un seul niveau de sous-menus est pris en charge.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)     | Oui |  Texte du menu. |
|  **ToolTip**    |Non|Info-bulle du menu. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un **élément String**. **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).|
|  [Supertip](supertip.md)  | Oui |  Super-bulle de ce menu.    |
|  [Icon](icon.md)      | Oui |  Image du menu.         |
|  **Éléments**     | Oui |  Collection d’éléments à afficher dans le menu. Contient **l’élément Item** pour chaque élément. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Non |  Spécifie si le menu doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés. S’il est utilisé, il doit s’agit du *premier* élément enfant. |

### <a name="label"></a>Étiquette

Spécifie le texte du nom du menu au moyen de son seul attribut, **resid**, qui ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **String** dans l’enfant **ShortStrings** de l’élément [Resources](resources.md) .

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) lorsque le parent **VersionOverrides** est de type Taskpane 1.0.
- [Boîte aux lettres 1.3 lorsque](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) le parent **VersionOverrides** est de type Mail 1.0.
- [Boîte aux lettres 1.5 lorsque](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) le parent **VersionOverrides** est de type Mail 1.1.

## <a name="examples"></a>Exemples

Dans l’exemple suivant, le menu comprend deux éléments. Le premier affiche un volet Des tâches. La seconde exécute une fonction. Le menu a été configuré pour ne  pas être visible lorsque le module est en cours d’exécution sur une plateforme qui prend en charge les onglets contextuels. Pour plus d’informations, lisez Implémenter une [autre expérience d’interface utilisateur lorsque les onglets contextuels personnalisés ne sont pas pris en charge](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

```xml
<Control xsi:type="Menu" id="Contoso.TestMenu2">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
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
    <Item id="ShowMainTaskPane">
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
    <Item id="GetData">
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
        <FunctionName>getData</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

Dans l’exemple suivant, le deuxième élément du menu est configuré pour ne  pas être visible lorsque le module est en cours d’exécution sur une plateforme qui prend en charge les onglets contextuels. Pour plus d’informations, lisez Implémenter une [autre expérience d’interface utilisateur lorsque les onglets contextuels personnalisés ne sont pas pris en charge](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

```xml
<Control xsi:type="Menu" id="Contoso.msgReadMenuButton">
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
    <Item id="ShowMainTaskPane">
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
    <Item id="msgReadMenuItem1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
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
  </Items>
</Control>
```
