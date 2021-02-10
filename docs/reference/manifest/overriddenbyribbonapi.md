---
title: Élément OverriddenByRibbonApi dans le fichier manifeste
description: Découvrez comment spécifier qu’un onglet, un groupe, un contrôle ou un élément de menu personnalisé ne doit pas apparaître lorsqu’il fait également partie d’un onglet contextuel personnalisé.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 62aa484057221f9cd7f41af9c8b9210cdb5b3376
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173997"
---
# <a name="overriddenbyribbonapi-element"></a>Élément OverriddenByRibbonApi

Spécifie si un contrôle [CustomTab,](customtab.md) [Group,](group.md) [Button,](control.md#button-control) [Menu](control.md#menu-dropdown-button-controls) ou menu sera masqué sur les combinaisons d’applications et de plateformes qui prendre en charge l’API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) qui installe des onglets contextuels personnalisés sur le ruban.

S’il est omis, la valeur par défaut est `false` . S’il est utilisé, il doit être le *premier* élément enfant de son élément parent.

> [!NOTE]
> Pour une compréhension complète de cet élément, lisez Implémenter une autre expérience d’interface utilisateur lorsque les [onglets contextuels personnalisés ne sont pas pris en charge.](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)

L’objectif de cet élément est de créer une expérience de retour dans un add-in qui implémente des onglets contextuels personnalisés lorsque le module est en cours d’exécution sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés. La stratégie essentielle consiste à dupliquer tout ou partie des groupes et contrôles de votre  onglet contextuel personnalisé sur un ou plusieurs onglets principaux personnalisés (c’est-à-dire, des onglets personnalisés nontexte). Ensuite, pour vous assurer que ces groupes et  contrôles apparaissent lorsque les onglets contextuels personnalisés ne sont pas pris en charge, mais n’apparaissent pas lorsque les *onglets* contextuels personnalisés sont pris en charge, vous ajoutez en tant que premier élément enfant des éléments `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **CustomTab,** **Group,** **Control** ou Menu **Item.** L’effet de cette utilisation est le suivant :

- Si le add-in s’exécute sur une application et une plateforme qui prend en charge les onglets contextuels personnalisés, les onglets, groupes et contrôles dupliqués n’apparaissent pas sur le ruban. Au lieu de cela, l’onglet contextuel personnalisé est installé lorsque le add-in appelle la `requestCreateControls` méthode.
- Si le add-in *s’exécute* sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés, les onglets, groupes et contrôles dupliqués apparaissent sur le ruban.

## <a name="examples"></a>Exemples

### <a name="overriding-an-entire-tab"></a>Remplacement d’un onglet entier

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-group"></a>Remplacement d’un groupe

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-control"></a>Remplacement d’un contrôle

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
        <!-- Other child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-menu-item"></a>Remplacement d’un élément de menu


```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Menu" id="MyMenu">
        <!-- Other child elements omitted. -->
        <Items>
          <Item id="showGallery">
            <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
            <!-- Other child elements omitted. -->
          </Item>
        </Items>
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
