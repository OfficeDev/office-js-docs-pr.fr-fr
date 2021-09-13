---
title: Élément OverriddenByRibbonApi dans le fichier manifeste
description: Découvrez comment spécifier qu’un onglet, un groupe, un contrôle ou un élément de menu personnalisé ne doit pas apparaître lorsqu’il fait également partie d’un onglet contextuel personnalisé.
ms.date: 09/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 35893bba5c00d8b6d63f02cc12ac6902197ab0d8
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153700"
---
# <a name="overriddenbyribbonapi-element"></a>Élément OverriddenByRibbonApi

Spécifie si un [](control.md#button-control) contrôle [groupe,](group.md) [bouton,](control.md#menu-dropdown-button-controls) menu ou élément de menu sera masqué sur les combinaisons d’application et de plateforme qui prendre en charge l’API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)) qui installe des onglets contextuels personnalisés sur le ruban.

S’il est omis, la valeur par défaut est `false` . S’il est utilisé, il doit être le *premier* élément enfant de son élément parent.

> [!NOTE]
> Pour une compréhension complète de cet élément, lisez Implémenter une autre expérience d’interface utilisateur lorsque les [onglets contextuels personnalisés ne sont pas pris en charge.](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)

L’objectif de cet élément est de créer une expérience de retour dans un add-in qui implémente des onglets contextuels personnalisés lorsque le module est en cours d’exécution sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés. La stratégie essentielle consiste à dupliquer tout ou partie des groupes et contrôles de votre  onglet contextuel personnalisé sur un ou plusieurs onglets principaux personnalisés (c’est-à-dire, des onglets personnalisés nontexte). Ensuite, pour vous assurer que ces groupes et  contrôles apparaissent lorsque les onglets contextuels personnalisés ne sont pas pris en charge, mais n’apparaissent pas lorsque les *onglets* contextuels personnalisés sont pris en charge, vous ajoutez en tant que premier élément enfant des éléments de groupe, de contrôle ou d’élément de `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` menu.    L’effet de cette utilisation est le suivant :

- Si le add-in s’exécute sur une application et une plateforme qui prend en charge les onglets contextuels personnalisés, les groupes et contrôles dupliqués n’apparaissent pas sur le ruban. Au lieu de cela, l’onglet contextuel personnalisé est installé lorsque le add-in appelle la `requestCreateControls` méthode.
- Si le add-in *s’exécute* sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés, les groupes et contrôles dupliqués apparaissent sur le ruban.

## <a name="examples"></a>Exemples

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
