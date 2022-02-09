---
title: Élément OverriddenByRibbonApi dans le fichier manifeste
description: Découvrez comment spécifier qu’un onglet, un groupe, un contrôle ou un élément de menu personnalisé ne doit pas apparaître lorsqu’il fait également partie d’un onglet contextuel personnalisé.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 48977691ee4bf2ccd71bc146647dae452ce9e2fc
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467686"
---
# <a name="overriddenbyribbonapi-element"></a>Élément OverriddenByRibbonApi

Spécifie si un contrôle [groupe, bouton](group.md)[, menu](control-button.md) ou élément de menu sera masqué sur les combinaisons d’applications et de plateformes qui prendre en charge l’API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1))) qui installe des onglets contextuels personnalisés sur le ruban. [](control-menu.md)

**Type de complément :** volet Office

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Taskpane 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [Ruban 1.2](../requirement-sets/add-in-commands-requirement-sets.md) (requis pour Excel, PowerPoint et Word.)

Si cet élément est omis, la valeur par défaut est `false`. S’il est utilisé, il doit être le *premier* élément enfant de son élément parent.

> [!NOTE]
> Pour une compréhension complète de cet élément, lisez Implémenter une autre expérience d’interface utilisateur lorsque les [onglets contextuels personnalisés ne sont pas pris en charge](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

L’objectif de cet élément est de créer une expérience de retour dans un application qui implémente des onglets contextuels personnalisés lorsque le module est en cours d’exécution sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés. La stratégie essentielle consiste à dupliquer tout ou partie des groupes et contrôles de votre onglet contextuel personnalisé sur un ou plusieurs onglets principaux personnalisés (c’est-à-dire, des onglets personnalisés *nontexte* ). Ensuite, pour vous assurer que ces groupes et contrôles apparaissent lorsque les onglets contextuels personnalisés ne sont pas pris en charge, mais qu’ils n’apparaissent pas lorsque  les *onglets* contextuels personnalisés sont pris en charge, `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` vous ajoutez en tant que premier élément enfant des éléments **Group**, **Control** ou **Menu Item**. L’effet de cette utilisation est le suivant :

- Si le add-in s’exécute sur une application et une plateforme qui prend en charge les onglets contextuels personnalisés, les groupes et contrôles dupliqués n’apparaissent pas sur le ruban. Au lieu de cela, l’onglet contextuel personnalisé est installé lorsque le add-in appelle la `requestCreateControls` méthode.
- Si le add-in *s’exécute* sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés, les groupes et contrôles dupliqués apparaissent sur le ruban.

## <a name="examples"></a>Exemples

### <a name="overriding-a-group"></a>Remplacement d’un groupe

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.CustomTab1.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="Contoso.MyButton1">
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
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.CustomTab2.group2">
      <Control  xsi:type="Button" id="Contoso.MyButton2">
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
  <CustomTab id="Contoso.TabCustom3">
    <Group id="Contoso.CustomTab3.group3">
      <Control  xsi:type="Menu" id="Contoso.MyMenu">
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
