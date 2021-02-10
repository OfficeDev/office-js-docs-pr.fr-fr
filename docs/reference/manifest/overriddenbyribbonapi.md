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
# <a name="overriddenbyribbonapi-element"></a><span data-ttu-id="d7300-103">Élément OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="d7300-103">OverriddenByRibbonApi element</span></span>

<span data-ttu-id="d7300-104">Spécifie si un contrôle [CustomTab,](customtab.md) [Group,](group.md) [Button,](control.md#button-control) [Menu](control.md#menu-dropdown-button-controls) ou menu sera masqué sur les combinaisons d’applications et de plateformes qui prendre en charge l’API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) qui installe des onglets contextuels personnalisés sur le ruban.</span><span class="sxs-lookup"><span data-stu-id="d7300-104">Specifies whether a [CustomTab](customtab.md), [Group](group.md), [Button](control.md#button-control) control, [Menu](control.md#menu-dropdown-button-controls) control, or menu item will be hidden on application and platform combinations that support the API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) that installs custom contextual tabs on the ribbon.</span></span>

<span data-ttu-id="d7300-105">S’il est omis, la valeur par défaut est `false` .</span><span class="sxs-lookup"><span data-stu-id="d7300-105">If it is omitted, the default is `false`.</span></span> <span data-ttu-id="d7300-106">S’il est utilisé, il doit être le *premier* élément enfant de son élément parent.</span><span class="sxs-lookup"><span data-stu-id="d7300-106">If it is used, it must be the *first* child element of its parent element.</span></span>

> [!NOTE]
> <span data-ttu-id="d7300-107">Pour une compréhension complète de cet élément, lisez Implémenter une autre expérience d’interface utilisateur lorsque les [onglets contextuels personnalisés ne sont pas pris en charge.](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)</span><span class="sxs-lookup"><span data-stu-id="d7300-107">For a full understanding of this element, please read [Implement an alternate UI experience when custom contextual tabs are not supported](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

<span data-ttu-id="d7300-108">L’objectif de cet élément est de créer une expérience de retour dans un add-in qui implémente des onglets contextuels personnalisés lorsque le module est en cours d’exécution sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés.</span><span class="sxs-lookup"><span data-stu-id="d7300-108">The purpose of this element is to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> <span data-ttu-id="d7300-109">La stratégie essentielle consiste à dupliquer tout ou partie des groupes et contrôles de votre  onglet contextuel personnalisé sur un ou plusieurs onglets principaux personnalisés (c’est-à-dire, des onglets personnalisés nontexte).</span><span class="sxs-lookup"><span data-stu-id="d7300-109">The essential strategy is that you duplicate some or all of the groups and controls from your custom contextual tab onto one or more custom core tabs (that is, *noncontextual* custom tabs).</span></span> <span data-ttu-id="d7300-110">Ensuite, pour vous assurer que ces groupes et  contrôles apparaissent lorsque les onglets contextuels personnalisés ne sont pas pris en charge, mais n’apparaissent pas lorsque les *onglets* contextuels personnalisés sont pris en charge, vous ajoutez en tant que premier élément enfant des éléments `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **CustomTab,** **Group,** **Control** ou Menu **Item.**</span><span class="sxs-lookup"><span data-stu-id="d7300-110">Then, to ensure that these groups and controls appear when custom contextual tabs are *not* supported, but do not appear when custom contextual tabs *are* supported, you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the **CustomTab**, **Group**, **Control**, or menu **Item** elements.</span></span> <span data-ttu-id="d7300-111">L’effet de cette utilisation est le suivant :</span><span class="sxs-lookup"><span data-stu-id="d7300-111">The effect of doing so is the following:</span></span>

- <span data-ttu-id="d7300-112">Si le add-in s’exécute sur une application et une plateforme qui prend en charge les onglets contextuels personnalisés, les onglets, groupes et contrôles dupliqués n’apparaissent pas sur le ruban.</span><span class="sxs-lookup"><span data-stu-id="d7300-112">If the add-in runs on an application and platform that support custom contextual tabs, then the duplicated tabs, groups, and controls won't appear on the ribbon.</span></span> <span data-ttu-id="d7300-113">Au lieu de cela, l’onglet contextuel personnalisé est installé lorsque le add-in appelle la `requestCreateControls` méthode.</span><span class="sxs-lookup"><span data-stu-id="d7300-113">Instead, the custom contextual tab will be installed when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="d7300-114">Si le add-in *s’exécute* sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés, les onglets, groupes et contrôles dupliqués apparaissent sur le ruban.</span><span class="sxs-lookup"><span data-stu-id="d7300-114">If the add-in runs on an application or platform that *doesn't* support custom contextual tabs, then the duplicated tabs, groups, and controls will appear on the ribbon.</span></span>

## <a name="examples"></a><span data-ttu-id="d7300-115">Exemples</span><span class="sxs-lookup"><span data-stu-id="d7300-115">Examples</span></span>

### <a name="overriding-an-entire-tab"></a><span data-ttu-id="d7300-116">Remplacement d’un onglet entier</span><span class="sxs-lookup"><span data-stu-id="d7300-116">Overriding an entire tab</span></span>

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

### <a name="overriding-a-group"></a><span data-ttu-id="d7300-117">Remplacement d’un groupe</span><span class="sxs-lookup"><span data-stu-id="d7300-117">Overriding a group</span></span>

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

### <a name="overriding-a-control"></a><span data-ttu-id="d7300-118">Remplacement d’un contrôle</span><span class="sxs-lookup"><span data-stu-id="d7300-118">Overriding a control</span></span>

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

### <a name="overriding-a-menu-item"></a><span data-ttu-id="d7300-119">Remplacement d’un élément de menu</span><span class="sxs-lookup"><span data-stu-id="d7300-119">Overriding a menu item</span></span>


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
