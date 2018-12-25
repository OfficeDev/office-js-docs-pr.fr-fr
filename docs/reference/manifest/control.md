---
title: Élément Control dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: e5d8574e322c21e768fb9f66fe9bbb0c12a34ed4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433935"
---
# <a name="control-element"></a><span data-ttu-id="24290-102">Élément Control</span><span class="sxs-lookup"><span data-stu-id="24290-102">Control element</span></span>

<span data-ttu-id="24290-p101">Définit une fonction JavaScript qui exécute une action ou lance un volet Office. Un élément **Control** peut être une option de bouton ou de menu. Au moins un élément **Control** doit être inclus dans un élément [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="24290-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="24290-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="24290-106">Attributes</span></span>

|  <span data-ttu-id="24290-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="24290-107">Attribute</span></span>  |  <span data-ttu-id="24290-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="24290-108">Required</span></span>  |  <span data-ttu-id="24290-109">Description</span><span class="sxs-lookup"><span data-stu-id="24290-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="24290-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="24290-110">**xsi:type**</span></span>|<span data-ttu-id="24290-111">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-111">Yes</span></span>|<span data-ttu-id="24290-p102">Type de contrôle défini. Peut être `Button`, `Menu` ou `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="24290-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="24290-114">**id**</span><span class="sxs-lookup"><span data-stu-id="24290-114">**id**</span></span>|<span data-ttu-id="24290-115">Non</span><span class="sxs-lookup"><span data-stu-id="24290-115">No</span></span>|<span data-ttu-id="24290-p103">ID de l’élément Control. Il doit comporter 125 caractères au maximum.</span><span class="sxs-lookup"><span data-stu-id="24290-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="24290-118">La valeur `MobileButton` de **xsi:type** est définie dans le schéma VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="24290-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing VersionOverrides element must have an  attribute value of .</span></span> <span data-ttu-id="24290-119">Elle s’applique uniquement aux éléments **Control** contenus dans un élément [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="24290-119">Note: The  value for xsi:type is defined in VersionOverrides schema 1.1. It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="24290-120">Contrôle bouton</span><span class="sxs-lookup"><span data-stu-id="24290-120">Button control</span></span>

<span data-ttu-id="24290-p105">Un bouton effectue une action unique quand il est sélectionné. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle bouton doit avoir un `id` unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="24290-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="24290-124">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="24290-124">Child elements</span></span>
|  <span data-ttu-id="24290-125">Élément</span><span class="sxs-lookup"><span data-stu-id="24290-125">Element</span></span> |  <span data-ttu-id="24290-126">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="24290-126">Required</span></span>  |  <span data-ttu-id="24290-127">Description</span><span class="sxs-lookup"><span data-stu-id="24290-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="24290-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="24290-128">**Label**</span></span>     | <span data-ttu-id="24290-129">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-129">Yes</span></span> |  <span data-ttu-id="24290-p106">Texte du bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="24290-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="24290-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="24290-132">**ToolTip**</span></span>  |<span data-ttu-id="24290-133">Non</span><span class="sxs-lookup"><span data-stu-id="24290-133">No</span></span>|<span data-ttu-id="24290-p107">Info-bulle pour le bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="24290-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="24290-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="24290-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="24290-138">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-138">Yes</span></span> |  <span data-ttu-id="24290-139">Info-bulle pour le bouton.</span><span class="sxs-lookup"><span data-stu-id="24290-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="24290-140">Icon</span><span class="sxs-lookup"><span data-stu-id="24290-140">Icon</span></span>](icon.md)      | <span data-ttu-id="24290-141">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-141">Yes</span></span> |  <span data-ttu-id="24290-142">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="24290-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="24290-143">Action</span><span class="sxs-lookup"><span data-stu-id="24290-143">Action</span></span>](action.md)    | <span data-ttu-id="24290-144">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-144">Yes</span></span> |  <span data-ttu-id="24290-145">Spécifie l’action à effectuer.</span><span class="sxs-lookup"><span data-stu-id="24290-145">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="24290-146">Exemple du bouton ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="24290-146">ExecuteFunction button example</span></span>

```xml
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
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="24290-147">Exemple du bouton ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="24290-147">ShowTaskpane button example</span></span>

```xml
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
```

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="24290-148">Contrôles de menu (bouton déroulant)</span><span class="sxs-lookup"><span data-stu-id="24290-148">Menu (dropdown button) controls</span></span>

<span data-ttu-id="24290-p108">Un menu définit une liste statique d’options. Chaque option de menu exécute une fonction ou affiche un volet Office. Les sous-menus ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="24290-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="24290-152">Lorsqu’il est utilisé avec un [point d’extension](extensionpoint.md) **PrimaryCommandSurface** ou **ContextMenu**, le contrôle de menu définit les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="24290-152">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="24290-153">une option de menu de niveau racine.</span><span class="sxs-lookup"><span data-stu-id="24290-153">A root-level menu item.</span></span>

- <span data-ttu-id="24290-154">une liste de sous-menus.</span><span class="sxs-lookup"><span data-stu-id="24290-154">A list of submenu items.</span></span>

<span data-ttu-id="24290-p109">Lorsqu’il est utilisé avec  **PrimaryCommandSurface**, l’élément de menu racine apparaît sous forme de bouton sur le ruban. Lorsque ce bouton est sélectionné, ce menu s’affiche comme une liste déroulante. Lorsqu’il est utilisé avec  **ContextMenu**, une option de menu comportant un sous-menu est inséré dans le menu contextuel. Dans les deux cas, les éléments de sous-menu individuels peuvent soit exécuter une fonction JavaScript, soit afficher un volet de tâches. Un seul niveau de sous-menus est actuellement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="24290-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="24290-p110">L’exemple suivant montre comment définir un élément de menu avec deux éléments de sous-menu. Le premier élément de sous-menu affiche un volet Office et le deuxième élément de sous-menu exécute une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="24290-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
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
    <Item id="showGallery2">
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
    <Item id="showGallery3">
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
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a><span data-ttu-id="24290-162">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="24290-162">Child elements</span></span>

|  <span data-ttu-id="24290-163">Élément</span><span class="sxs-lookup"><span data-stu-id="24290-163">Element</span></span> |  <span data-ttu-id="24290-164">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="24290-164">Required</span></span>  |  <span data-ttu-id="24290-165">Description</span><span class="sxs-lookup"><span data-stu-id="24290-165">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="24290-166">**Label**</span><span class="sxs-lookup"><span data-stu-id="24290-166">**Label**</span></span>     | <span data-ttu-id="24290-167">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-167">Yes</span></span> |  <span data-ttu-id="24290-p111">Texte du bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="24290-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="24290-170">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="24290-170">**ToolTip**</span></span>  |<span data-ttu-id="24290-171">Non</span><span class="sxs-lookup"><span data-stu-id="24290-171">No</span></span>|<span data-ttu-id="24290-p112">Info-bulle pour le bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="24290-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="24290-175">Supertip</span><span class="sxs-lookup"><span data-stu-id="24290-175">Supertip</span></span>](supertip.md)  | <span data-ttu-id="24290-176">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-176">Yes</span></span> |  <span data-ttu-id="24290-177">Info-bulle pour ce bouton.</span><span class="sxs-lookup"><span data-stu-id="24290-177">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="24290-178">Icon</span><span class="sxs-lookup"><span data-stu-id="24290-178">Icon</span></span>](icon.md)      | <span data-ttu-id="24290-179">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-179">Yes</span></span> |  <span data-ttu-id="24290-180">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="24290-180">An image for the button.</span></span>         |
|  <span data-ttu-id="24290-181">**Items**</span><span class="sxs-lookup"><span data-stu-id="24290-181">**Items**</span></span>     | <span data-ttu-id="24290-182">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-182">Yes</span></span> |  <span data-ttu-id="24290-p113">Ensemble de boutons à afficher dans le menu Contient les éléments **Item** pour chaque élément de sous-menu. Chaque élément **Item** contient les éléments enfants du [contrôle de bouton](#button-control).</span><span class="sxs-lookup"><span data-stu-id="24290-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="24290-186">Exemples de contrôle de menu</span><span class="sxs-lookup"><span data-stu-id="24290-186">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
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
    <Item id="showGallery2">
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
    <Item id="showGallery3">
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
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
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
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a><span data-ttu-id="24290-187">Contrôle MobileButton</span><span class="sxs-lookup"><span data-stu-id="24290-187">MobileButton control</span></span>

<span data-ttu-id="24290-p114">Un bouton mobile effectue une action unique lorsque l’utilisateur le sélectionne. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle de bouton mobile doit avoir un `id` unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="24290-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="24290-p115">La valeur `MobileButton` de **xsi:type** est définie dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="24290-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="24290-193">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="24290-193">Child elements</span></span>
|  <span data-ttu-id="24290-194">Élément</span><span class="sxs-lookup"><span data-stu-id="24290-194">Element</span></span> |  <span data-ttu-id="24290-195">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="24290-195">Required</span></span>  |  <span data-ttu-id="24290-196">Description</span><span class="sxs-lookup"><span data-stu-id="24290-196">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="24290-197">**Label**</span><span class="sxs-lookup"><span data-stu-id="24290-197">**Label**</span></span>     | <span data-ttu-id="24290-198">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-198">Yes</span></span> |  <span data-ttu-id="24290-p116">Texte du bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="24290-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="24290-201">Icon</span><span class="sxs-lookup"><span data-stu-id="24290-201">Icon</span></span>](icon.md)      | <span data-ttu-id="24290-202">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-202">Yes</span></span> |  <span data-ttu-id="24290-203">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="24290-203">An image for the button.</span></span>         |
|  [<span data-ttu-id="24290-204">Action</span><span class="sxs-lookup"><span data-stu-id="24290-204">Action</span></span>](action.md)    | <span data-ttu-id="24290-205">Oui</span><span class="sxs-lookup"><span data-stu-id="24290-205">Yes</span></span> |  <span data-ttu-id="24290-206">Spécifie l’action à effectuer.</span><span class="sxs-lookup"><span data-stu-id="24290-206">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="24290-207">Exemple de bouton mobile ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="24290-207">ExecuteFunction mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="24290-208">Exemple de bouton mobile ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="24290-208">ShowTaskpane mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```