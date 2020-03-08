---
title: Élément Control dans le fichier manifeste
description: ''
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: ccf7c3065db13a311825498292713b619f1cd745
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/07/2020
ms.locfileid: "42562050"
---
# <a name="control-element"></a><span data-ttu-id="bfc9b-102">Élément Control</span><span class="sxs-lookup"><span data-stu-id="bfc9b-102">Control element</span></span>

<span data-ttu-id="bfc9b-p101">Définit une fonction JavaScript qui exécute une action ou lance un volet Office. Un élément **Control** peut être une option de bouton ou de menu. Au moins un élément **Control** doit être inclus dans un élément [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="bfc9b-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="bfc9b-106">Attributes</span></span>

|  <span data-ttu-id="bfc9b-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="bfc9b-107">Attribute</span></span>  |  <span data-ttu-id="bfc9b-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="bfc9b-108">Required</span></span>  |  <span data-ttu-id="bfc9b-109">Description</span><span class="sxs-lookup"><span data-stu-id="bfc9b-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="bfc9b-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-110">**xsi:type**</span></span>|<span data-ttu-id="bfc9b-111">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-111">Yes</span></span>|<span data-ttu-id="bfc9b-p102">Type de contrôle défini. Peut être `Button`, `Menu` ou `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="bfc9b-114">**id**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-114">**id**</span></span>|<span data-ttu-id="bfc9b-115">Non</span><span class="sxs-lookup"><span data-stu-id="bfc9b-115">No</span></span>|<span data-ttu-id="bfc9b-p103">ID de l’élément Control. Il doit comporter 125 caractères au maximum.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="bfc9b-118">La valeur `MobileButton` de **xsi:type** est définie dans le schéma VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="bfc9b-119">Elle s’applique uniquement aux éléments **Control** contenus dans un élément [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="bfc9b-119">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="bfc9b-120">Contrôle bouton</span><span class="sxs-lookup"><span data-stu-id="bfc9b-120">Button control</span></span>

<span data-ttu-id="bfc9b-p105">Un bouton effectue une action unique quand il est sélectionné. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle bouton doit avoir un `id` unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="bfc9b-124">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="bfc9b-124">Child elements</span></span>
|  <span data-ttu-id="bfc9b-125">Élément</span><span class="sxs-lookup"><span data-stu-id="bfc9b-125">Element</span></span> |  <span data-ttu-id="bfc9b-126">Requis</span><span class="sxs-lookup"><span data-stu-id="bfc9b-126">Required</span></span>  |  <span data-ttu-id="bfc9b-127">Description</span><span class="sxs-lookup"><span data-stu-id="bfc9b-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bfc9b-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-128">**Label**</span></span>     | <span data-ttu-id="bfc9b-129">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-129">Yes</span></span> |  <span data-ttu-id="bfc9b-130">Texte du bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-130">The text for the button.</span></span> <span data-ttu-id="bfc9b-131">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="bfc9b-131">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="bfc9b-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-132">**ToolTip**</span></span>  |<span data-ttu-id="bfc9b-133">Non</span><span class="sxs-lookup"><span data-stu-id="bfc9b-133">No</span></span>|<span data-ttu-id="bfc9b-134">Info-bulle pour le bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-134">The tooltip for the button.</span></span> <span data-ttu-id="bfc9b-135">L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="bfc9b-136">**String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="bfc9b-136">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="bfc9b-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="bfc9b-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="bfc9b-138">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-138">Yes</span></span> |  <span data-ttu-id="bfc9b-139">Info-bulle pour le bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="bfc9b-140">Icon</span><span class="sxs-lookup"><span data-stu-id="bfc9b-140">Icon</span></span>](icon.md)      | <span data-ttu-id="bfc9b-141">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-141">Yes</span></span> |  <span data-ttu-id="bfc9b-142">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="bfc9b-143">Action</span><span class="sxs-lookup"><span data-stu-id="bfc9b-143">Action</span></span>](action.md)    | <span data-ttu-id="bfc9b-144">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-144">Yes</span></span> |  <span data-ttu-id="bfc9b-145">Spécifie l’action à effectuer.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-145">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="bfc9b-146">Enabled</span><span class="sxs-lookup"><span data-stu-id="bfc9b-146">Enabled</span></span>](enabled.md)    | <span data-ttu-id="bfc9b-147">Non</span><span class="sxs-lookup"><span data-stu-id="bfc9b-147">No</span></span> |  <span data-ttu-id="bfc9b-148">Indique si le contrôle est activé au lancement du complément.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-148">Specifies whether the control is enabled when the add-in launches.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="bfc9b-149">Exemple du bouton ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="bfc9b-149">ExecuteFunction button example</span></span>

<span data-ttu-id="bfc9b-150">Dans l’exemple suivant, le bouton est désactivé au lancement du complément.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-150">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="bfc9b-151">Il peut être activé par programmation.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-151">It can be programmatically enabled.</span></span> <span data-ttu-id="bfc9b-152">Pour plus d’informations, consultez la rubrique [activer et désactiver les commandes de complément](/office/dev/add-ins/design/disable-add-in-commands).</span><span class="sxs-lookup"><span data-stu-id="bfc9b-152">For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).</span></span>

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
  <Enabled>false</Enabled>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="bfc9b-153">Exemple du bouton ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="bfc9b-153">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="bfc9b-154">Contrôles de menu (bouton déroulant)</span><span class="sxs-lookup"><span data-stu-id="bfc9b-154">Menu (dropdown button) controls</span></span>

<span data-ttu-id="bfc9b-p109">Un menu définit une liste statique d’options. Chaque option de menu exécute une fonction ou affiche un volet Office. Les sous-menus ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p109">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="bfc9b-158">Lorsqu’il est utilisé avec un **point d’extension** **PrimaryCommandSurface** ou [ContextMenu](extensionpoint.md), le contrôle de menu définit les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="bfc9b-158">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="bfc9b-159">une option de menu de niveau racine.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-159">A root-level menu item.</span></span>

- <span data-ttu-id="bfc9b-160">une liste de sous-menus.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-160">A list of submenu items.</span></span>

<span data-ttu-id="bfc9b-p110">Lorsqu’il est utilisé avec **PrimaryCommandSurface**, l’option de menu de niveau racine s’affiche sous la forme d’un bouton dans le ruban. Lorsque le bouton est sélectionné, le sous-menu s’affiche sous la forme d’une liste déroulante. Lorsqu’il est utilisé avec **ContextMenu**, un élément de menu avec un sous-menu est inséré dans le menu contextuel. Dans les deux cas, les éléments individuels du sous-menu peuvent exécuter une fonction JavaScript ou afficher un volet de tâches. Un seul niveau de sous-menus est pris en charge pour l’instant.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p110">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="bfc9b-p111">L’exemple suivant montre comment définir un élément de menu avec deux éléments de sous-menu. Le premier élément de sous-menu affiche un volet Office et le deuxième élément de sous-menu exécute une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p111">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="bfc9b-168">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="bfc9b-168">Child elements</span></span>

|  <span data-ttu-id="bfc9b-169">Élément</span><span class="sxs-lookup"><span data-stu-id="bfc9b-169">Element</span></span> |  <span data-ttu-id="bfc9b-170">Requis</span><span class="sxs-lookup"><span data-stu-id="bfc9b-170">Required</span></span>  |  <span data-ttu-id="bfc9b-171">Description</span><span class="sxs-lookup"><span data-stu-id="bfc9b-171">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bfc9b-172">**Label**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-172">**Label**</span></span>     | <span data-ttu-id="bfc9b-173">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-173">Yes</span></span> |  <span data-ttu-id="bfc9b-174">Texte du bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-174">The text for the button.</span></span> <span data-ttu-id="bfc9b-175">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="bfc9b-175">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="bfc9b-176">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-176">**ToolTip**</span></span>  |<span data-ttu-id="bfc9b-177">Non</span><span class="sxs-lookup"><span data-stu-id="bfc9b-177">No</span></span>|<span data-ttu-id="bfc9b-178">Info-bulle pour le bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-178">The tooltip for the button.</span></span> <span data-ttu-id="bfc9b-179">L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-179">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="bfc9b-180">**String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="bfc9b-180">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="bfc9b-181">Supertip</span><span class="sxs-lookup"><span data-stu-id="bfc9b-181">Supertip</span></span>](supertip.md)  | <span data-ttu-id="bfc9b-182">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-182">Yes</span></span> |  <span data-ttu-id="bfc9b-183">Info-bulle pour ce bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-183">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="bfc9b-184">Icon</span><span class="sxs-lookup"><span data-stu-id="bfc9b-184">Icon</span></span>](icon.md)      | <span data-ttu-id="bfc9b-185">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-185">Yes</span></span> |  <span data-ttu-id="bfc9b-186">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-186">An image for the button.</span></span>         |
|  <span data-ttu-id="bfc9b-187">**Éléments**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-187">**Items**</span></span>     | <span data-ttu-id="bfc9b-188">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-188">Yes</span></span> |  <span data-ttu-id="bfc9b-189">Collection de boutons à afficher dans le menu.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-189">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="bfc9b-190">Contient les éléments **Élément** pour chaque élément de sous-menu.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-190">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="bfc9b-191">Chaque élément **Item** contient les éléments enfants du [contrôle Button](#button-control).</span><span class="sxs-lookup"><span data-stu-id="bfc9b-191">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="bfc9b-192">Exemples de contrôle de menu</span><span class="sxs-lookup"><span data-stu-id="bfc9b-192">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="bfc9b-193">Contrôle MobileButton</span><span class="sxs-lookup"><span data-stu-id="bfc9b-193">MobileButton control</span></span>

<span data-ttu-id="bfc9b-p115">Un bouton mobile effectue une action unique lorsque l’utilisateur le sélectionne. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle de bouton mobile doit avoir un `id` unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p115">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="bfc9b-p116">La valeur `MobileButton` de **xsi:type** est définie dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p116">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="bfc9b-199">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="bfc9b-199">Child elements</span></span>
|  <span data-ttu-id="bfc9b-200">Élément</span><span class="sxs-lookup"><span data-stu-id="bfc9b-200">Element</span></span> |  <span data-ttu-id="bfc9b-201">Requis</span><span class="sxs-lookup"><span data-stu-id="bfc9b-201">Required</span></span>  |  <span data-ttu-id="bfc9b-202">Description</span><span class="sxs-lookup"><span data-stu-id="bfc9b-202">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bfc9b-203">**Label**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-203">**Label**</span></span>     | <span data-ttu-id="bfc9b-204">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-204">Yes</span></span> |  <span data-ttu-id="bfc9b-205">Texte du bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-205">The text for the button.</span></span> <span data-ttu-id="bfc9b-206">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="bfc9b-206">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="bfc9b-207">Icon</span><span class="sxs-lookup"><span data-stu-id="bfc9b-207">Icon</span></span>](icon.md)      | <span data-ttu-id="bfc9b-208">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-208">Yes</span></span> |  <span data-ttu-id="bfc9b-209">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-209">An image for the button.</span></span>         |
|  [<span data-ttu-id="bfc9b-210">Action</span><span class="sxs-lookup"><span data-stu-id="bfc9b-210">Action</span></span>](action.md)    | <span data-ttu-id="bfc9b-211">Oui</span><span class="sxs-lookup"><span data-stu-id="bfc9b-211">Yes</span></span> |  <span data-ttu-id="bfc9b-212">Spécifie l’action à effectuer.</span><span class="sxs-lookup"><span data-stu-id="bfc9b-212">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="bfc9b-213">Exemple de bouton mobile ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="bfc9b-213">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="bfc9b-214">Exemple de bouton mobile ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="bfc9b-214">ShowTaskpane mobile button example</span></span>

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
