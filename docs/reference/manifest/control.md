---
title: Élément Control dans le fichier manifeste
description: Définit une fonction JavaScript qui exécute une action ou lance un volet Office.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 0d13fc4fd30f32b4998d385058cff941d3f72353
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720672"
---
# <a name="control-element"></a><span data-ttu-id="470c7-103">Élément Control</span><span class="sxs-lookup"><span data-stu-id="470c7-103">Control element</span></span>

<span data-ttu-id="470c7-p101">Définit une fonction JavaScript qui exécute une action ou lance un volet Office. Un élément **Control** peut être une option de bouton ou de menu. Au moins un élément **Control** doit être inclus dans un élément [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="470c7-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="470c7-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="470c7-107">Attributes</span></span>

|  <span data-ttu-id="470c7-108">Attribut</span><span class="sxs-lookup"><span data-stu-id="470c7-108">Attribute</span></span>  |  <span data-ttu-id="470c7-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="470c7-109">Required</span></span>  |  <span data-ttu-id="470c7-110">Description</span><span class="sxs-lookup"><span data-stu-id="470c7-110">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="470c7-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="470c7-111">**xsi:type**</span></span>|<span data-ttu-id="470c7-112">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-112">Yes</span></span>|<span data-ttu-id="470c7-p102">Type de contrôle défini. Peut être `Button`, `Menu` ou `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="470c7-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="470c7-115">**id**</span><span class="sxs-lookup"><span data-stu-id="470c7-115">**id**</span></span>|<span data-ttu-id="470c7-116">Non</span><span class="sxs-lookup"><span data-stu-id="470c7-116">No</span></span>|<span data-ttu-id="470c7-p103">ID de l’élément Control. Il doit comporter 125 caractères au maximum.</span><span class="sxs-lookup"><span data-stu-id="470c7-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="470c7-119">La valeur `MobileButton` de **xsi:type** est définie dans le schéma VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="470c7-119">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="470c7-120">Elle s’applique uniquement aux éléments **Control** contenus dans un élément [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="470c7-120">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="470c7-121">Contrôle bouton</span><span class="sxs-lookup"><span data-stu-id="470c7-121">Button control</span></span>

<span data-ttu-id="470c7-p105">Un bouton effectue une action unique quand il est sélectionné. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle bouton doit avoir un `id` unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="470c7-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="470c7-125">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="470c7-125">Child elements</span></span>
|  <span data-ttu-id="470c7-126">Élément</span><span class="sxs-lookup"><span data-stu-id="470c7-126">Element</span></span> |  <span data-ttu-id="470c7-127">Requis</span><span class="sxs-lookup"><span data-stu-id="470c7-127">Required</span></span>  |  <span data-ttu-id="470c7-128">Description</span><span class="sxs-lookup"><span data-stu-id="470c7-128">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="470c7-129">**Label**</span><span class="sxs-lookup"><span data-stu-id="470c7-129">**Label**</span></span>     | <span data-ttu-id="470c7-130">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-130">Yes</span></span> |  <span data-ttu-id="470c7-131">Texte du bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-131">The text for the button.</span></span> <span data-ttu-id="470c7-132">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="470c7-132">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="470c7-133">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="470c7-133">**ToolTip**</span></span>    |<span data-ttu-id="470c7-134">Non</span><span class="sxs-lookup"><span data-stu-id="470c7-134">No</span></span>|<span data-ttu-id="470c7-135">Info-bulle pour le bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-135">The tooltip for the button.</span></span> <span data-ttu-id="470c7-136">L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**.</span><span class="sxs-lookup"><span data-stu-id="470c7-136">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="470c7-137">**String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="470c7-137">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="470c7-138">Supertip</span><span class="sxs-lookup"><span data-stu-id="470c7-138">Supertip</span></span>](supertip.md)  | <span data-ttu-id="470c7-139">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-139">Yes</span></span> |  <span data-ttu-id="470c7-140">Info-bulle pour le bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-140">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="470c7-141">Icon</span><span class="sxs-lookup"><span data-stu-id="470c7-141">Icon</span></span>](icon.md)      | <span data-ttu-id="470c7-142">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-142">Yes</span></span> |  <span data-ttu-id="470c7-143">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-143">An image for the button.</span></span>         |
|  [<span data-ttu-id="470c7-144">Action</span><span class="sxs-lookup"><span data-stu-id="470c7-144">Action</span></span>](action.md)    | <span data-ttu-id="470c7-145">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-145">Yes</span></span> |  <span data-ttu-id="470c7-146">Spécifie l’action à effectuer.</span><span class="sxs-lookup"><span data-stu-id="470c7-146">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="470c7-147">Enabled</span><span class="sxs-lookup"><span data-stu-id="470c7-147">Enabled</span></span>](enabled.md)    | <span data-ttu-id="470c7-148">Non</span><span class="sxs-lookup"><span data-stu-id="470c7-148">No</span></span> |  <span data-ttu-id="470c7-149">Indique si le contrôle est activé au lancement du complément.</span><span class="sxs-lookup"><span data-stu-id="470c7-149">Specifies whether the control is enabled when the add-in launches.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="470c7-150">Exemple du bouton ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="470c7-150">ExecuteFunction button example</span></span>

<span data-ttu-id="470c7-151">Dans l’exemple suivant, le bouton est désactivé au lancement du complément.</span><span class="sxs-lookup"><span data-stu-id="470c7-151">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="470c7-152">Il peut être activé par programmation.</span><span class="sxs-lookup"><span data-stu-id="470c7-152">It can be programmatically enabled.</span></span> <span data-ttu-id="470c7-153">Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="470c7-153">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

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

### <a name="showtaskpane-button-example"></a><span data-ttu-id="470c7-154">Exemple du bouton ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="470c7-154">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="470c7-155">Contrôles de menu (bouton déroulant)</span><span class="sxs-lookup"><span data-stu-id="470c7-155">Menu (dropdown button) controls</span></span>

<span data-ttu-id="470c7-p109">Un menu définit une liste statique d’options. Chaque option de menu exécute une fonction ou affiche un volet Office. Les sous-menus ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="470c7-p109">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="470c7-159">Lorsqu’il est utilisé avec un **point d’extension** **PrimaryCommandSurface** ou [ContextMenu](extensionpoint.md), le contrôle de menu définit les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="470c7-159">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="470c7-160">une option de menu de niveau racine.</span><span class="sxs-lookup"><span data-stu-id="470c7-160">A root-level menu item.</span></span>

- <span data-ttu-id="470c7-161">une liste de sous-menus.</span><span class="sxs-lookup"><span data-stu-id="470c7-161">A list of submenu items.</span></span>

<span data-ttu-id="470c7-p110">Lorsqu’il est utilisé avec **PrimaryCommandSurface**, l’option de menu de niveau racine s’affiche sous la forme d’un bouton dans le ruban. Lorsque le bouton est sélectionné, le sous-menu s’affiche sous la forme d’une liste déroulante. Lorsqu’il est utilisé avec **ContextMenu**, un élément de menu avec un sous-menu est inséré dans le menu contextuel. Dans les deux cas, les éléments individuels du sous-menu peuvent exécuter une fonction JavaScript ou afficher un volet de tâches. Un seul niveau de sous-menus est pris en charge pour l’instant.</span><span class="sxs-lookup"><span data-stu-id="470c7-p110">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="470c7-p111">L’exemple suivant montre comment définir un élément de menu avec deux éléments de sous-menu. Le premier élément de sous-menu affiche un volet Office et le deuxième élément de sous-menu exécute une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="470c7-p111">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="470c7-169">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="470c7-169">Child elements</span></span>

|  <span data-ttu-id="470c7-170">Élément</span><span class="sxs-lookup"><span data-stu-id="470c7-170">Element</span></span> |  <span data-ttu-id="470c7-171">Requis</span><span class="sxs-lookup"><span data-stu-id="470c7-171">Required</span></span>  |  <span data-ttu-id="470c7-172">Description</span><span class="sxs-lookup"><span data-stu-id="470c7-172">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="470c7-173">**Label**</span><span class="sxs-lookup"><span data-stu-id="470c7-173">**Label**</span></span>     | <span data-ttu-id="470c7-174">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-174">Yes</span></span> |  <span data-ttu-id="470c7-175">Texte du bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-175">The text for the button.</span></span> <span data-ttu-id="470c7-176">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="470c7-176">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="470c7-177">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="470c7-177">**ToolTip**</span></span>    |<span data-ttu-id="470c7-178">Non</span><span class="sxs-lookup"><span data-stu-id="470c7-178">No</span></span>|<span data-ttu-id="470c7-179">Info-bulle pour le bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-179">The tooltip for the button.</span></span> <span data-ttu-id="470c7-180">L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**.</span><span class="sxs-lookup"><span data-stu-id="470c7-180">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="470c7-181">**String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="470c7-181">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="470c7-182">Supertip</span><span class="sxs-lookup"><span data-stu-id="470c7-182">Supertip</span></span>](supertip.md)  | <span data-ttu-id="470c7-183">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-183">Yes</span></span> |  <span data-ttu-id="470c7-184">Info-bulle pour ce bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-184">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="470c7-185">Icon</span><span class="sxs-lookup"><span data-stu-id="470c7-185">Icon</span></span>](icon.md)      | <span data-ttu-id="470c7-186">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-186">Yes</span></span> |  <span data-ttu-id="470c7-187">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-187">An image for the button.</span></span>         |
|  <span data-ttu-id="470c7-188">**Éléments**</span><span class="sxs-lookup"><span data-stu-id="470c7-188">**Items**</span></span>     | <span data-ttu-id="470c7-189">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-189">Yes</span></span> |  <span data-ttu-id="470c7-190">Collection de boutons à afficher dans le menu.</span><span class="sxs-lookup"><span data-stu-id="470c7-190">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="470c7-191">Contient les éléments **Élément** pour chaque élément de sous-menu.</span><span class="sxs-lookup"><span data-stu-id="470c7-191">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="470c7-192">Chaque élément **Item** contient les éléments enfants du [contrôle Button](#button-control).</span><span class="sxs-lookup"><span data-stu-id="470c7-192">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="470c7-193">Exemples de contrôle de menu</span><span class="sxs-lookup"><span data-stu-id="470c7-193">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="470c7-194">Contrôle MobileButton</span><span class="sxs-lookup"><span data-stu-id="470c7-194">MobileButton control</span></span>

<span data-ttu-id="470c7-p115">Un bouton mobile effectue une action unique lorsque l’utilisateur le sélectionne. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle de bouton mobile doit avoir un `id` unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="470c7-p115">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="470c7-p116">La valeur `MobileButton` de **xsi:type** est définie dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="470c7-p116">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="470c7-200">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="470c7-200">Child elements</span></span>
|  <span data-ttu-id="470c7-201">Élément</span><span class="sxs-lookup"><span data-stu-id="470c7-201">Element</span></span> |  <span data-ttu-id="470c7-202">Requis</span><span class="sxs-lookup"><span data-stu-id="470c7-202">Required</span></span>  |  <span data-ttu-id="470c7-203">Description</span><span class="sxs-lookup"><span data-stu-id="470c7-203">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="470c7-204">**Label**</span><span class="sxs-lookup"><span data-stu-id="470c7-204">**Label**</span></span>     | <span data-ttu-id="470c7-205">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-205">Yes</span></span> |  <span data-ttu-id="470c7-206">Texte du bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-206">The text for the button.</span></span> <span data-ttu-id="470c7-207">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="470c7-207">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="470c7-208">Icon</span><span class="sxs-lookup"><span data-stu-id="470c7-208">Icon</span></span>](icon.md)      | <span data-ttu-id="470c7-209">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-209">Yes</span></span> |  <span data-ttu-id="470c7-210">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="470c7-210">An image for the button.</span></span>         |
|  [<span data-ttu-id="470c7-211">Action</span><span class="sxs-lookup"><span data-stu-id="470c7-211">Action</span></span>](action.md)    | <span data-ttu-id="470c7-212">Oui</span><span class="sxs-lookup"><span data-stu-id="470c7-212">Yes</span></span> |  <span data-ttu-id="470c7-213">Spécifie l’action à effectuer.</span><span class="sxs-lookup"><span data-stu-id="470c7-213">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="470c7-214">Exemple de bouton mobile ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="470c7-214">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="470c7-215">Exemple de bouton mobile ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="470c7-215">ShowTaskpane mobile button example</span></span>

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
