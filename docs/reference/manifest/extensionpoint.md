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
# <a name="extensionpoint-element"></a><span data-ttu-id="135dc-103">Élément ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="135dc-103">ExtensionPoint element</span></span>

 <span data-ttu-id="135dc-104">Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="135dc-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="135dc-105">L’élément **ExtensionPoint** est un élément enfant de [AllFormFactors](allformfactors.md) ou [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="135dc-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="135dc-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="135dc-106">Attributes</span></span>

|  <span data-ttu-id="135dc-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="135dc-107">Attribute</span></span>  |  <span data-ttu-id="135dc-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="135dc-108">Required</span></span>  |  <span data-ttu-id="135dc-109">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="135dc-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="135dc-110">**xsi:type**</span></span>  |  <span data-ttu-id="135dc-111">Oui</span><span class="sxs-lookup"><span data-stu-id="135dc-111">Yes</span></span>  | <span data-ttu-id="135dc-112">Type de point d’extension défini.</span><span class="sxs-lookup"><span data-stu-id="135dc-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="135dc-113">Points d’extension pour Excel uniquement</span><span class="sxs-lookup"><span data-stu-id="135dc-113">Extension points for Excel only</span></span>

- <span data-ttu-id="135dc-114">**CustomFunctions** – fonction personnalisée écrite en JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="135dc-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="135dc-115">[L’exemple de code XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) montre comment utiliser l’élément **ExtensionPoint** avec la valeur d’attribut **CustomFunctions** et les éléments enfants à utiliser.</span><span class="sxs-lookup"><span data-stu-id="135dc-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="135dc-116">Points d’extension pour les commandes de complément Word, Excel, PowerPoint et OneNote</span><span class="sxs-lookup"><span data-stu-id="135dc-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="135dc-117">**PrimaryCommandSurface** : ruban dans Office.</span><span class="sxs-lookup"><span data-stu-id="135dc-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="135dc-118">**ContextMenu** : menu contextuel qui apparaît lorsque vous cliquez avec le bouton droit de la souris dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="135dc-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="135dc-119">Les exemples suivants montrent comment utiliser l’élément **ExtensionPoint** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.</span><span class="sxs-lookup"><span data-stu-id="135dc-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="135dc-p102">Pour les éléments qui contiennent un attribut ID, assurez-vous que vous indiquez un ID unique. Nous vous recommandons d’utiliser le nom de votre organisation, ainsi que votre ID. Par exemple, utilisez le format suivant. <CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="135dc-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="135dc-123">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="135dc-123">Child elements</span></span>
 
|<span data-ttu-id="135dc-124">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-124">Element</span></span>|<span data-ttu-id="135dc-125">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-125">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="135dc-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="135dc-126">**CustomTab**</span></span>|<span data-ttu-id="135dc-p103">Obligatoire si vous souhaitez ajouter un onglet personnalisé au ruban (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **CustomTab**, vous ne pouvez pas utiliser l’élément **OfficeTab**. L’attribut **id** est obligatoire. </span><span class="sxs-lookup"><span data-stu-id="135dc-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="135dc-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="135dc-130">**OfficeTab**</span></span>|<span data-ttu-id="135dc-131">Obligatoire si vous souhaitez étendre un onglet de ruban d’application Office par défaut (à l’aide **de PrimaryCommandSurface).**</span><span class="sxs-lookup"><span data-stu-id="135dc-131">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="135dc-132">Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**.</span><span class="sxs-lookup"><span data-stu-id="135dc-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="135dc-133">Pour plus d’informations, voir [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="135dc-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="135dc-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="135dc-134">**OfficeMenu**</span></span>|<span data-ttu-id="135dc-p105">Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : </span><span class="sxs-lookup"><span data-stu-id="135dc-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="135dc-p106">- **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. </span><span class="sxs-lookup"><span data-stu-id="135dc-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="135dc-p107">- **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="135dc-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="135dc-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="135dc-141">**Group**</span></span>|<span data-ttu-id="135dc-p108">Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est obligatoire. Il s’agit d’une chaîne avec un maximum de 125 caractères. </span><span class="sxs-lookup"><span data-stu-id="135dc-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="135dc-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="135dc-145">**Label**</span></span>|<span data-ttu-id="135dc-146">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="135dc-146">Required.</span></span> <span data-ttu-id="135dc-147">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="135dc-147">The label of the group.</span></span> <span data-ttu-id="135dc-148">**L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un **élément String.**</span><span class="sxs-lookup"><span data-stu-id="135dc-148">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="135dc-149">L’élément  **String** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément  **Resources**.</span><span class="sxs-lookup"><span data-stu-id="135dc-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="135dc-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="135dc-150">**Icon**</span></span>|<span data-ttu-id="135dc-151">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="135dc-151">Required.</span></span> <span data-ttu-id="135dc-152">Indique l’icône du groupe qui doit être utilisée sur les périphériques de petit facteur de forme ou lorsque les boutons sont affichés en trop grand nombre.</span><span class="sxs-lookup"><span data-stu-id="135dc-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span></span> <span data-ttu-id="135dc-153">**L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un **élément Image.**</span><span class="sxs-lookup"><span data-stu-id="135dc-153">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="135dc-154">L’élément  **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément  **Resources**.</span><span class="sxs-lookup"><span data-stu-id="135dc-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="135dc-155">L’attribut **size** donne la taille, en pixels, de l’image.</span><span class="sxs-lookup"><span data-stu-id="135dc-155">The **size** attribute gives the size, in pixels, of the image.</span></span> <span data-ttu-id="135dc-156">Trois tailles d’image, en pixels, sont obligatoires : 16, 32 et 80.</span><span class="sxs-lookup"><span data-stu-id="135dc-156">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="135dc-157">Cinq tailles facultatives, en pixels, sont également prises en charge : 20, 24, 40, 48 et 64.</span><span class="sxs-lookup"><span data-stu-id="135dc-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="135dc-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="135dc-158">**Tooltip**</span></span>|<span data-ttu-id="135dc-159">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="135dc-159">Optional.</span></span> <span data-ttu-id="135dc-160">Info-bulle du groupe.</span><span class="sxs-lookup"><span data-stu-id="135dc-160">The tooltip of the group.</span></span> <span data-ttu-id="135dc-161">**L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un **élément String.**</span><span class="sxs-lookup"><span data-stu-id="135dc-161">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="135dc-162">L’élément  **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément  **Resources**.</span><span class="sxs-lookup"><span data-stu-id="135dc-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="135dc-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="135dc-163">**Control**</span></span>|<span data-ttu-id="135dc-164">Chaque groupe requiert au moins un contrôle.</span><span class="sxs-lookup"><span data-stu-id="135dc-164">Each group requires at least one control.</span></span> <span data-ttu-id="135dc-165">Un élément **Control** peut être de type **Button** ou **Menu**.</span><span class="sxs-lookup"><span data-stu-id="135dc-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="135dc-166">Utilisez **Menu** pour spécifier une liste déroulante de contrôles de bouton.</span><span class="sxs-lookup"><span data-stu-id="135dc-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="135dc-167">Actuellement, seuls les boutons et les menus sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="135dc-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="135dc-168">Pour plus d’informations, reportez-vous aux sections [Contrôles de bouton](control.md#button-control) et [Contrôles de menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="135dc-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="135dc-169">**Remarque :**  Pour faciliter la résolution des problèmes, nous vous recommandons d’ajouter un élément **Control** et les éléments enfants **Resources** associés un par un.</span><span class="sxs-lookup"><span data-stu-id="135dc-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="135dc-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="135dc-170">**Script**</span></span>|<span data-ttu-id="135dc-171">Liens vers le fichier JavaScript avec la définition de la fonction personnalisée et le code d’inscription.</span><span class="sxs-lookup"><span data-stu-id="135dc-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="135dc-172">Cet élément n’est pas utilisé dans l’aperçu pour les développeurs.</span><span class="sxs-lookup"><span data-stu-id="135dc-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="135dc-173">À la place, la page HTML est responsable du chargement de tous les fichiers JavaScript.</span><span class="sxs-lookup"><span data-stu-id="135dc-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="135dc-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="135dc-174">**Page**</span></span>|<span data-ttu-id="135dc-175">Liens vers la page HTML de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="135dc-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="135dc-176">Points d’extension pour Outlook</span><span class="sxs-lookup"><span data-stu-id="135dc-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="135dc-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="135dc-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="135dc-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="135dc-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="135dc-181">[Module](#module) (peut uniquement être utilisé dans [DesktopFormFactor](desktopformfactor.md).)</span><span class="sxs-lookup"><span data-stu-id="135dc-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="135dc-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="135dc-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface)
- [<span data-ttu-id="135dc-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="135dc-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="135dc-185">Événements</span><span class="sxs-lookup"><span data-stu-id="135dc-185">Events</span></span>](#events)
- [<span data-ttu-id="135dc-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="135dc-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="135dc-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="135dc-p114">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique. Dans l’application de bureau Outlook, cela apparaît dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="135dc-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="135dc-190">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="135dc-190">Child elements</span></span>

|  <span data-ttu-id="135dc-191">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-191">Element</span></span> |  <span data-ttu-id="135dc-192">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="135dc-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="135dc-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="135dc-194">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="135dc-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="135dc-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="135dc-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="135dc-196">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="135dc-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="135dc-197">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="135dc-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="135dc-198">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="135dc-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="135dc-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="135dc-200">Ce point d’extension place des boutons sur le ruban pour les compléments à l’aide du formulaire de composition de messagerie.</span><span class="sxs-lookup"><span data-stu-id="135dc-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="135dc-201">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="135dc-201">Child elements</span></span>

|  <span data-ttu-id="135dc-202">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-202">Element</span></span> |  <span data-ttu-id="135dc-203">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="135dc-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="135dc-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="135dc-205">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="135dc-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="135dc-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="135dc-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="135dc-207">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="135dc-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="135dc-208">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="135dc-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="135dc-209">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="135dc-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="135dc-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="135dc-211">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention de l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="135dc-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="135dc-212">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="135dc-212">Child elements</span></span>

|  <span data-ttu-id="135dc-213">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-213">Element</span></span> |  <span data-ttu-id="135dc-214">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="135dc-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="135dc-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="135dc-216">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="135dc-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="135dc-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="135dc-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="135dc-218">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="135dc-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="135dc-219">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="135dc-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="135dc-220">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="135dc-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="135dc-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="135dc-222">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention du participant à la réunion.</span><span class="sxs-lookup"><span data-stu-id="135dc-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="135dc-223">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="135dc-223">Child elements</span></span>

|  <span data-ttu-id="135dc-224">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-224">Element</span></span> |  <span data-ttu-id="135dc-225">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="135dc-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="135dc-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="135dc-227">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="135dc-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="135dc-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="135dc-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="135dc-229">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="135dc-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="135dc-230">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="135dc-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="135dc-231">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="135dc-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="135dc-232">Module</span><span class="sxs-lookup"><span data-stu-id="135dc-232">Module</span></span>

<span data-ttu-id="135dc-233">Ce point d’extension place des boutons sur le ruban pour l’extension de module.</span><span class="sxs-lookup"><span data-stu-id="135dc-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="135dc-234">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="135dc-234">Child elements</span></span>

|  <span data-ttu-id="135dc-235">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-235">Element</span></span> |  <span data-ttu-id="135dc-236">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-236">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="135dc-237">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="135dc-237">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="135dc-238">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="135dc-238">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="135dc-239">CustomTab</span><span class="sxs-lookup"><span data-stu-id="135dc-239">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="135dc-240">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="135dc-240">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="135dc-241">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-241">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="135dc-242">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique dans le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="135dc-242">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="135dc-243">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="135dc-243">Child elements</span></span>

|  <span data-ttu-id="135dc-244">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-244">Element</span></span> |  <span data-ttu-id="135dc-245">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-245">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="135dc-246">Group</span><span class="sxs-lookup"><span data-stu-id="135dc-246">Group</span></span>](group.md) |  <span data-ttu-id="135dc-247">Ajoute un groupe de boutons à la surface de commande.</span><span class="sxs-lookup"><span data-stu-id="135dc-247">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="135dc-248">Les éléments **ExtensionPoint** de ce type peuvent uniquement avoir un élément enfant, à savoir un élément **Group**.</span><span class="sxs-lookup"><span data-stu-id="135dc-248">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="135dc-249">Pour les éléments **Control** contenus dans ce point d’extension, l’attribut **xsi:type** doit avoir la valeur `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="135dc-249">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="135dc-250">Exemple</span><span class="sxs-lookup"><span data-stu-id="135dc-250">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface"></a><span data-ttu-id="135dc-251">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="135dc-251">MobileOnlineMeetingCommandSurface</span></span>

<span data-ttu-id="135dc-252">Ce point d’extension place un basculement approprié en mode dans l’surface de commande d’un rendez-vous dans le facteur de forme mobile.</span><span class="sxs-lookup"><span data-stu-id="135dc-252">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="135dc-253">Un organisateur de réunion peut créer une réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="135dc-253">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="135dc-254">Un participant peut ensuite participer à la réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="135dc-254">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="135dc-255">Pour en savoir plus sur ce scénario, consultez l’article Créer un application mobile Outlook pour un fournisseur de réunion [en ligne.](../../outlook/online-meeting.md)</span><span class="sxs-lookup"><span data-stu-id="135dc-255">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

> [!NOTE]
> <span data-ttu-id="135dc-256">Ce point d’extension est uniquement pris en charge sur Android avec un abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="135dc-256">This extension point is only supported on Android with a Microsoft 365 subscription.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="135dc-257">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="135dc-257">Child elements</span></span>

|  <span data-ttu-id="135dc-258">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-258">Element</span></span> |  <span data-ttu-id="135dc-259">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-259">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="135dc-260">Control</span><span class="sxs-lookup"><span data-stu-id="135dc-260">Control</span></span>](control.md) |  <span data-ttu-id="135dc-261">Ajoute un bouton à la surface de commande.</span><span class="sxs-lookup"><span data-stu-id="135dc-261">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="135dc-262">`ExtensionPoint` les éléments de ce type ne peuvent avoir qu’un seul élément enfant : un `Control` élément.</span><span class="sxs-lookup"><span data-stu-id="135dc-262">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="135dc-263">L’attribut doit être attribué à l’élément contenu dans ce `Control` point `xsi:type` d’extension. `MobileButton`</span><span class="sxs-lookup"><span data-stu-id="135dc-263">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="135dc-264">Les images doivent être en échelles de gris à l’aide de code hex ou de son équivalent `Icon` `#919191` dans [d’autres formats de couleur.](https://convertingcolors.com/hex-color-919191.html)</span><span class="sxs-lookup"><span data-stu-id="135dc-264">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="135dc-265">Exemple</span><span class="sxs-lookup"><span data-stu-id="135dc-265">Example</span></span>

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

### <a name="launchevent-preview"></a><span data-ttu-id="135dc-266">LaunchEvent (aperçu)</span><span class="sxs-lookup"><span data-stu-id="135dc-266">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="135dc-267">Ce point d’extension est uniquement pris en charge en [prévisualisation](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) dans Outlook sur le web et sur Windows avec un abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="135dc-267">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="135dc-268">Ce point d’extension permet à un application de s’activer en fonction des événements pris en charge dans le facteur de forme de bureau.</span><span class="sxs-lookup"><span data-stu-id="135dc-268">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="135dc-269">Actuellement, les seuls événements pris en charge sont `OnNewMessageCompose` et `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="135dc-269">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="135dc-270">Pour en savoir plus sur ce scénario, consultez l’article Configurer votre complément Outlook pour [l’activation](../../outlook/autolaunch.md) basée sur un événement.</span><span class="sxs-lookup"><span data-stu-id="135dc-270">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="135dc-271">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="135dc-271">Child elements</span></span>

|  <span data-ttu-id="135dc-272">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-272">Element</span></span> |  <span data-ttu-id="135dc-273">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-273">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="135dc-274">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="135dc-274">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="135dc-275">Liste de [LaunchEvent pour](launchevent.md) l’activation basée sur des événements.</span><span class="sxs-lookup"><span data-stu-id="135dc-275">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="135dc-276">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="135dc-276">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="135dc-277">Emplacement du fichier JavaScript source.</span><span class="sxs-lookup"><span data-stu-id="135dc-277">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="135dc-278">Exemple</span><span class="sxs-lookup"><span data-stu-id="135dc-278">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="135dc-279">Événements</span><span class="sxs-lookup"><span data-stu-id="135dc-279">Events</span></span>

<span data-ttu-id="135dc-280">Ce point d’extension ajoute un gestionnaire d’événements pour un événement spécifié.</span><span class="sxs-lookup"><span data-stu-id="135dc-280">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="135dc-281">Pour plus d’informations sur l’utilisation de ce point d’extension, voir La fonctionnalité d’envoi pour [les modules complémentaires Outlook.](../../outlook/outlook-on-send-addins.md)</span><span class="sxs-lookup"><span data-stu-id="135dc-281">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

| <span data-ttu-id="135dc-282">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-282">Element</span></span> | <span data-ttu-id="135dc-283">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-283">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="135dc-284">Event</span><span class="sxs-lookup"><span data-stu-id="135dc-284">Event</span></span>](event.md) |  <span data-ttu-id="135dc-285">Indique l’événement et la fonction gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="135dc-285">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="135dc-286">Exemple d’événement ItemSend</span><span class="sxs-lookup"><span data-stu-id="135dc-286">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="135dc-287">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="135dc-287">DetectedEntity</span></span>

<span data-ttu-id="135dc-288">Ce point d’extension ajoute une activation de complément contextuel sur un type d’entité spécifié.</span><span class="sxs-lookup"><span data-stu-id="135dc-288">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="135dc-289">Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, `xsi:type`l’attribut doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="135dc-289">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="135dc-290">Ce type d’élément est disponible pour [les clients Outlook qui prennent en charge les ensembles de conditions requises 1.6 ou version ultérieure](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="135dc-290">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="135dc-291">Élément</span><span class="sxs-lookup"><span data-stu-id="135dc-291">Element</span></span> |  <span data-ttu-id="135dc-292">Description</span><span class="sxs-lookup"><span data-stu-id="135dc-292">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="135dc-293">Label</span><span class="sxs-lookup"><span data-stu-id="135dc-293">Label</span></span>](#label) |  <span data-ttu-id="135dc-294">Spécifie l’étiquette pour le complément dans la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="135dc-294">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="135dc-295">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="135dc-295">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="135dc-296">Spécifie l’URL de la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="135dc-296">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="135dc-297">Règle</span><span class="sxs-lookup"><span data-stu-id="135dc-297">Rule</span></span>](rule.md) |  <span data-ttu-id="135dc-298">Spécifie la ou les règles qui déterminent lorsqu’un complément s’active.</span><span class="sxs-lookup"><span data-stu-id="135dc-298">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="135dc-299">Étiquette</span><span class="sxs-lookup"><span data-stu-id="135dc-299">Label</span></span>

<span data-ttu-id="135dc-300">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="135dc-300">Required.</span></span> <span data-ttu-id="135dc-301">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="135dc-301">The label of the group.</span></span> <span data-ttu-id="135dc-302">**L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="135dc-302">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="135dc-303">Exigences relatives à la mise en surbrillance</span><span class="sxs-lookup"><span data-stu-id="135dc-303">Highlight requirements</span></span>

<span data-ttu-id="135dc-p119">Le seul moyen pour qu’un utilisateur puisse activer un complément contextuel consiste à interagir avec une entité en surbrillance. Les développeurs peuvent contrôler les entités qui sont mises en surbrillance à l’aide de l’attribut `Highlight` de l’élément `Rule` pour les types de règles `ItemHasKnownEntity` et `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="135dc-p119">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="135dc-p120">Toutefois, il existe certaines limitations à connaître. Ces limitations sont en place pour vous assurer qu’il y aura toujours une entité en surbrillance dans les messages ou rendez-vous applicables pour permettre à l’utilisateur d’activer le complément.</span><span class="sxs-lookup"><span data-stu-id="135dc-p120">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="135dc-308">Les types d’entité `EmailAddress` et `Url` ne peuvent pas être mis en surbrillance et par conséquent ne peuvent pas être utilisés pour activer un complément.</span><span class="sxs-lookup"><span data-stu-id="135dc-308">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="135dc-309">Si vous utilisez une seule règle, la valeur `Highlight` DOIT être définie sur `all`.</span><span class="sxs-lookup"><span data-stu-id="135dc-309">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="135dc-310">Si vous utilisez un type de règle `RuleCollection` avec `Mode="AND"` pour combiner plusieurs règles, au moins l’une des règles DOIT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="135dc-310">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="135dc-311">Si vous utilisez un type de règle `RuleCollection` avec `Mode="OR"` pour combiner plusieurs règles, toutes les règles DOIVENT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="135dc-311">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="135dc-312">Exemple d’événement DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="135dc-312">DetectedEntity event example</span></span>

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
