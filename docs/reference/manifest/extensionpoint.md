---
title: Élément Extension dans le fichier manifeste
description: Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 56a7365f986060a225cc0d20c89a310deb4a52b3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611868"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="9c6ee-103">Élément ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="9c6ee-103">ExtensionPoint element</span></span>

 <span data-ttu-id="9c6ee-104">Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="9c6ee-105">L’élément **ExtensionPoint** est un élément enfant de [AllFormFactors](allformfactors.md) ou [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="9c6ee-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="9c6ee-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="9c6ee-106">Attributes</span></span>

|  <span data-ttu-id="9c6ee-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="9c6ee-107">Attribute</span></span>  |  <span data-ttu-id="9c6ee-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="9c6ee-108">Required</span></span>  |  <span data-ttu-id="9c6ee-109">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="9c6ee-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-110">**xsi:type**</span></span>  |  <span data-ttu-id="9c6ee-111">Oui</span><span class="sxs-lookup"><span data-stu-id="9c6ee-111">Yes</span></span>  | <span data-ttu-id="9c6ee-112">Type de point d’extension défini.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="9c6ee-113">Points d’extension pour Excel uniquement</span><span class="sxs-lookup"><span data-stu-id="9c6ee-113">Extension points for Excel only</span></span>

- <span data-ttu-id="9c6ee-114">**CustomFunctions** – fonction personnalisée écrite en JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="9c6ee-115">[L’exemple de code XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) montre comment utiliser l’élément **ExtensionPoint** avec la valeur d’attribut **CustomFunctions** et les éléments enfants à utiliser.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="9c6ee-116">Points d’extension pour les commandes de complément Word, Excel, PowerPoint et OneNote</span><span class="sxs-lookup"><span data-stu-id="9c6ee-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="9c6ee-117">**PrimaryCommandSurface** : ruban dans Office.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="9c6ee-118">**ContextMenu** : menu contextuel qui apparaît lorsque vous cliquez avec le bouton droit de la souris dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="9c6ee-119">Les exemples suivants montrent comment utiliser l’élément **ExtensionPoint** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9c6ee-p102">Pour les éléments qui contiennent un attribut ID, assurez-vous que vous indiquez un ID unique. Nous vous recommandons d’utiliser le nom de votre organisation, ainsi que votre ID. Par exemple, utilisez le format suivant. <CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="9c6ee-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="9c6ee-123">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9c6ee-123">Child elements</span></span>
 
|<span data-ttu-id="9c6ee-124">**Élément**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-124">**Element**</span></span>|<span data-ttu-id="9c6ee-125">**Description**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9c6ee-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-126">**CustomTab**</span></span>|<span data-ttu-id="9c6ee-p103">Obligatoire si vous souhaitez ajouter un onglet personnalisé au ruban (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **CustomTab**, vous ne pouvez pas utiliser l’élément **OfficeTab**. L’attribut **id** est obligatoire. </span><span class="sxs-lookup"><span data-stu-id="9c6ee-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="9c6ee-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-130">**OfficeTab**</span></span>|<span data-ttu-id="9c6ee-131">Obligatoire pour étendre un onglet du ruban Office par défaut (en utilisant **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="9c6ee-131">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="9c6ee-132">Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="9c6ee-133">Pour plus d’informations, voir [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="9c6ee-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="9c6ee-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-134">**OfficeMenu**</span></span>|<span data-ttu-id="9c6ee-p105">Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : </span><span class="sxs-lookup"><span data-stu-id="9c6ee-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="9c6ee-p106">- **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. </span><span class="sxs-lookup"><span data-stu-id="9c6ee-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="9c6ee-p107">- **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="9c6ee-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-141">**Group**</span></span>|<span data-ttu-id="9c6ee-p108">Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est obligatoire. Il s’agit d’une chaîne avec un maximum de 125 caractères. </span><span class="sxs-lookup"><span data-stu-id="9c6ee-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="9c6ee-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-145">**Label**</span></span>|<span data-ttu-id="9c6ee-p109">Obligatoire. L’étiquette du groupe. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="9c6ee-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="9c6ee-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-150">**Icon**</span></span>|<span data-ttu-id="9c6ee-p110">Obligatoire. Spécifie l’icône du groupe à utiliser sur de petits appareils, ou lorsqu’un nombre trop important de boutons est affiché. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image**. **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément **Ressources**. L’attribut **size** donne la taille, en pixels, de l’image. Trois tailles d’images sont obligatoires : 16, 32 et 80. 5 tailles facultatives sont également prises en charge : 20, 24, 40, 48 et 64. </span><span class="sxs-lookup"><span data-stu-id="9c6ee-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="9c6ee-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-158">**Tooltip**</span></span>|<span data-ttu-id="9c6ee-p111">Facultatif. Info-bulle du groupe. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="9c6ee-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="9c6ee-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-163">**Control**</span></span>|<span data-ttu-id="9c6ee-164">Chaque groupe requiert au moins un contrôle.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-164">Each group requires at least one control.</span></span> <span data-ttu-id="9c6ee-165">Un élément **Control** peut être de type **Button** ou **Menu**.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="9c6ee-166">Utilisez **Menu** pour spécifier une liste déroulante de contrôles de bouton.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="9c6ee-167">Actuellement, seuls les boutons et les menus sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="9c6ee-168">Pour plus d’informations, reportez-vous aux sections [Contrôles de bouton](control.md#button-control) et [Contrôles de menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="9c6ee-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="9c6ee-169">**Remarque :**  Pour faciliter la résolution des problèmes, nous vous recommandons d’ajouter un élément **Control** et les éléments enfants **Resources** associés un par un.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="9c6ee-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-170">**Script**</span></span>|<span data-ttu-id="9c6ee-171">Liens vers le fichier JavaScript avec la définition de la fonction personnalisée et le code d’inscription.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="9c6ee-172">Cet élément n’est pas utilisé dans l’aperçu pour les développeurs.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="9c6ee-173">À la place, la page HTML est responsable du chargement de tous les fichiers JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="9c6ee-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="9c6ee-174">**Page**</span></span>|<span data-ttu-id="9c6ee-175">Liens vers la page HTML de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="9c6ee-176">Points d’extension pour Outlook</span><span class="sxs-lookup"><span data-stu-id="9c6ee-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="9c6ee-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="9c6ee-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="9c6ee-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="9c6ee-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="9c6ee-181">[Module](#module) (peut uniquement être utilisé dans [DesktopFormFactor](desktopformfactor.md).)</span><span class="sxs-lookup"><span data-stu-id="9c6ee-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="9c6ee-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="9c6ee-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface-preview)
- [<span data-ttu-id="9c6ee-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="9c6ee-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="9c6ee-185">Événements</span><span class="sxs-lookup"><span data-stu-id="9c6ee-185">Events</span></span>](#events)
- [<span data-ttu-id="9c6ee-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="9c6ee-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="9c6ee-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="9c6ee-p114">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique. Dans l’application de bureau Outlook, cela apparaît dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="9c6ee-190">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9c6ee-190">Child elements</span></span>

|  <span data-ttu-id="9c6ee-191">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-191">Element</span></span> |  <span data-ttu-id="9c6ee-192">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c6ee-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c6ee-194">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c6ee-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c6ee-196">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9c6ee-197">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9c6ee-198">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="9c6ee-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="9c6ee-200">Ce point d’extension place des boutons sur le ruban pour les compléments à l’aide du formulaire de composition de messagerie.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9c6ee-201">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9c6ee-201">Child elements</span></span>

|  <span data-ttu-id="9c6ee-202">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-202">Element</span></span> |  <span data-ttu-id="9c6ee-203">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c6ee-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c6ee-205">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c6ee-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c6ee-207">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9c6ee-208">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9c6ee-209">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="9c6ee-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="9c6ee-211">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention de l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9c6ee-212">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9c6ee-212">Child elements</span></span>

|  <span data-ttu-id="9c6ee-213">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-213">Element</span></span> |  <span data-ttu-id="9c6ee-214">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c6ee-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c6ee-216">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c6ee-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c6ee-218">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9c6ee-219">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9c6ee-220">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="9c6ee-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="9c6ee-222">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention du participant à la réunion.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9c6ee-223">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9c6ee-223">Child elements</span></span>

|  <span data-ttu-id="9c6ee-224">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-224">Element</span></span> |  <span data-ttu-id="9c6ee-225">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c6ee-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c6ee-227">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c6ee-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c6ee-229">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9c6ee-230">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9c6ee-231">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="9c6ee-232">Module</span><span class="sxs-lookup"><span data-stu-id="9c6ee-232">Module</span></span>

<span data-ttu-id="9c6ee-233">Ce point d’extension place des boutons sur le ruban pour l’extension de module.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="9c6ee-234">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9c6ee-234">Child elements</span></span>

|  <span data-ttu-id="9c6ee-235">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-235">Element</span></span> |  <span data-ttu-id="9c6ee-236">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-236">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c6ee-237">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-237">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c6ee-238">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-238">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c6ee-239">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c6ee-239">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c6ee-240">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-240">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="9c6ee-241">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c6ee-241">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="9c6ee-242">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique dans le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-242">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="9c6ee-243">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9c6ee-243">Child elements</span></span>

|  <span data-ttu-id="9c6ee-244">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-244">Element</span></span> |  <span data-ttu-id="9c6ee-245">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-245">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c6ee-246">Group</span><span class="sxs-lookup"><span data-stu-id="9c6ee-246">Group</span></span>](group.md) |  <span data-ttu-id="9c6ee-247">Ajoute un groupe de boutons à la surface de commande.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-247">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="9c6ee-248">Les éléments **ExtensionPoint** de ce type peuvent uniquement avoir un élément enfant, à savoir un élément **Group**.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-248">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="9c6ee-249">Pour les éléments **Control** contenus dans ce point d’extension, l’attribut **xsi:type** doit avoir la valeur `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-249">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="9c6ee-250">Exemple</span><span class="sxs-lookup"><span data-stu-id="9c6ee-250">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a><span data-ttu-id="9c6ee-251">MobileOnlineMeetingCommandSurface (aperçu)</span><span class="sxs-lookup"><span data-stu-id="9c6ee-251">MobileOnlineMeetingCommandSurface (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="9c6ee-252">Ce point d’extension est uniquement pris en charge en [Aperçu](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur Android avec un abonnement Office 365.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-252">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="9c6ee-253">Ce point d’extension place un bouton bascule mode-approprié dans la surface de commande pour un rendez-vous dans le facteur de forme mobile.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-253">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="9c6ee-254">Un organisateur de réunion peut créer une réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-254">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="9c6ee-255">Un participant peut ensuite participer à la réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-255">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="9c6ee-256">Pour en savoir plus sur ce scénario, consultez l’article [créer un complément Outlook Mobile pour un fournisseur de réunions en ligne](../../outlook/online-meeting.md) .</span><span class="sxs-lookup"><span data-stu-id="9c6ee-256">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="9c6ee-257">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9c6ee-257">Child elements</span></span>

|  <span data-ttu-id="9c6ee-258">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-258">Element</span></span> |  <span data-ttu-id="9c6ee-259">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-259">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c6ee-260">Control</span><span class="sxs-lookup"><span data-stu-id="9c6ee-260">Control</span></span>](control.md) |  <span data-ttu-id="9c6ee-261">Ajoute un bouton à la surface de commande.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-261">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="9c6ee-262">`ExtensionPoint`les éléments de ce type ne peuvent avoir qu’un seul élément enfant : un `Control` élément.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-262">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="9c6ee-263">L' `Control` attribut de l’élément contenu dans ce point d’extension doit être `xsi:type` défini sur `MobileButton` .</span><span class="sxs-lookup"><span data-stu-id="9c6ee-263">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="9c6ee-264">Les `Icon` images doivent être en nuances de gris à l’aide de code hexadécimal `#919191` ou de leur équivalent dans d' [autres formats de couleur](https://convertingcolors.com/hex-color-919191.html).</span><span class="sxs-lookup"><span data-stu-id="9c6ee-264">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="9c6ee-265">Exemple</span><span class="sxs-lookup"><span data-stu-id="9c6ee-265">Example</span></span>

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

### <a name="launchevent-preview"></a><span data-ttu-id="9c6ee-266">LaunchEvent (aperçu)</span><span class="sxs-lookup"><span data-stu-id="9c6ee-266">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="9c6ee-267">Ce point d’extension est uniquement pris en [charge dans Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur le Web avec un abonnement Office 365.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-267">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with an Office 365 subscription.</span></span>

<span data-ttu-id="9c6ee-268">Ce point d’extension permet à un complément de s’activer en fonction des événements pris en charge dans le facteur de forme de bureau.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-268">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="9c6ee-269">Actuellement, les seuls événements pris en charge sont `OnNewMessageCompose` et `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="9c6ee-269">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="9c6ee-270">Pour en savoir plus sur ce scénario, reportez-vous à l’article [configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="9c6ee-270">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="9c6ee-271">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9c6ee-271">Child elements</span></span>

|  <span data-ttu-id="9c6ee-272">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-272">Element</span></span> |  <span data-ttu-id="9c6ee-273">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-273">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="9c6ee-274">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="9c6ee-274">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="9c6ee-275">Liste des [LaunchEvent](launchevent.md) pour l’activation basée sur un événement.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-275">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="9c6ee-276">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9c6ee-276">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="9c6ee-277">Emplacement du fichier JavaScript source.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-277">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="9c6ee-278">Exemple</span><span class="sxs-lookup"><span data-stu-id="9c6ee-278">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="9c6ee-279">Événements</span><span class="sxs-lookup"><span data-stu-id="9c6ee-279">Events</span></span>

<span data-ttu-id="9c6ee-280">Ce point d’extension ajoute un gestionnaire d’événements pour un événement spécifié.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-280">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="9c6ee-281">Pour plus d’informations sur l’utilisation de ce point d’extension, consultez la rubrique relative à la [fonctionnalité d’envoi pour les compléments Outlook](../../outlook/outlook-on-send-addins.md).</span><span class="sxs-lookup"><span data-stu-id="9c6ee-281">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

| <span data-ttu-id="9c6ee-282">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-282">Element</span></span> | <span data-ttu-id="9c6ee-283">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-283">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c6ee-284">Event</span><span class="sxs-lookup"><span data-stu-id="9c6ee-284">Event</span></span>](event.md) |  <span data-ttu-id="9c6ee-285">Indique l’événement et la fonction gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-285">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="9c6ee-286">Exemple d’événement ItemSend</span><span class="sxs-lookup"><span data-stu-id="9c6ee-286">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="9c6ee-287">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="9c6ee-287">DetectedEntity</span></span>

<span data-ttu-id="9c6ee-288">Ce point d’extension ajoute une activation de complément contextuel sur un type d’entité spécifié.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-288">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="9c6ee-289">Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, `xsi:type`l’attribut doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-289">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="9c6ee-290">Ce type d’élément est disponible pour [les clients Outlook qui prennent en charge les ensembles de conditions requises 1.6 ou version ultérieure](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="9c6ee-290">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="9c6ee-291">Élément</span><span class="sxs-lookup"><span data-stu-id="9c6ee-291">Element</span></span> |  <span data-ttu-id="9c6ee-292">Description</span><span class="sxs-lookup"><span data-stu-id="9c6ee-292">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c6ee-293">Label</span><span class="sxs-lookup"><span data-stu-id="9c6ee-293">Label</span></span>](#label) |  <span data-ttu-id="9c6ee-294">Spécifie l’étiquette pour le complément dans la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-294">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="9c6ee-295">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9c6ee-295">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="9c6ee-296">Spécifie l’URL de la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-296">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="9c6ee-297">Règle</span><span class="sxs-lookup"><span data-stu-id="9c6ee-297">Rule</span></span>](rule.md) |  <span data-ttu-id="9c6ee-298">Spécifie la ou les règles qui déterminent lorsqu’un complément s’active.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-298">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="9c6ee-299">Étiquette</span><span class="sxs-lookup"><span data-stu-id="9c6ee-299">Label</span></span>

<span data-ttu-id="9c6ee-300">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-300">Required.</span></span> <span data-ttu-id="9c6ee-301">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-301">The label of the group.</span></span> <span data-ttu-id="9c6ee-302">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="9c6ee-302">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="9c6ee-303">Exigences relatives à la mise en surbrillance</span><span class="sxs-lookup"><span data-stu-id="9c6ee-303">Highlight requirements</span></span>

<span data-ttu-id="9c6ee-p119">Le seul moyen pour qu’un utilisateur puisse activer un complément contextuel consiste à interagir avec une entité en surbrillance. Les développeurs peuvent contrôler les entités qui sont mises en surbrillance à l’aide de l’attribut `Highlight` de l’élément `Rule` pour les types de règles `ItemHasKnownEntity` et `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-p119">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="9c6ee-p120">Toutefois, il existe certaines limitations à connaître. Ces limitations sont en place pour vous assurer qu’il y aura toujours une entité en surbrillance dans les messages ou rendez-vous applicables pour permettre à l’utilisateur d’activer le complément.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-p120">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="9c6ee-308">Les types d’entité `EmailAddress` et `Url` ne peuvent pas être mis en surbrillance et par conséquent ne peuvent pas être utilisés pour activer un complément.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-308">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="9c6ee-309">Si vous utilisez une seule règle, la valeur `Highlight` DOIT être définie sur `all`.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-309">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="9c6ee-310">Si vous utilisez un type de règle `RuleCollection` avec `Mode="AND"` pour combiner plusieurs règles, au moins l’une des règles DOIT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-310">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="9c6ee-311">Si vous utilisez un type de règle `RuleCollection` avec `Mode="OR"` pour combiner plusieurs règles, toutes les règles DOIVENT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="9c6ee-311">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="9c6ee-312">Exemple d’événement DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="9c6ee-312">DetectedEntity event example</span></span>

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
