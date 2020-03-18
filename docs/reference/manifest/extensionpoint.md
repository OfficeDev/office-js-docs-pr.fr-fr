---
title: Élément Extension dans le fichier manifeste
description: Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.
ms.date: 09/05/2019
localization_priority: Normal
ms.openlocfilehash: c945875140fdbdb7ba6aaeed7bb0a7bf5d06e050
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720567"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="b93fe-103">Élément ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="b93fe-103">ExtensionPoint element</span></span>

 <span data-ttu-id="b93fe-104">Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="b93fe-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="b93fe-105">L’élément **ExtensionPoint** est un élément enfant de [AllFormFactors](allformfactors.md) ou [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="b93fe-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="b93fe-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="b93fe-106">Attributes</span></span>

|  <span data-ttu-id="b93fe-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="b93fe-107">Attribute</span></span>  |  <span data-ttu-id="b93fe-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="b93fe-108">Required</span></span>  |  <span data-ttu-id="b93fe-109">Description</span><span class="sxs-lookup"><span data-stu-id="b93fe-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b93fe-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="b93fe-110">**xsi:type**</span></span>  |  <span data-ttu-id="b93fe-111">Oui</span><span class="sxs-lookup"><span data-stu-id="b93fe-111">Yes</span></span>  | <span data-ttu-id="b93fe-112">Type de point d’extension défini.</span><span class="sxs-lookup"><span data-stu-id="b93fe-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="b93fe-113">Points d’extension pour Excel uniquement</span><span class="sxs-lookup"><span data-stu-id="b93fe-113">Extension points for Excel only</span></span>

- <span data-ttu-id="b93fe-114">**CustomFunctions** – fonction personnalisée écrite en JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="b93fe-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="b93fe-115">[L’exemple de code XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) montre comment utiliser l’élément **ExtensionPoint** avec la valeur d’attribut **CustomFunctions** et les éléments enfants à utiliser.</span><span class="sxs-lookup"><span data-stu-id="b93fe-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="b93fe-116">Points d’extension pour les commandes de complément Word, Excel, PowerPoint et OneNote</span><span class="sxs-lookup"><span data-stu-id="b93fe-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="b93fe-117">**PrimaryCommandSurface** : ruban dans Office.</span><span class="sxs-lookup"><span data-stu-id="b93fe-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="b93fe-118">**ContextMenu** : menu contextuel qui apparaît lorsque vous cliquez avec le bouton droit de la souris dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="b93fe-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="b93fe-119">Les exemples suivants montrent comment utiliser l’élément **ExtensionPoint** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.</span><span class="sxs-lookup"><span data-stu-id="b93fe-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="b93fe-p102">Pour les éléments qui contiennent un attribut ID, assurez-vous que vous indiquez un ID unique. Nous vous recommandons d’utiliser le nom de votre organisation, ainsi que votre ID. Par exemple, utilisez le format suivant. <CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="b93fe-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="b93fe-123">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b93fe-123">Child elements</span></span>
 
|<span data-ttu-id="b93fe-124">**Élément**</span><span class="sxs-lookup"><span data-stu-id="b93fe-124">**Element**</span></span>|<span data-ttu-id="b93fe-125">**Description**</span><span class="sxs-lookup"><span data-stu-id="b93fe-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="b93fe-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="b93fe-126">**CustomTab**</span></span>|<span data-ttu-id="b93fe-p103">Obligatoire si vous souhaitez ajouter un onglet personnalisé au ruban (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **CustomTab**, vous ne pouvez pas utiliser l’élément **OfficeTab**. L’attribut **id** est obligatoire. </span><span class="sxs-lookup"><span data-stu-id="b93fe-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="b93fe-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="b93fe-130">**OfficeTab**</span></span>|<span data-ttu-id="b93fe-131">Obligatoire pour étendre un onglet du ruban Office par défaut (en utilisant **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="b93fe-131">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="b93fe-132">Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**.</span><span class="sxs-lookup"><span data-stu-id="b93fe-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="b93fe-133">Pour plus d’informations, voir [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="b93fe-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="b93fe-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="b93fe-134">**OfficeMenu**</span></span>|<span data-ttu-id="b93fe-p105">Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : </span><span class="sxs-lookup"><span data-stu-id="b93fe-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="b93fe-p106">- **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. </span><span class="sxs-lookup"><span data-stu-id="b93fe-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="b93fe-p107">- **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b93fe-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="b93fe-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="b93fe-141">**Group**</span></span>|<span data-ttu-id="b93fe-p108">Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est obligatoire. Il s’agit d’une chaîne avec un maximum de 125 caractères. </span><span class="sxs-lookup"><span data-stu-id="b93fe-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="b93fe-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="b93fe-145">**Label**</span></span>|<span data-ttu-id="b93fe-p109">Obligatoire. L’étiquette du groupe. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="b93fe-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="b93fe-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="b93fe-150">**Icon**</span></span>|<span data-ttu-id="b93fe-p110">Obligatoire. Spécifie l’icône du groupe à utiliser sur de petits appareils, ou lorsqu’un nombre trop important de boutons est affiché. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image**. **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément **Ressources**. L’attribut **size** donne la taille, en pixels, de l’image. Trois tailles d’images sont obligatoires : 16, 32 et 80. 5 tailles facultatives sont également prises en charge : 20, 24, 40, 48 et 64. </span><span class="sxs-lookup"><span data-stu-id="b93fe-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="b93fe-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="b93fe-158">**Tooltip**</span></span>|<span data-ttu-id="b93fe-p111">Facultatif. Info-bulle du groupe. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="b93fe-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="b93fe-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="b93fe-163">**Control**</span></span>|<span data-ttu-id="b93fe-164">Chaque groupe requiert au moins un contrôle.</span><span class="sxs-lookup"><span data-stu-id="b93fe-164">Each group requires at least one control.</span></span> <span data-ttu-id="b93fe-165">Un élément **Control** peut être de type **Button** ou **Menu**.</span><span class="sxs-lookup"><span data-stu-id="b93fe-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="b93fe-166">Utilisez **Menu** pour spécifier une liste déroulante de contrôles de bouton.</span><span class="sxs-lookup"><span data-stu-id="b93fe-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="b93fe-167">Actuellement, seuls les boutons et les menus sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="b93fe-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="b93fe-168">Pour plus d’informations, reportez-vous aux sections [Contrôles de bouton](control.md#button-control) et [Contrôles de menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="b93fe-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="b93fe-169">**Remarque :**  Pour faciliter la résolution des problèmes, nous vous recommandons d’ajouter un élément **Control** et les éléments enfants **Resources** associés un par un.</span><span class="sxs-lookup"><span data-stu-id="b93fe-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="b93fe-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="b93fe-170">**Script**</span></span>|<span data-ttu-id="b93fe-171">Liens vers le fichier JavaScript avec la définition de la fonction personnalisée et le code d’inscription.</span><span class="sxs-lookup"><span data-stu-id="b93fe-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="b93fe-172">Cet élément n’est pas utilisé dans l’aperçu pour les développeurs.</span><span class="sxs-lookup"><span data-stu-id="b93fe-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="b93fe-173">À la place, la page HTML est responsable du chargement de tous les fichiers JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b93fe-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="b93fe-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="b93fe-174">**Page**</span></span>|<span data-ttu-id="b93fe-175">Liens vers la page HTML de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b93fe-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="b93fe-176">Points d’extension pour Outlook</span><span class="sxs-lookup"><span data-stu-id="b93fe-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="b93fe-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="b93fe-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="b93fe-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="b93fe-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="b93fe-181">[Module](#module) (peut uniquement être utilisé dans [DesktopFormFactor](desktopformfactor.md).)</span><span class="sxs-lookup"><span data-stu-id="b93fe-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="b93fe-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="b93fe-183">Événements</span><span class="sxs-lookup"><span data-stu-id="b93fe-183">Events</span></span>](#events)
- [<span data-ttu-id="b93fe-184">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="b93fe-184">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="b93fe-185">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-185">MessageReadCommandSurface</span></span>
<span data-ttu-id="b93fe-p114">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique. Dans l’application de bureau Outlook, cela apparaît dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="b93fe-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="b93fe-188">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b93fe-188">Child elements</span></span>

|  <span data-ttu-id="b93fe-189">Élément</span><span class="sxs-lookup"><span data-stu-id="b93fe-189">Element</span></span> |  <span data-ttu-id="b93fe-190">Description</span><span class="sxs-lookup"><span data-stu-id="b93fe-190">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="b93fe-191">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-191">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="b93fe-192">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="b93fe-192">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="b93fe-193">CustomTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-193">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="b93fe-194">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b93fe-194">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="b93fe-195">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-195">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="b93fe-196">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-196">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="b93fe-197">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-197">MessageComposeCommandSurface</span></span>
<span data-ttu-id="b93fe-198">Ce point d’extension place des boutons sur le ruban pour les compléments à l’aide du formulaire de composition de messagerie.</span><span class="sxs-lookup"><span data-stu-id="b93fe-198">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="b93fe-199">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b93fe-199">Child elements</span></span>

|  <span data-ttu-id="b93fe-200">Élément</span><span class="sxs-lookup"><span data-stu-id="b93fe-200">Element</span></span> |  <span data-ttu-id="b93fe-201">Description</span><span class="sxs-lookup"><span data-stu-id="b93fe-201">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="b93fe-202">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-202">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="b93fe-203">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="b93fe-203">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="b93fe-204">CustomTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-204">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="b93fe-205">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b93fe-205">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="b93fe-206">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-206">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="b93fe-207">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-207">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="b93fe-208">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-208">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="b93fe-209">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention de l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="b93fe-209">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="b93fe-210">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b93fe-210">Child elements</span></span>

|  <span data-ttu-id="b93fe-211">Élément</span><span class="sxs-lookup"><span data-stu-id="b93fe-211">Element</span></span> |  <span data-ttu-id="b93fe-212">Description</span><span class="sxs-lookup"><span data-stu-id="b93fe-212">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="b93fe-213">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-213">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="b93fe-214">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="b93fe-214">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="b93fe-215">CustomTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-215">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="b93fe-216">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b93fe-216">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="b93fe-217">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-217">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="b93fe-218">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-218">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="b93fe-219">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-219">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="b93fe-220">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention du participant à la réunion.</span><span class="sxs-lookup"><span data-stu-id="b93fe-220">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="b93fe-221">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b93fe-221">Child elements</span></span>

|  <span data-ttu-id="b93fe-222">Élément</span><span class="sxs-lookup"><span data-stu-id="b93fe-222">Element</span></span> |  <span data-ttu-id="b93fe-223">Description</span><span class="sxs-lookup"><span data-stu-id="b93fe-223">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="b93fe-224">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-224">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="b93fe-225">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="b93fe-225">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="b93fe-226">CustomTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-226">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="b93fe-227">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b93fe-227">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="b93fe-228">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-228">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="b93fe-229">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-229">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="b93fe-230">Module</span><span class="sxs-lookup"><span data-stu-id="b93fe-230">Module</span></span>

<span data-ttu-id="b93fe-231">Ce point d’extension place des boutons sur le ruban pour l’extension de module.</span><span class="sxs-lookup"><span data-stu-id="b93fe-231">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="b93fe-232">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b93fe-232">Child elements</span></span>

|  <span data-ttu-id="b93fe-233">Élément</span><span class="sxs-lookup"><span data-stu-id="b93fe-233">Element</span></span> |  <span data-ttu-id="b93fe-234">Description</span><span class="sxs-lookup"><span data-stu-id="b93fe-234">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="b93fe-235">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-235">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="b93fe-236">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="b93fe-236">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="b93fe-237">CustomTab</span><span class="sxs-lookup"><span data-stu-id="b93fe-237">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="b93fe-238">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b93fe-238">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="b93fe-239">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="b93fe-239">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="b93fe-240">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique dans le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="b93fe-240">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="b93fe-241">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b93fe-241">Child elements</span></span>

|  <span data-ttu-id="b93fe-242">Élément</span><span class="sxs-lookup"><span data-stu-id="b93fe-242">Element</span></span> |  <span data-ttu-id="b93fe-243">Description</span><span class="sxs-lookup"><span data-stu-id="b93fe-243">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="b93fe-244">Group</span><span class="sxs-lookup"><span data-stu-id="b93fe-244">Group</span></span>](group.md) |  <span data-ttu-id="b93fe-245">Ajoute un groupe de boutons à la surface de commande.</span><span class="sxs-lookup"><span data-stu-id="b93fe-245">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="b93fe-246">Les éléments **ExtensionPoint** de ce type peuvent uniquement avoir un élément enfant, à savoir un élément **Group**.</span><span class="sxs-lookup"><span data-stu-id="b93fe-246">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="b93fe-247">Pour les éléments **Control** contenus dans ce point d’extension, l’attribut **xsi:type** doit avoir la valeur `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="b93fe-247">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="b93fe-248">Exemple</span><span class="sxs-lookup"><span data-stu-id="b93fe-248">Example</span></span>
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

### <a name="events"></a><span data-ttu-id="b93fe-249">Événements</span><span class="sxs-lookup"><span data-stu-id="b93fe-249">Events</span></span>

<span data-ttu-id="b93fe-250">Ce point d’extension ajoute un gestionnaire d’événements pour un événement spécifié.</span><span class="sxs-lookup"><span data-stu-id="b93fe-250">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="b93fe-251">Ce type d’élément est pris en charge par la version classique d’Outlook sur le Web, et en mode [aperçu](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sous Windows, Mac et Outlook moderne sur le Web.</span><span class="sxs-lookup"><span data-stu-id="b93fe-251">This element type is supported by classic Outlook on the web, and in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Windows, Mac, and modern Outlook on the web.</span></span> <span data-ttu-id="b93fe-252">Un abonnement Office 365 est également requis.</span><span class="sxs-lookup"><span data-stu-id="b93fe-252">An Office 365 subscription is also required.</span></span>

| <span data-ttu-id="b93fe-253">Élément</span><span class="sxs-lookup"><span data-stu-id="b93fe-253">Element</span></span> | <span data-ttu-id="b93fe-254">Description</span><span class="sxs-lookup"><span data-stu-id="b93fe-254">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="b93fe-255">Event</span><span class="sxs-lookup"><span data-stu-id="b93fe-255">Event</span></span>](event.md) |  <span data-ttu-id="b93fe-256">Indique l’événement et la fonction gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="b93fe-256">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="b93fe-257">Exemple d’événement ItemSend</span><span class="sxs-lookup"><span data-stu-id="b93fe-257">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="b93fe-258">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="b93fe-258">DetectedEntity</span></span>

<span data-ttu-id="b93fe-259">Ce point d’extension ajoute une activation de complément contextuel sur un type d’entité spécifié.</span><span class="sxs-lookup"><span data-stu-id="b93fe-259">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="b93fe-260">Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, `xsi:type`l’attribut doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="b93fe-260">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="b93fe-261">Ce type d’élément est disponible pour [les clients Outlook qui prennent en charge les ensembles de conditions requises 1.6 ou version ultérieure](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="b93fe-261">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="b93fe-262">Élément</span><span class="sxs-lookup"><span data-stu-id="b93fe-262">Element</span></span> |  <span data-ttu-id="b93fe-263">Description</span><span class="sxs-lookup"><span data-stu-id="b93fe-263">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="b93fe-264">Label</span><span class="sxs-lookup"><span data-stu-id="b93fe-264">Label</span></span>](#label) |  <span data-ttu-id="b93fe-265">Spécifie l’étiquette pour le complément dans la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="b93fe-265">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="b93fe-266">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="b93fe-266">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="b93fe-267">Spécifie l’URL de la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="b93fe-267">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="b93fe-268">Règle</span><span class="sxs-lookup"><span data-stu-id="b93fe-268">Rule</span></span>](rule.md) |  <span data-ttu-id="b93fe-269">Spécifie la ou les règles qui déterminent lorsqu’un complément s’active.</span><span class="sxs-lookup"><span data-stu-id="b93fe-269">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="b93fe-270">Étiquette</span><span class="sxs-lookup"><span data-stu-id="b93fe-270">Label</span></span>

<span data-ttu-id="b93fe-271">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="b93fe-271">Required.</span></span> <span data-ttu-id="b93fe-272">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="b93fe-272">The label of the group.</span></span> <span data-ttu-id="b93fe-273">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="b93fe-273">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="b93fe-274">Exigences relatives à la mise en surbrillance</span><span class="sxs-lookup"><span data-stu-id="b93fe-274">Highlight requirements</span></span>

<span data-ttu-id="b93fe-p117">Le seul moyen pour qu’un utilisateur puisse activer un complément contextuel consiste à interagir avec une entité en surbrillance. Les développeurs peuvent contrôler les entités qui sont mises en surbrillance à l’aide de l’attribut `Highlight` de l’élément `Rule` pour les types de règles `ItemHasKnownEntity` et `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="b93fe-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="b93fe-p118">Toutefois, il existe certaines limitations à connaître. Ces limitations sont en place pour vous assurer qu’il y aura toujours une entité en surbrillance dans les messages ou rendez-vous applicables pour permettre à l’utilisateur d’activer le complément.</span><span class="sxs-lookup"><span data-stu-id="b93fe-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="b93fe-279">Les types d’entité `EmailAddress` et `Url` ne peuvent pas être mis en surbrillance et par conséquent ne peuvent pas être utilisés pour activer un complément.</span><span class="sxs-lookup"><span data-stu-id="b93fe-279">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="b93fe-280">Si vous utilisez une seule règle, la valeur `Highlight` DOIT être définie sur `all`.</span><span class="sxs-lookup"><span data-stu-id="b93fe-280">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="b93fe-281">Si vous utilisez un type de règle `RuleCollection` avec `Mode="AND"` pour combiner plusieurs règles, au moins l’une des règles DOIT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="b93fe-281">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="b93fe-282">Si vous utilisez un type de règle `RuleCollection` avec `Mode="OR"` pour combiner plusieurs règles, toutes les règles DOIVENT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="b93fe-282">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="b93fe-283">Exemple d’événement DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="b93fe-283">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```
