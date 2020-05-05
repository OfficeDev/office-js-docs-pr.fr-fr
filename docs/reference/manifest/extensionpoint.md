---
title: Élément Extension dans le fichier manifeste
description: Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.
ms.date: 05/04/2020
localization_priority: Normal
ms.openlocfilehash: ede99ad73beb1e4a46c9b08188ca79efb556acb0
ms.sourcegitcommit: 800dacf0399465318489c9d949e259b5cf0f81ca
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/05/2020
ms.locfileid: "44022175"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="1de49-103">Élément ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="1de49-103">ExtensionPoint element</span></span>

 <span data-ttu-id="1de49-104">Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="1de49-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="1de49-105">L’élément **ExtensionPoint** est un élément enfant de [AllFormFactors](allformfactors.md) ou [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="1de49-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="1de49-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="1de49-106">Attributes</span></span>

|  <span data-ttu-id="1de49-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="1de49-107">Attribute</span></span>  |  <span data-ttu-id="1de49-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="1de49-108">Required</span></span>  |  <span data-ttu-id="1de49-109">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1de49-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="1de49-110">**xsi:type**</span></span>  |  <span data-ttu-id="1de49-111">Oui</span><span class="sxs-lookup"><span data-stu-id="1de49-111">Yes</span></span>  | <span data-ttu-id="1de49-112">Type de point d’extension défini.</span><span class="sxs-lookup"><span data-stu-id="1de49-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="1de49-113">Points d’extension pour Excel uniquement</span><span class="sxs-lookup"><span data-stu-id="1de49-113">Extension points for Excel only</span></span>

- <span data-ttu-id="1de49-114">**CustomFunctions** – fonction personnalisée écrite en JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="1de49-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="1de49-115">[L’exemple de code XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) montre comment utiliser l’élément **ExtensionPoint** avec la valeur d’attribut **CustomFunctions** et les éléments enfants à utiliser.</span><span class="sxs-lookup"><span data-stu-id="1de49-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="1de49-116">Points d’extension pour les commandes de complément Word, Excel, PowerPoint et OneNote</span><span class="sxs-lookup"><span data-stu-id="1de49-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="1de49-117">**PrimaryCommandSurface** : ruban dans Office.</span><span class="sxs-lookup"><span data-stu-id="1de49-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="1de49-118">**ContextMenu** : menu contextuel qui apparaît lorsque vous cliquez avec le bouton droit de la souris dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="1de49-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="1de49-119">Les exemples suivants montrent comment utiliser l’élément **ExtensionPoint** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.</span><span class="sxs-lookup"><span data-stu-id="1de49-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1de49-p102">Pour les éléments qui contiennent un attribut ID, assurez-vous que vous indiquez un ID unique. Nous vous recommandons d’utiliser le nom de votre organisation, ainsi que votre ID. Par exemple, utilisez le format suivant. <CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="1de49-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="1de49-123">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1de49-123">Child elements</span></span>
 
|<span data-ttu-id="1de49-124">**Élément**</span><span class="sxs-lookup"><span data-stu-id="1de49-124">**Element**</span></span>|<span data-ttu-id="1de49-125">**Description**</span><span class="sxs-lookup"><span data-stu-id="1de49-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="1de49-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="1de49-126">**CustomTab**</span></span>|<span data-ttu-id="1de49-p103">Obligatoire si vous souhaitez ajouter un onglet personnalisé au ruban (à l’aide de **PrimaryCommandSurface**). Si vous utilisez l’élément **CustomTab**, vous ne pouvez pas utiliser l’élément **OfficeTab**. L’attribut **id** est obligatoire. </span><span class="sxs-lookup"><span data-stu-id="1de49-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="1de49-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="1de49-130">**OfficeTab**</span></span>|<span data-ttu-id="1de49-131">Obligatoire pour étendre un onglet du ruban Office par défaut (en utilisant **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="1de49-131">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="1de49-132">Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**.</span><span class="sxs-lookup"><span data-stu-id="1de49-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="1de49-133">Pour plus d’informations, voir [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="1de49-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="1de49-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="1de49-134">**OfficeMenu**</span></span>|<span data-ttu-id="1de49-p105">Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : </span><span class="sxs-lookup"><span data-stu-id="1de49-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="1de49-p106">- **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. </span><span class="sxs-lookup"><span data-stu-id="1de49-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="1de49-p107">- **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="1de49-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="1de49-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="1de49-141">**Group**</span></span>|<span data-ttu-id="1de49-p108">Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est obligatoire. Il s’agit d’une chaîne avec un maximum de 125 caractères. </span><span class="sxs-lookup"><span data-stu-id="1de49-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="1de49-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="1de49-145">**Label**</span></span>|<span data-ttu-id="1de49-p109">Obligatoire. L’étiquette du groupe. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="1de49-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="1de49-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="1de49-150">**Icon**</span></span>|<span data-ttu-id="1de49-p110">Obligatoire. Spécifie l’icône du groupe à utiliser sur de petits appareils, ou lorsqu’un nombre trop important de boutons est affiché. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image**. **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément **Ressources**. L’attribut **size** donne la taille, en pixels, de l’image. Trois tailles d’images sont obligatoires : 16, 32 et 80. 5 tailles facultatives sont également prises en charge : 20, 24, 40, 48 et 64. </span><span class="sxs-lookup"><span data-stu-id="1de49-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="1de49-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="1de49-158">**Tooltip**</span></span>|<span data-ttu-id="1de49-p111">Facultatif. Info-bulle du groupe. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Chaîne**. **Chaîne** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément **Ressources**. </span><span class="sxs-lookup"><span data-stu-id="1de49-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="1de49-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="1de49-163">**Control**</span></span>|<span data-ttu-id="1de49-164">Chaque groupe requiert au moins un contrôle.</span><span class="sxs-lookup"><span data-stu-id="1de49-164">Each group requires at least one control.</span></span> <span data-ttu-id="1de49-165">Un élément **Control** peut être de type **Button** ou **Menu**.</span><span class="sxs-lookup"><span data-stu-id="1de49-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="1de49-166">Utilisez **Menu** pour spécifier une liste déroulante de contrôles de bouton.</span><span class="sxs-lookup"><span data-stu-id="1de49-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="1de49-167">Actuellement, seuls les boutons et les menus sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="1de49-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="1de49-168">Pour plus d’informations, reportez-vous aux sections [Contrôles de bouton](control.md#button-control) et [Contrôles de menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="1de49-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="1de49-169">**Remarque :**  Pour faciliter la résolution des problèmes, nous vous recommandons d’ajouter un élément **Control** et les éléments enfants **Resources** associés un par un.</span><span class="sxs-lookup"><span data-stu-id="1de49-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="1de49-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="1de49-170">**Script**</span></span>|<span data-ttu-id="1de49-171">Liens vers le fichier JavaScript avec la définition de la fonction personnalisée et le code d’inscription.</span><span class="sxs-lookup"><span data-stu-id="1de49-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="1de49-172">Cet élément n’est pas utilisé dans l’aperçu pour les développeurs.</span><span class="sxs-lookup"><span data-stu-id="1de49-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="1de49-173">À la place, la page HTML est responsable du chargement de tous les fichiers JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1de49-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="1de49-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="1de49-174">**Page**</span></span>|<span data-ttu-id="1de49-175">Liens vers la page HTML de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1de49-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="1de49-176">Points d’extension pour Outlook</span><span class="sxs-lookup"><span data-stu-id="1de49-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="1de49-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="1de49-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="1de49-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="1de49-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="1de49-181">[Module](#module) (peut uniquement être utilisé dans [DesktopFormFactor](desktopformfactor.md).)</span><span class="sxs-lookup"><span data-stu-id="1de49-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="1de49-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="1de49-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface-preview)
- [<span data-ttu-id="1de49-184">Événements</span><span class="sxs-lookup"><span data-stu-id="1de49-184">Events</span></span>](#events)
- [<span data-ttu-id="1de49-185">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="1de49-185">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="1de49-186">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-186">MessageReadCommandSurface</span></span>

<span data-ttu-id="1de49-p114">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique. Dans l’application de bureau Outlook, cela apparaît dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="1de49-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="1de49-189">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1de49-189">Child elements</span></span>

|  <span data-ttu-id="1de49-190">Élément</span><span class="sxs-lookup"><span data-stu-id="1de49-190">Element</span></span> |  <span data-ttu-id="1de49-191">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-191">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1de49-192">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1de49-192">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1de49-193">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="1de49-193">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1de49-194">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1de49-194">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1de49-195">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="1de49-195">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="1de49-196">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1de49-196">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="1de49-197">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="1de49-197">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="1de49-198">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-198">MessageComposeCommandSurface</span></span>

<span data-ttu-id="1de49-199">Ce point d’extension place des boutons sur le ruban pour les compléments à l’aide du formulaire de composition de messagerie.</span><span class="sxs-lookup"><span data-stu-id="1de49-199">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="1de49-200">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1de49-200">Child elements</span></span>

|  <span data-ttu-id="1de49-201">Élément</span><span class="sxs-lookup"><span data-stu-id="1de49-201">Element</span></span> |  <span data-ttu-id="1de49-202">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-202">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1de49-203">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1de49-203">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1de49-204">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="1de49-204">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1de49-205">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1de49-205">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1de49-206">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="1de49-206">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="1de49-207">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1de49-207">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="1de49-208">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="1de49-208">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="1de49-209">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-209">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="1de49-210">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention de l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="1de49-210">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="1de49-211">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1de49-211">Child elements</span></span>

|  <span data-ttu-id="1de49-212">Élément</span><span class="sxs-lookup"><span data-stu-id="1de49-212">Element</span></span> |  <span data-ttu-id="1de49-213">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-213">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1de49-214">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1de49-214">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1de49-215">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="1de49-215">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1de49-216">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1de49-216">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1de49-217">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="1de49-217">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="1de49-218">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1de49-218">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="1de49-219">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="1de49-219">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="1de49-220">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-220">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="1de49-221">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention du participant à la réunion.</span><span class="sxs-lookup"><span data-stu-id="1de49-221">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="1de49-222">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1de49-222">Child elements</span></span>

|  <span data-ttu-id="1de49-223">Élément</span><span class="sxs-lookup"><span data-stu-id="1de49-223">Element</span></span> |  <span data-ttu-id="1de49-224">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-224">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1de49-225">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1de49-225">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1de49-226">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="1de49-226">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1de49-227">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1de49-227">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1de49-228">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="1de49-228">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="1de49-229">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1de49-229">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="1de49-230">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="1de49-230">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="1de49-231">Module</span><span class="sxs-lookup"><span data-stu-id="1de49-231">Module</span></span>

<span data-ttu-id="1de49-232">Ce point d’extension place des boutons sur le ruban pour l’extension de module.</span><span class="sxs-lookup"><span data-stu-id="1de49-232">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="1de49-233">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1de49-233">Child elements</span></span>

|  <span data-ttu-id="1de49-234">Élément</span><span class="sxs-lookup"><span data-stu-id="1de49-234">Element</span></span> |  <span data-ttu-id="1de49-235">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-235">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1de49-236">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1de49-236">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1de49-237">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="1de49-237">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1de49-238">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1de49-238">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1de49-239">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="1de49-239">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="1de49-240">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1de49-240">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="1de49-241">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique dans le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="1de49-241">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="1de49-242">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1de49-242">Child elements</span></span>

|  <span data-ttu-id="1de49-243">Élément</span><span class="sxs-lookup"><span data-stu-id="1de49-243">Element</span></span> |  <span data-ttu-id="1de49-244">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-244">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1de49-245">Group</span><span class="sxs-lookup"><span data-stu-id="1de49-245">Group</span></span>](group.md) |  <span data-ttu-id="1de49-246">Ajoute un groupe de boutons à la surface de commande.</span><span class="sxs-lookup"><span data-stu-id="1de49-246">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="1de49-247">Les éléments **ExtensionPoint** de ce type peuvent uniquement avoir un élément enfant, à savoir un élément **Group**.</span><span class="sxs-lookup"><span data-stu-id="1de49-247">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="1de49-248">Pour les éléments **Control** contenus dans ce point d’extension, l’attribut **xsi:type** doit avoir la valeur `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="1de49-248">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="1de49-249">Exemple</span><span class="sxs-lookup"><span data-stu-id="1de49-249">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a><span data-ttu-id="1de49-250">MobileOnlineMeetingCommandSurface (aperçu)</span><span class="sxs-lookup"><span data-stu-id="1de49-250">MobileOnlineMeetingCommandSurface (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="1de49-251">Ce point d’extension est uniquement pris en charge en [Aperçu](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur Android avec un abonnement Office 365.</span><span class="sxs-lookup"><span data-stu-id="1de49-251">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="1de49-252">Ce point d’extension place un bouton bascule mode-approprié dans la surface de commande pour un rendez-vous dans le facteur de forme mobile.</span><span class="sxs-lookup"><span data-stu-id="1de49-252">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="1de49-253">Un organisateur de réunion peut créer une réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="1de49-253">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="1de49-254">Un participant peut ensuite participer à la réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="1de49-254">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="1de49-255">Pour en savoir plus sur ce scénario, consultez l’article [créer un complément Outlook Mobile pour un fournisseur de réunions en ligne](../../outlook/online-meeting.md) .</span><span class="sxs-lookup"><span data-stu-id="1de49-255">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="1de49-256">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1de49-256">Child elements</span></span>

|  <span data-ttu-id="1de49-257">Élément</span><span class="sxs-lookup"><span data-stu-id="1de49-257">Element</span></span> |  <span data-ttu-id="1de49-258">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-258">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1de49-259">Control</span><span class="sxs-lookup"><span data-stu-id="1de49-259">Control</span></span>](control.md) |  <span data-ttu-id="1de49-260">Ajoute un bouton à la surface de commande.</span><span class="sxs-lookup"><span data-stu-id="1de49-260">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="1de49-261">`ExtensionPoint`les éléments de ce type ne peuvent avoir qu’un seul élément `Control` enfant : un élément.</span><span class="sxs-lookup"><span data-stu-id="1de49-261">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="1de49-262">L' `Control` `xsi:type` attribut de l’élément contenu dans ce point d’extension doit être `MobileButton`défini sur.</span><span class="sxs-lookup"><span data-stu-id="1de49-262">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="1de49-263">Les `Icon` images doivent être en nuances de gris à `#919191` l’aide de code hexadécimal ou de leur équivalent dans d' [autres formats de couleur](https://convertingcolors.com/hex-color-919191.html).</span><span class="sxs-lookup"><span data-stu-id="1de49-263">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="1de49-264">Exemple</span><span class="sxs-lookup"><span data-stu-id="1de49-264">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="1de49-265">Événements</span><span class="sxs-lookup"><span data-stu-id="1de49-265">Events</span></span>

<span data-ttu-id="1de49-266">Ce point d’extension ajoute un gestionnaire d’événements pour un événement spécifié.</span><span class="sxs-lookup"><span data-stu-id="1de49-266">This extension point adds an event handler for a specified event.</span></span>

| <span data-ttu-id="1de49-267">Élément</span><span class="sxs-lookup"><span data-stu-id="1de49-267">Element</span></span> | <span data-ttu-id="1de49-268">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-268">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1de49-269">Event</span><span class="sxs-lookup"><span data-stu-id="1de49-269">Event</span></span>](event.md) |  <span data-ttu-id="1de49-270">Indique l’événement et la fonction gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="1de49-270">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="1de49-271">Exemple d’événement ItemSend</span><span class="sxs-lookup"><span data-stu-id="1de49-271">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="1de49-272">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="1de49-272">DetectedEntity</span></span>

<span data-ttu-id="1de49-273">Ce point d’extension ajoute une activation de complément contextuel sur un type d’entité spécifié.</span><span class="sxs-lookup"><span data-stu-id="1de49-273">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="1de49-274">Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, `xsi:type`l’attribut doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="1de49-274">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="1de49-275">Ce type d’élément est disponible pour [les clients Outlook qui prennent en charge les ensembles de conditions requises 1.6 ou version ultérieure](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="1de49-275">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="1de49-276">Élément</span><span class="sxs-lookup"><span data-stu-id="1de49-276">Element</span></span> |  <span data-ttu-id="1de49-277">Description</span><span class="sxs-lookup"><span data-stu-id="1de49-277">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1de49-278">Label</span><span class="sxs-lookup"><span data-stu-id="1de49-278">Label</span></span>](#label) |  <span data-ttu-id="1de49-279">Spécifie l’étiquette pour le complément dans la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="1de49-279">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="1de49-280">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1de49-280">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="1de49-281">Spécifie l’URL de la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="1de49-281">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="1de49-282">Règle</span><span class="sxs-lookup"><span data-stu-id="1de49-282">Rule</span></span>](rule.md) |  <span data-ttu-id="1de49-283">Spécifie la ou les règles qui déterminent lorsqu’un complément s’active.</span><span class="sxs-lookup"><span data-stu-id="1de49-283">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="1de49-284">Étiquette</span><span class="sxs-lookup"><span data-stu-id="1de49-284">Label</span></span>

<span data-ttu-id="1de49-285">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="1de49-285">Required.</span></span> <span data-ttu-id="1de49-286">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="1de49-286">The label of the group.</span></span> <span data-ttu-id="1de49-287">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="1de49-287">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="1de49-288">Exigences relatives à la mise en surbrillance</span><span class="sxs-lookup"><span data-stu-id="1de49-288">Highlight requirements</span></span>

<span data-ttu-id="1de49-p117">Le seul moyen pour qu’un utilisateur puisse activer un complément contextuel consiste à interagir avec une entité en surbrillance. Les développeurs peuvent contrôler les entités qui sont mises en surbrillance à l’aide de l’attribut `Highlight` de l’élément `Rule` pour les types de règles `ItemHasKnownEntity` et `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="1de49-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="1de49-p118">Toutefois, il existe certaines limitations à connaître. Ces limitations sont en place pour vous assurer qu’il y aura toujours une entité en surbrillance dans les messages ou rendez-vous applicables pour permettre à l’utilisateur d’activer le complément.</span><span class="sxs-lookup"><span data-stu-id="1de49-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="1de49-293">Les types d’entité `EmailAddress` et `Url` ne peuvent pas être mis en surbrillance et par conséquent ne peuvent pas être utilisés pour activer un complément.</span><span class="sxs-lookup"><span data-stu-id="1de49-293">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="1de49-294">Si vous utilisez une seule règle, la valeur `Highlight` DOIT être définie sur `all`.</span><span class="sxs-lookup"><span data-stu-id="1de49-294">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="1de49-295">Si vous utilisez un type de règle `RuleCollection` avec `Mode="AND"` pour combiner plusieurs règles, au moins l’une des règles DOIT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="1de49-295">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="1de49-296">Si vous utilisez un type de règle `RuleCollection` avec `Mode="OR"` pour combiner plusieurs règles, toutes les règles DOIVENT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="1de49-296">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="1de49-297">Exemple d’événement DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="1de49-297">DetectedEntity event example</span></span>

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
