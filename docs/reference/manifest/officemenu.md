---
title: Élément OfficeMenu dans le fichier manifeste
description: L’élément OfficeMenu définit une collection de contrôles à ajouter au menu contextuel Office.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d181e0c6f489997a149b9713bdc257f4a2baeb16
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641441"
---
# <a name="officemenu-element"></a><span data-ttu-id="d135b-103">Élément OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="d135b-103">OfficeMenu element</span></span>

<span data-ttu-id="d135b-p101">Définit un ensemble d’options à ajouter au menu contextuel Office. S’applique aux compléments Word, Excel, PowerPoint et OneNote.</span><span class="sxs-lookup"><span data-stu-id="d135b-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="d135b-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="d135b-106">Attributes</span></span>

| <span data-ttu-id="d135b-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="d135b-107">Attribute</span></span>            | <span data-ttu-id="d135b-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="d135b-108">Required</span></span> | <span data-ttu-id="d135b-109">Description</span><span class="sxs-lookup"><span data-stu-id="d135b-109">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="d135b-110">xsi:type</span><span class="sxs-lookup"><span data-stu-id="d135b-110">xsi:type</span></span>](#xsitype) | <span data-ttu-id="d135b-111">Oui</span><span class="sxs-lookup"><span data-stu-id="d135b-111">Yes</span></span>      | <span data-ttu-id="d135b-112">Type d’OfficeMenu défini.</span><span class="sxs-lookup"><span data-stu-id="d135b-112">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="d135b-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="d135b-113">Child elements</span></span>

|  <span data-ttu-id="d135b-114">Élément</span><span class="sxs-lookup"><span data-stu-id="d135b-114">Element</span></span> |  <span data-ttu-id="d135b-115">Requis</span><span class="sxs-lookup"><span data-stu-id="d135b-115">Required</span></span>  |  <span data-ttu-id="d135b-116">Description</span><span class="sxs-lookup"><span data-stu-id="d135b-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d135b-117">Control</span><span class="sxs-lookup"><span data-stu-id="d135b-117">Control</span></span>](#control)    | <span data-ttu-id="d135b-118">Oui</span><span class="sxs-lookup"><span data-stu-id="d135b-118">Yes</span></span> |  <span data-ttu-id="d135b-119">Ensemble d’un ou de plusieurs objets Control</span><span class="sxs-lookup"><span data-stu-id="d135b-119">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="d135b-120">xsi:type</span><span class="sxs-lookup"><span data-stu-id="d135b-120">xsi:type</span></span>

<span data-ttu-id="d135b-121">Indique un menu prédéfini de l’application cliente Office sur laquelle ajouter ce complément Office.</span><span class="sxs-lookup"><span data-stu-id="d135b-121">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="d135b-p102">`ContextMenuText` -  Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur ouvre le menu contextuel (clique dessus avec le bouton droit de la souris) du texte sélectionné. S’applique à Word, Excel, PowerPoint et OneNote.</span><span class="sxs-lookup"><span data-stu-id="d135b-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="d135b-p103">`ContextMenuCell` -  Affiche l’élément dans le menu contextuel lorsque l’utilisateur ouvre le menu contextuel (clique avec le bouton droit de la souris) dans une cellule de la feuille de calcul. S’applique à Excel.</span><span class="sxs-lookup"><span data-stu-id="d135b-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span>

## <a name="control"></a><span data-ttu-id="d135b-126">Contrôle</span><span class="sxs-lookup"><span data-stu-id="d135b-126">Control</span></span>

<span data-ttu-id="d135b-127">Chaque élément **OfficeMenu** requiert une ou plusieurs options de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="d135b-127">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="d135b-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="d135b-128">Example</span></span>

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />
          </Action>
        </Item>
      </Items>
    </Control>
</OfficeMenu>
```
