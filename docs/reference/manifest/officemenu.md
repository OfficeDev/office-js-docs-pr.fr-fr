---
title: Élément OfficeMenu dans le fichier manifeste
description: L’élément OfficeMenu définit une collection de contrôles à ajouter au menu contextuel Office.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 89503533f7310898a420eb805d5fd66f096ad5f2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718047"
---
# <a name="officemenu-element"></a><span data-ttu-id="a166e-103">Élément OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="a166e-103">OfficeMenu element</span></span>

<span data-ttu-id="a166e-p101">Définit un ensemble d’options à ajouter au menu contextuel Office. S’applique aux compléments Word, Excel, PowerPoint et OneNote.</span><span class="sxs-lookup"><span data-stu-id="a166e-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="a166e-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="a166e-106">Attributes</span></span>

| <span data-ttu-id="a166e-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="a166e-107">Attribute</span></span>            | <span data-ttu-id="a166e-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a166e-108">Required</span></span> | <span data-ttu-id="a166e-109">Description</span><span class="sxs-lookup"><span data-stu-id="a166e-109">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="a166e-110">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a166e-110">xsi:type</span></span>](#xsitype) | <span data-ttu-id="a166e-111">Oui</span><span class="sxs-lookup"><span data-stu-id="a166e-111">Yes</span></span>      | <span data-ttu-id="a166e-112">Type d’OfficeMenu défini.</span><span class="sxs-lookup"><span data-stu-id="a166e-112">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="a166e-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="a166e-113">Child elements</span></span>

|  <span data-ttu-id="a166e-114">Élément</span><span class="sxs-lookup"><span data-stu-id="a166e-114">Element</span></span> |  <span data-ttu-id="a166e-115">Requis</span><span class="sxs-lookup"><span data-stu-id="a166e-115">Required</span></span>  |  <span data-ttu-id="a166e-116">Description</span><span class="sxs-lookup"><span data-stu-id="a166e-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a166e-117">Control</span><span class="sxs-lookup"><span data-stu-id="a166e-117">Control</span></span>](#control)    | <span data-ttu-id="a166e-118">Oui</span><span class="sxs-lookup"><span data-stu-id="a166e-118">Yes</span></span> |  <span data-ttu-id="a166e-119">Ensemble d’un ou de plusieurs objets Control</span><span class="sxs-lookup"><span data-stu-id="a166e-119">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="a166e-120">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a166e-120">xsi:type</span></span>

<span data-ttu-id="a166e-121">Indique un menu prédéfini de l’application cliente Office sur laquelle ajouter ce complément Office.</span><span class="sxs-lookup"><span data-stu-id="a166e-121">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="a166e-p102">`ContextMenuText` -  Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur ouvre le menu contextuel (clique dessus avec le bouton droit de la souris) du texte sélectionné. S’applique à Word, Excel, PowerPoint et OneNote.</span><span class="sxs-lookup"><span data-stu-id="a166e-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="a166e-p103">`ContextMenuCell` -  Affiche l’élément dans le menu contextuel lorsque l’utilisateur ouvre le menu contextuel (clique avec le bouton droit de la souris) dans une cellule de la feuille de calcul. S’applique à Excel.</span><span class="sxs-lookup"><span data-stu-id="a166e-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="a166e-126">Contrôle</span><span class="sxs-lookup"><span data-stu-id="a166e-126">Control</span></span>

<span data-ttu-id="a166e-127">Chaque élément **OfficeMenu** requiert une ou plusieurs options de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="a166e-127">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="a166e-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="a166e-128">Example</span></span>

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
