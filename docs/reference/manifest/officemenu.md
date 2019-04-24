---
title: Élément OfficeMenu dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 20d020b8ab826049ef0271cbdb8d51201ee88ab4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452017"
---
# <a name="officemenu-element"></a>Élément OfficeMenu

Définit un ensemble d’options à ajouter au menu contextuel Office. S’applique aux compléments Word, Excel, PowerPoint et OneNote.

## <a name="attributes"></a>Attributs

| Attribut            | Obligatoire | Description                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | Oui      | Type d’OfficeMenu défini.|

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Control](#control)    | Oui |  Ensemble d’un ou de plusieurs objets Control  |

## <a name="xsitype"></a>xsi:type

Indique un menu prédéfini de l’application cliente Office sur laquelle ajouter ce complément Office.

- `ContextMenuText` -  Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur ouvre le menu contextuel (clique dessus avec le bouton droit de la souris) du texte sélectionné. S’applique à Word, Excel, PowerPoint et OneNote.
- `ContextMenuCell` -  Affiche l’élément dans le menu contextuel lorsque l’utilisateur ouvre le menu contextuel (clique avec le bouton droit de la souris) dans une cellule de la feuille de calcul. S’applique à Excel. 

## <a name="control"></a>Contrôle

Chaque élément **OfficeMenu** requiert une ou plusieurs options de [menu](control.md#menu-dropdown-button-controls). 

## <a name="example"></a>Exemple

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
