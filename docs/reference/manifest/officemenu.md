---
title: Élément OfficeMenu dans le fichier manifeste
description: L’élément OfficeMenu définit une collection de contrôles à ajouter au menu context Office de travail.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="officemenu-element"></a>Élément OfficeMenu

Définit un ensemble d’options à ajouter au menu contextuel Office. S’applique aux compléments Word, Excel, PowerPoint et OneNote.

**Type de complément :** volet Office

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Taskpane 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

## <a name="attributes"></a>Attributs

| Attribut            | Obligatoire | Description                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | Oui      | Type d’OfficeMenu défini.|

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
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
    <Control xsi:type="Menu" id="Contoso.myMenu">
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
