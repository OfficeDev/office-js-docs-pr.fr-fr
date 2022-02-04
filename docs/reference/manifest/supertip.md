---
title: Élément Supertip dans le fichier manifest
description: L’élément Supertip définit une boîte à outils enrichie (titre et description).
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="supertip"></a>Supertip

Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Taskpane 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) lorsque le parent **VersionOverrides** est de type Taskpane 1.0.
- [Boîte aux lettres 1.3 lorsque](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) le parent **VersionOverrides** est de type Mail 1.0.
- [Boîte aux lettres 1.5 lorsque](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) le parent **VersionOverrides** est de type Mail 1.1.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
| [Titre](#title) | Oui | Texte de l’info-bulle. |
| [Description](#description) | Oui | Description de l’info-bulle.<br>**Remarque** : (Outlook) seuls Windows clients Mac et mac sont pris en charge. |

### <a name="title"></a>Titre

Obligatoire. Texte de la propriété SuperTip. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md).

### <a name="description"></a>Description

Obligatoire. Description de l’info-bulle. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **String** dans l’élément **LongStrings** dans l’élément [Resources](resources.md).

> [!NOTE]
> Par Outlook, seuls Windows clients Mac et les clients Mac supportent **l’élément Description**.

## <a name="example"></a>Exemple

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
