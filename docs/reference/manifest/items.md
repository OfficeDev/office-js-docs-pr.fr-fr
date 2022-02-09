---
title: Élément Items dans le fichier manifeste
description: Spécifie les éléments d’un menu.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2249bc55db662a36cf3986ebb0b90353237d4985
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467898"
---
# <a name="items-element"></a>Élément Items

Spécifie les éléments d’un menu.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) lorsque le parent **VersionOverrides** est de type Taskpane 1.0.
- [Boîte aux lettres 1.3 lorsque](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) le parent **VersionOverrides** est de type Mail 1.0.
- [Boîte aux lettres 1.5 lorsque](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) le parent **VersionOverrides** est de type Mail 1.1.

## <a name="syntax"></a>Syntaxe

```XML
<Items>
...  
</Items>  
```

## <a name="contained-in"></a>Contenu dans

[Élément Control de type Menu](control-menu.md)

## <a name="must-contain"></a>Doit contenir

[Item](item.md)

## <a name="examples"></a>Exemples

Pour obtenir des exemples, [voir Contrôle de type Menu](control-menu.md).