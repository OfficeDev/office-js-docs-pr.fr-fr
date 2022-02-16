---
title: Élément Namespace dans le fichier manifest
description: L’élément Namespace définit l’espace de noms qu’une fonction personnalisée utilise dans Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: f9fddaca6ec8ce6128ae638c9b798efb06319ba0
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855624"
---
# <a name="namespace-element"></a>Élément Namespace

Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.

**Type de add-in :** Fonction personnalisée

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Taskpane 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Non  | Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md). Ne peut pas faire plus de 32 caractères. |

## <a name="child-elements"></a>Éléments enfants

Aucun

## <a name="example"></a>Exemple

```xml
<Namespace resid="namespace" />
```
