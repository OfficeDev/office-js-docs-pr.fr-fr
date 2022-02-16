---
title: Élément Script dans le fichier manifeste
description: L’élément Script définit les paramètres de script qu’une fonction personnalisée utilise dans Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0f32314912dd66d8578750bf4818af8483c8ef36
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855540"
---
# <a name="script-element"></a>Élément Script

Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.

**Type de add-in :** Fonction personnalisée

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Taskpane 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Attributs

Aucun

## <a name="child-elements"></a>Éléments enfants

|Éléments  |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne avec l’ID de ressource du fichier JavaScript utilisé par les fonctions personnalisées.|

## <a name="example"></a>Exemple

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
