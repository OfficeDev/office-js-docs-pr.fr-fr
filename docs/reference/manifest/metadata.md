---
title: Élément de métadonnées dans le fichier manifest
description: L’élément Metadata définit les paramètres de métadonnées qu’une fonction personnalisée utilise dans Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52938155442bb5424a170634d1324de77de2b788
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855533"
---
# <a name="metadata-element"></a>Élément de métadonnées

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

|  Élément  |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne contenant l’ID de ressource du fichier JSON utilisé par les fonctions personnalisées. |

## <a name="example"></a>Exemple

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
