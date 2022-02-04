---
title: Élément Page dans le fichier manifeste
description: L’élément Page définit les paramètres de page HTML qu’une fonction personnalisée utilise dans Excel.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="page-element"></a>Élément Page

Définit les paramètres de la page HTML utilisés par une fonction personnalisée dans Excel.

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
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne contenant l’ID de ressource du fichier HTML utilisé par les fonctions personnalisées. |

## <a name="example"></a>Exemple

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
