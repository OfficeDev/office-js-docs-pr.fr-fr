---
title: Élément SourceLocation pour les fonctions personnalisées dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="sourcelocation-element-custom-functions"></a>Élément SourceLocation (fonctions personnalisées)

Définit l’emplacement d’une ressource requise par les éléments **Script** ou **Page** utilisés par les fonctions personnalisées dans Excel.

> [!IMPORTANT]
> Cet article fait uniquement référence à **sourcelocation** qui est un enfant des éléments **Page** ou **Script** . Pour [plus d’informations sur](sourcelocation.md) l’élément SourceLocation du manifeste de base, voir **SourceLocation** .

**Type de add-in :** Fonction personnalisée

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Taskpane 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Attributs

| Attribut | Obligatoire | Description                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Oui      | Nom d’une ressource d’URL définie dans la section **Ressources** du manifeste. Ne peut pas faire plus de 32 caractères. |

## <a name="child-elements"></a>Éléments enfants

Aucun

## <a name="example"></a>Exemple

```xml
<SourceLocation resid="pageURL"/>
```
