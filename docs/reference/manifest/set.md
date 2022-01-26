---
title: Élément Set dans le fichier manifeste
description: L’élément Set spécifie un ensemble de conditions requises de l’API JavaScript Office requise par votre Office pour être activé par Office ou pour remplacer les paramètres de manifeste de base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 55e1b25765bfbe53108bc9201c0c851c6ef9161d
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222233"
---
# <a name="set-element"></a>Élément Set

La signification de cet élément dépend de l’endroit où il est utilisé dans le manifeste.

## <a name="in-the-base-manifest"></a>Dans le manifeste de base

Lorsqu’il est utilisé dans le manifeste de base (autrement dit, si l’élément **Requirements** est un enfant direct d’OfficeApp ), l’élément **Set** spécifie un ensemble de conditions requises de l’API JavaScript Office dont votre Office a besoin pour être activé par Office. [](officeapp.md) [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)

**Type de complément :** Application de contenu, de volet Office, de messagerie

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>En tant qu’arrière-petit-enfant d’un élément VersionOverrides

Spécifie un [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) ensemble de conditions requises à partir de l’API JavaScript Office qui doit être pris en charge par la version et la plateforme Office (telles que Windows, Mac, web et iOS ou iPad) pour que [versionOverrides](versionoverrides.md) prenne effet.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans ces schémas VersionOverrides**:

- Identique à [l’élément Requirements.](requirements.md)

**Associés à ces ensembles de conditions requises**:

- Identique à [l’élément Requirements.](requirements.md)

## <a name="syntax"></a>Syntaxe

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Contenu dans

[Ensembles](sets.md)

## <a name="attributes"></a>Attributs

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|Nom|string|obligatoire|Nom d’un [ensemble de conditions requises](../../develop/office-versions-and-requirement-sets.md).|
|MinVersion|chaîne|facultatif|Spécifie la version minimale de l’ensemble d’API requis par votre complément. Remplace la valeur de **DefaultMinVersion,** si elle est spécifiée dans l’élément [Sets](sets.md) parent.|

## <a name="remarks"></a>Remarques

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et l’attribut **DefaultMinVersion** de l’élément **Sets,** voir Spécifier quelles versions et plateformes Office peuvent héberger votre [add-in.](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)

