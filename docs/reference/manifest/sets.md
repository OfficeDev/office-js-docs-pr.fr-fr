---
title: Élément Sets dans le fichier manifeste
description: L’élément Sets spécifie l’ensemble minimal d’API JavaScript Office dont votre application Office a besoin pour être activé par Office ou pour remplacer les paramètres de manifeste de base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: df0cf686fe213a51321595a000438ca2a411f2c7
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222142"
---
# <a name="sets-element"></a>Élément Sets

La signification de cet élément dépend de l’endroit où il est utilisé dans le manifeste.

## <a name="in-the-base-manifest"></a>Dans le manifeste de base

Lorsqu’il est utilisé dans le manifeste de base (c’est-à-dire que l’élément **Requirements** parent est un enfant direct [d’OfficeApp](officeapp.md)), l’élément **Sets** spécifie le sous-ensemble minimal des conditions requises de l’API JavaScript Office [(ensembles](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)de conditions requises) dont votre application Office a besoin pour être activée par Office.

**Type de complément :** Application de contenu, de volet Office, de messagerie

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>En tant que petit-enfant d’un élément VersionOverrides

Spécifie l’ensemble minimal de conditions requises[](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)de l’API JavaScript Office (ensembles de conditions requises) qui doivent être pris en charge par la version et la plateforme Office (telles que Windows, Mac, web et iOS ou iPad) pour que [versionOverrides](versionoverrides.md) prenne effet.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans ces schémas VersionOverrides**:

- Identique à [l’élément Requirements](requirements.md) parent.

**Associés à ces ensembles de conditions requises**:

- Identique à [l’élément Requirements](requirements.md) parent.

## <a name="syntax"></a>Syntaxe

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Contenu dans

[Configuration requise](requirements.md)

## <a name="can-contain"></a>Peut contenir

[Ensemble](set.md)

## <a name="attributes"></a>Attributs

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|chaîne|facultatif|Spécifie la valeur **d’attribut MinVersion** par défaut pour tous les éléments [Set](set.md) enfants. La valeur par défaut est « 1.1 ».|

## <a name="remarks"></a>Remarques

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et l’attribut **DefaultMinVersion** de l’élément **Sets,** voir Spécifier quelles versions et plateformes Office peuvent héberger votre [add-in.](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)

