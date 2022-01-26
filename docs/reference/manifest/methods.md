---
title: Élément Methods dans le fichier manifeste
description: L’élément Methods spécifie la liste des méthodes de l’API JavaScript Office dont votre application Office a besoin pour être activée par Office ou pour remplacer les paramètres de manifeste de base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4c39c6363cd33e103cf40c0f7f047fa694db1411
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222275"
---
# <a name="methods-element"></a>Élément Methods

La signification de cet élément dépend de l’endroit où il est utilisé dans le manifeste.

## <a name="in-the-base-manifest"></a>Dans le manifeste de base

Lorsqu’il est utilisé dans le manifeste de base (c’est-à-dire que l’élément **Requirements** parent est un enfant direct [d’OfficeApp](officeapp.md)), l’élément **Methods** spécifie la liste des méthodes d’API JavaScript Office dont votre application Office a besoin pour être activée par Office.

**Type de complément :** Application de contenu et de volet Office

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>En tant que petit-enfant d’un élément VersionOverrides

Spécifie l’ensemble minimal de méthodes d’API JavaScript Office qui doivent être pris en charge par la version et la plateforme Office (telles que Windows, Mac, web et iOS ou iPad) pour que [versionOverrides](versionoverrides.md) prenne effet.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans ces schémas VersionOverrides**:

- Identique à [l’élément Requirements](requirements.md) parent.

**Associés à ces ensembles de conditions requises**:

- Identique à [l’élément Requirements](requirements.md) parent.

## <a name="syntax"></a>Syntaxe

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a>Contenu dans

[Configuration requise](requirements.md)

## <a name="can-contain"></a>Peut contenir

[Méthod](method.md)

## <a name="remarks"></a>Remarques

Les **éléments Methods** et **Method** ne sont pas pris en charge dans les modules de messagerie lorsqu’ils sont utilisés dans le manifeste de base. Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).
