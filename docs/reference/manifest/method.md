---
title: Élément Method dans le fichier manifeste
description: L’élément Method spécifie une méthode individuelle à partir de l’API JavaScript Office dont vos Office Add-ins ont besoin pour être activés par Office ou pour remplacer les paramètres de manifeste de base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 052fb41a7077781843ea7e63d9601a819058dfa6
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222268"
---
# <a name="method-element"></a>Élément Method

La signification de cet élément dépend de l’endroit où il est utilisé dans le manifeste.

## <a name="in-the-base-manifest"></a>Dans le manifeste de base

Lorsqu’il est utilisé dans le manifeste de base (autrement dit, si l’élément **Requirements** est un enfant direct [d’OfficeApp](officeapp.md)), l’élément **Method** spécifie une méthode individuelle à partir de l’API JavaScript Office dont votre application Office a besoin pour être activée par Office.

**Type de complément :** Application de contenu et de volet Office

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>En tant qu’arrière-petit-enfant d’un élément VersionOverrides

Spécifie une méthode individuelle à partir de l’API JavaScript Office qui doit être prise en charge par la version et la plateforme Office (telles que Windows, Mac, web et iOS ou iPad) pour que [versionOverrides](versionoverrides.md) prenne effet.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans ces schémas VersionOverrides**:

- Identique à [l’élément Requirements.](requirements.md)

**Associés à ces ensembles de conditions requises**:

- Identique à [l’élément Requirements.](requirements.md)

## <a name="syntax"></a>Syntaxe

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contenu dans

[Méthodes](methods.md)

## <a name="attributes"></a>Attributs

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|Nom|string|obligatoire|Spécifie le nom de la méthode qualifiée requise avec son objet parent. Par exemple, pour spécifier la `getSelectedDataAsync` méthode, vous devez spécifier `"Document.getSelectedDataAsync"` .|

## <a name="remarks"></a>Remarques

Les **éléments Methods** et **Method** ne sont pas pris en charge par les modules de messagerie lorsqu’ils sont utilisés dans le manifeste de base. Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

> [!IMPORTANT]
> Étant donné qu’il n’existe aucun moyen de spécifier la version minimale requise pour les différentes méthodes, afin de vous assurer qu’une méthode est disponible lors de l’exécution, vous devez également utiliser une instruction **if** lorsque vous appelez cette méthode dans le script de votre complément. Pour plus d’informations sur la façon de le faire, voir [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).
