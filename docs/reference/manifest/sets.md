---
title: Élément Sets dans le fichier manifeste
description: L’élément sets spécifie l’ensemble minimal d’API JavaScript pour Office requis pour l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bd8f8311bb06a8e9e98fc408aece6395ab5643b1
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641423"
---
# <a name="sets-element"></a>Élément Sets

Spécifie le sous-ensemble minimal de l’API JavaScript Office requise pour l’activation de votre complément Office.

**Type de complément :** application de contenu, de volet Office, de messagerie

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
|DefaultMinVersion|chaîne|facultatif|Spécifie la valeur par défaut de l’attribut **MinVersion** pour tous les éléments [Set](set.md) enfants. La valeur par défaut est « 1.1 ».|

## <a name="remarks"></a>Remarques

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **sets** , voir [Set the requirements ELEMENT dans le manifeste](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

