---
title: Élément Sets dans le fichier manifeste
description: L’élément Sets spécifie l’ensemble minimal de Office’API JavaScript dont votre Office a besoin pour s’activer.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a0a7edf6543cc74ac69ee6dc430c0a7497b6911ed43d66ea1082c0d477255948
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095018"
---
# <a name="sets-element"></a>Élément Sets

Spécifie le sous-ensemble minimal de l’API JavaScript Office que votre Office nécessite pour être activé.

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
|DefaultMinVersion|chaîne|facultatif|Spécifie la valeur **d’attribut MinVersion** par défaut pour tous les éléments [Set](set.md) enfants. La valeur par défaut est « 1.1 ».|

## <a name="remarks"></a>Remarques

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et l’attribut **DefaultMinVersion** de l’élément **Sets,** voir Définir l’élément [Requirements dans le manifeste.](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)

