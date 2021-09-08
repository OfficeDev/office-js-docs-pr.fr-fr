---
title: Élément Set dans le fichier manifeste
description: L’élément Set spécifie un ensemble Office conditions requises de l’API JavaScript requises Office votre Office pour activer.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 608830e1ebc0d2e2d4c170b48bba00b3a19e87af
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937119"
---
# <a name="set-element"></a>Élément Set

Spécifie un ensemble de conditions requises à partir Office’API JavaScript requise par votre Office pour l’activer.

**Type de complément :** application de contenu, de volet Office, de messagerie

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

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et l’attribut **DefaultMinVersion** de l’élément **Sets,** voir Définir l’élément [Requirements dans le manifeste.](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)

> [!IMPORTANT]
> Pour les compléments de messagerie, il n'existe qu’un seul `"Mailbox"`ensemble de conditions requises disponible. Cet ensemble de conditions requises contient le sous-ensemble complet de l’API pris en charge dans les compléments de messagerie pour Outlook, et vous devez spécifier `"Mailbox"`l’ensemble de conditions requises dans le manifeste de votre complément de messagerie (ce n’est pas facultatif, comme c’est le cas pour les compléments de contenu et de volet Office).  De même, vous ne pouvez pas déclarer une prise en charge pour des méthodes spécifiques dans les compléments de messagerie.
