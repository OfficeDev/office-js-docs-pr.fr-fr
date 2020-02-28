---
title: Élément Set dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d86b3123ff856e8618f92629308787b543e8228b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324805"
---
# <a name="set-element"></a>Élément Set

Spécifie un ensemble de conditions requises de l’API JavaScript Office que votre complément Office doit activer.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Contenu dans

[Ensembles](sets.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|Nom|string|obligatoire|Nom d’un [ensemble de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|chaîne|facultatif|Spécifie la version minimale de l’ensemble d’API requis par votre complément. Remplace la valeur de **DefaultMinVersion**, si elle est spécifiée dans l’élément parent [sets](sets.md) .|

## <a name="remarks"></a>Remarques

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **sets** , voir [Set the requirements ELEMENT dans le manifeste](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Pour les compléments de messagerie, il n'existe qu’un seul `"Mailbox"`ensemble de conditions requises disponible. Cet ensemble de conditions requises contient le sous-ensemble complet de l’API pris en charge dans les compléments de messagerie pour Outlook, et vous devez spécifier `"Mailbox"`l’ensemble de conditions requises dans le manifeste de votre complément de messagerie (ce n’est pas facultatif, comme c’est le cas pour les compléments de contenu et de volet Office).  De même, vous ne pouvez pas déclarer une prise en charge pour des méthodes spécifiques dans les compléments de messagerie.
