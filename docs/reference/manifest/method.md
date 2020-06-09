---
title: Élément Method dans le fichier manifeste
description: L’élément Method spécifie une méthode individuelle de l’API JavaScript Office requise pour l’activation de vos compléments Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c3531475a920fd24ce8390170b5f4728d4dcd0e0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611756"
---
# <a name="method-element"></a>Élément Method

Spécifie une méthode individuelle de l’API JavaScript Office requise pour l’activation de votre complément Office.

**Type de complément :** Application de contenu et de volet Office

## <a name="syntax"></a>Syntaxe

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contenu dans

[Méthodes](methods.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|Nom|string|obligatoire|Spécifie le nom de la méthode qualifiée requise avec son objet parent. Par exemple, pour spécifier la `getSelectedDataAsync` méthode, vous devez spécifier `"Document.getSelectedDataAsync"` .|

## <a name="remarks"></a>Remarques

Les `Methods` `Method` éléments et ne sont pas pris en charge par les compléments de messagerie. Pour plus d’informations sur les ensembles de conditions requises, voir [versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

> [!IMPORTANT]
> Étant donné qu’il n’existe aucun moyen de spécifier la version minimale requise pour les différentes méthodes, afin de vous assurer qu’une méthode est disponible lors de l’exécution, vous devez également utiliser une instruction **if** lorsque vous appelez cette méthode dans le script de votre complément. Pour plus d’informations sur la façon de procéder, consultez [la rubrique Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).
