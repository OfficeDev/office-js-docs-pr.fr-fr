---
title: Élément Method dans le fichier manifeste
description: L’élément Method spécifie une méthode individuelle à partir de l’API JavaScript Office dont vos Office de développement ont besoin pour s’activer.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 037446f5027a97214d2b1be6ee99c8f6822b33b9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150324"
---
# <a name="method-element"></a>Élément Method

Spécifie une méthode individuelle à partir de l’API JavaScript Office que votre Office nécessite pour s’activer.

**Type de complément :** Application de contenu et de volet Office

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

Les éléments et les éléments ne sont pas pris en charge par `Methods` `Method` les modules de messagerie. Pour plus d’informations sur les ensembles de conditions requises, [voir Office versions et les ensembles de conditions requises.](../../develop/office-versions-and-requirement-sets.md)

> [!IMPORTANT]
> Étant donné qu’il n’existe aucun moyen de spécifier la version minimale requise pour les différentes méthodes, afin de vous assurer qu’une méthode est disponible lors de l’exécution, vous devez également utiliser une instruction **if** lorsque vous appelez cette méthode dans le script de votre complément. Pour plus d’informations sur la façon de le faire, voir [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).
