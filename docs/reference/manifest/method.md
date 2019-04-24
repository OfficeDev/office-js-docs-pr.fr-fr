---
title: Élément Method dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 19234b35e1faf8a8cc52a9e893fcc720793cadae
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450652"
---
# <a name="method-element"></a>Method, élément

Spécifie une méthode individuelle de l’API JavaScript pour Office requise pour l’activation de votre complément Office.

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
|Nom|string|obligatoire|Spécifie le nom de la méthode qualifiée requise avec son objet parent. Par exemple, pour spécifier la méthode **getSelectedDataAsync**, vous devez spécifier `"Document.getSelectedDataAsync"`.|

## <a name="remarks"></a>Remarques

Les éléments **Methods** et **Method** ne sont pas pris en charge par les compléments de messagerie. Pour plus d’informations sur les ensembles de spécifications, voir l’article [Versions Office et jeux de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT] 
> Étant donné qu’il n’existe aucun moyen de spécifier la version minimale requise pour les différentes méthodes, afin de vous assurer qu’une méthode est disponible lors de l’exécution, vous devez également utiliser une instruction **if** lorsque vous appelez cette méthode dans le script de votre complément. Pour plus d’informations sur la procédure à suivre, consultez l’article décrivant l’[API JavaScript pour Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

