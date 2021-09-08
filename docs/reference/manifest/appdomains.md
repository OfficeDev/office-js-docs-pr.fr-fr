---
title: Élément AppDomains dans le fichier manifeste
description: Répertorie tous les domaines en plus du domaine spécifié dans l’élément que votre complément Office utilisera et doit être approuvé par `SourceLocation` Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937910"
---
# <a name="appdomains-element"></a>AppDomains, élément

Répertorie tous les domaines, en plus du domaine spécifié dans l’élément, que votre complément Office utilisera et qui doivent être `SourceLocation` Office. Cela permet aux pages des domaines d’effectuer des appels à Office.js API à partir d’IFrames au sein du module et a d’autres effets. Pour chaque domaine supplémentaire, indiquez un élément **AppDomain**.

 **Type de complément :** Application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> Il existe des restrictions sur ce qui peut être la valeur d’un **élément AppDomain.** Pour plus d’informations, [voir AppDomain](appdomain.md).

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

[AppDomain](appdomain.md)

## <a name="remarks"></a>Remarques

Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément[SourceLocation](sourcelocation.md). Vous devez indiquer une valeur pour cet élément.
