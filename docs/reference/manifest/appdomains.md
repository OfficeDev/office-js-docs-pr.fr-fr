---
title: Élément AppDomains dans le fichier manifeste
description: Répertorie tous les domaines en plus du domaine spécifié dans l' `SourceLocation` élément que votre complément Office utilisera et doit être approuvé par Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778654"
---
# <a name="appdomains-element"></a>AppDomains, élément

Répertorie tous les domaines, en plus du domaine spécifié dans l' `SourceLocation` élément, que votre complément Office utilisera et qui doit être approuvé par Office. Cela permet aux pages des domaines d’effectuer des appels à Office.js API depuis des IFrames dans le complément et présente d’autres effets. Pour chaque domaine supplémentaire, indiquez un élément **AppDomain**.

 **Type de complément :** Application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> Il existe des restrictions quant à ce qui peut être la valeur d’un élément **AppDomain** . Pour plus d’informations, consultez la rubrique [AppDomain](appdomain.md).

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

[AppDomain](appdomain.md)

## <a name="remarks"></a>Remarques

Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément[SourceLocation](sourcelocation.md). Vous devez indiquer une valeur pour cet élément.
