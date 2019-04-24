---
title: Élément AppDomains dans le fichier manifeste
description: ''
ms.date: 12/13/2018
localization_priority: Normal
ms.openlocfilehash: 65391c9529e7ddaa9726d0b58accf90c5b9babef
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450645"
---
# <a name="appdomains-element"></a>AppDomains, élément

Répertorie tout domaine supplémentaire qui sera utilisé par votre complément Office pour charger des pages en plus du domaine spécifié dans l’élément SourceLocation. Pour chaque domaine supplémentaire, indiquez un élément AppDomain.

 **Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> La valeur de chaque élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain<AppDomain>`).

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

[AppDomain](appdomain.md)

## <a name="remarks"></a>Remarques

Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément[SourceLocation](sourcelocation.md). Pour charger des pages qui ne sont pas dans le même domaine que le complément, spécifiez les domaines à l’aide des éléments **AppDomains** et **AppDomain**. Vous devez indiquer une valeur pour cet élément.
