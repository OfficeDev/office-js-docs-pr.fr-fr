---
title: Élément SupportUrl dans le fichier manifest
description: L’élément SupportUrl spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e38030062c48936f925126e896cd74e660164a5d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720343"
---
# <a name="supporturl-element"></a>SupportUrl, élément

Spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.

## <a name="syntax"></a>Syntaxe

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

|  Élément | Requis | Description  |
|:-----|:-----|:-----|
|  [Override](override.md)   | Non | Spécifie le paramètre pour les URL de paramètres régionaux supplémentaires |

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|
