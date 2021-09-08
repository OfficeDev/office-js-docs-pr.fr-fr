---
title: Élément SupportUrl dans le fichier manifest
description: L’élément SupportUrl spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre add-in.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: be516fe5848d775dacb0d424a92be02d59f85512
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937013"
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

|  Élément | Obligatoire | Description  |
|:-----|:-----|:-----|
|  [Override](override.md)   | Non | Spécifie le paramètre pour les URL de paramètres régionaux supplémentaires |

## <a name="attributes"></a>Attributs

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|
