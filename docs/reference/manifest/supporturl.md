---
title: Élément SupportUrl dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 18b9b7c4df9def70ab42ae213066188ac04c07a7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450414"
---
# <a name="supporturl-element"></a>Élément SupportUrl

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

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|
