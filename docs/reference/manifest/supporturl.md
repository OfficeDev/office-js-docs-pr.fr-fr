---
title: Élément SupportUrl dans le fichier manifest
description: L’élément SupportUrl spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre add-in.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 1d76afeaaceafc9e8786070338d69cea1b73635d20cd5a729d7e3d859b952494
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096352"
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
