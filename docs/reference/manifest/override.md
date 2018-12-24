---
title: Élément Override dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1d2400312f12116b1ac5f4010135541e783dcc7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432864"
---
# <a name="override-element"></a>Élément Override

Fournit une manière de spécifier la valeur d’un paramètre pour d’autres paramètres régionaux.

**Type de complément:** application de contenu, de volet Office, de messagerie (Mail)

## <a name="syntax"></a>Syntaxe

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>Contenu dans

|**Élément**|
|:-----|
|[CitationText](citationtext.md)|
|[Description](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|Locale|string|obligatoire|Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.|
|Valeur|string|obligatoire|Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.|

## <a name="see-also"></a>Voir aussi

- [Localisation des compléments Office](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
