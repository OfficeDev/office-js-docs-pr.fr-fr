---
title: Élément Override dans le fichier manifest
description: L’élément override vous permet de spécifier la valeur d’un paramètre pour des paramètres régionaux supplémentaires.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: aa5d023169389670d15e36f8bee4445529d84711
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611504"
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
|Paramètres régionaux|string|obligatoire|Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.|
|Valeur|string|obligatoire|Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.|

## <a name="see-also"></a>Voir aussi

- [Localisation des compléments Office](../../develop/localization.md)
