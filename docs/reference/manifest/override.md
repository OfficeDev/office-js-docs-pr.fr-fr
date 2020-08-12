---
title: Élément Override dans le fichier manifest
description: L’élément override vous permet de spécifier la valeur d’un paramètre pour des paramètres régionaux supplémentaires.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 139a4089a36d8a8adfa71d4a0947b02f5b163b52
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641451"
---
# <a name="override-element"></a>Élément Override

Fournit une manière de spécifier la valeur d’un paramètre pour d’autres paramètres régionaux.

**Type de complément:** application de contenu, de volet Office, de messagerie (Mail)

## <a name="syntax"></a>Syntaxe

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>Contenu dans

|Élément|
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

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|Paramètres régionaux|string|obligatoire|Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.|
|Valeur|string|obligatoire|Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.|

## <a name="see-also"></a>Voir aussi

- [Localisation des compléments Office](../../develop/localization.md)
