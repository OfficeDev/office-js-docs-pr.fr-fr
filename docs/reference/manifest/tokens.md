---
title: Élément Tokens dans le fichier manifeste
description: Spécifie les jetons ou les caractères génériques qui peuvent être utilisés avec des modèles d’URL dans le manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5d42abab46ecc6e7ab465144f061d26da52c0eb3e2623acd8a8a2912ecc13312
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095785"
---
# <a name="tokens-element"></a>Élément Tokens

Définit les jetons qui peuvent être utilisés dans les URL de modèle. Pour plus d’informations sur l’utilisation de cet élément, voir [Work with extended overrides of the manifest](../../develop/extended-overrides.md).

**Type de complément :** volet Office

## <a name="syntax"></a>Syntaxe

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a>Contenu dans

[ExtendedOverrides](extendedoverrides.md)

## <a name="must-contain"></a>Doit contenir

|Élément|Contenu|Courrier Outlook|TaskPane|
|:-----|:-----|:-----|:-----|
|[Jeton](token.md)|||x|

## <a name="example"></a>Exemple

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```