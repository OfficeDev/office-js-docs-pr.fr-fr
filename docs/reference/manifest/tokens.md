---
title: Élément Tokens dans le fichier manifeste
description: Spécifie les jetons ou les caractères génériques qui peuvent être utilisés avec des modèles d’URL dans le manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 8680b985068c44e93f601a2b24e2f28899eb483d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505324"
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

|Élément|Contenu|Courrier|TaskPane|
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