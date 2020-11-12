---
title: Élément Tokens dans le fichier manifeste
description: Spécifie les jetons ou les caractères génériques qui peuvent être utilisés avec les modèles d’URL dans le manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: a50de7c2c3e8ebeb9425c1677a94bbcc62281d3b
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996681"
---
# <a name="tokens-element"></a>Élément Tokens

Définit les jetons qui peuvent être utilisés dans les URL de modèles.

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