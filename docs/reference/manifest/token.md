---
title: Élément Token dans le fichier manifeste
description: Spécifie un jeton ou un caractère générique qui peut être utilisé avec des modèles d’URL dans le manifeste.
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: 69f626f5f6f57dd155756812bcd56267a1da3ffa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150299"
---
# <a name="token-element"></a>Élément Token

Définit un jeton d’URL individuel. Pour plus d’informations sur l’utilisation de cet élément, voir [Work with extended overrides of the manifest](../../develop/extended-overrides.md).

**Type de complément :** volet Office

## <a name="syntax"></a>Syntaxe

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>Contenu dans

[Jetons](tokens.md)

## <a name="can-contain"></a>Peut contenir

|Élément|Contenu|Courrier|TaskPane|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>Attributs

|Attribut|Description|
|:-----|:-----|
|DefaultValue|Valeur par défaut de ce jeton si aucune condition dans un élément `<Override>` enfant ne correspond.|
|Nom|Nom du jeton. Ce nom est défini par l’utilisateur. Le type du jeton est déterminé par l’attribut de type.|
|xsi:type|Définit le type de jeton. Cet attribut doit être définie sur l’une des valeurs  `"RequirementsToken"` : ou  `"LocaleToken"` .|

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