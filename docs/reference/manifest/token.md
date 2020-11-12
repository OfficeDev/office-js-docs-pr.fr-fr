---
title: Élément de jeton dans le fichier manifeste
description: Spécifie un jeton ou un caractère générique qui peut être utilisé avec les modèles d’URL dans le manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5e26af44c566ab09ac81c8194e1ae7d85aaac327
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996685"
---
# <a name="token-element"></a>Élément Token

Définit un jeton d’URL individuel.

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
|DefaultValue|Valeur par défaut de ce jeton si aucune condition n’est `<Override>` correspondante dans un élément enfant.|
|Nom|Nom du jeton. Ce nom est défini par l’utilisateur. Le type du jeton est déterminé par l’attribut type.|
|xsi:type|Définit le type de jeton. Cet attribut doit être défini sur l’un des éléments suivants :  `"RequirementsToken"` , ou  `"LocaleToken"` .|

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