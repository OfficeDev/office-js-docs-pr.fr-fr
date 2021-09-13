---
title: Élément ExtendedOverrides dans le fichier manifeste
description: Spécifie les URL d’une extension au format JSON du manifeste.
ms.date: 02/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 062d62284d786e81f230b1d094c46a73208a6c2c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153033"
---
# <a name="extendedoverrides-element"></a>Élément ExtendedOverrides

Spécifie les URL complètes pour les fichiers au format JSON qui étendent le manifeste. Pour plus d’informations sur l’utilisation de cet élément et de ses éléments descendants, voir [Work with extended overrides of the manifest](../../develop/extended-overrides.md).

**Type de complément :** volet Office

## <a name="syntax"></a>Syntaxe

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

|Élément|Contenu|Courrier|TaskPane|
|:-----|:-----|:-----|:-----|
|[Jetons](tokens.md)|||x|

## <a name="attributes"></a>Attributs

|Attribut|Description|
|:-----|:-----|
|Url (obligatoire)| URL complète du fichier JSON de remplacements étendu. À l’avenir, cette valeur pourrait être un modèle d’URL qui utilise des jetons définis par [l’élément Tokens.](tokens.md) Voir [exemples](#examples).|
|ResourcesUrl (facultatif) | URL complète d’un fichier qui fournit des ressources supplémentaires, telles que des chaînes localisées, pour le fichier spécifié dans `Url` l’attribut. Il peut s’agit d’un modèle d’URL qui utilise des jetons définis par [l’élément Tokens.](tokens.md)|

## <a name="examples"></a>Exemples

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

À l’avenir, cette valeur pourrait être un modèle d’URL qui utilise des jetons définis par [l’élément Tokens.](tokens.md) Voici un exemple.

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
