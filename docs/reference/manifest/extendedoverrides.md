---
title: Élément ExtendedOverrides dans le fichier manifeste
description: Spécifie les URL pour une extension au format JSON du manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 76491af34d1caf0ec266826df97a5363e336b85d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996687"
---
# <a name="extendedoverrides-element"></a>Élément ExtendedOverrides

Spécifie les URL complètes des fichiers au format JSON qui étendent le manifeste.

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
|URL (obligatoire)| URL complète du fichier JSON des substitutions étendues. Il peut s’agir d’un modèle d’URL qui utilise des jetons définis par l’élément [tokens](tokens.md) .|
|ResourcesUrl (facultatif) | URL complète d’un fichier qui fournit des ressources supplémentaires, telles que des chaînes localisées, pour le fichier spécifié dans l' `Url` attribut. Il peut s’agir d’un modèle d’URL qui utilise des jetons définis par l’élément [tokens](tokens.md) .|

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
