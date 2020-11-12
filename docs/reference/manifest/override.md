---
title: Élément Override dans le fichier manifest
description: L’élément override vous permet de spécifier la valeur d’un paramètre en fonction d’une condition spécifiée.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 2c66503f9f95155a096b1b6fb23332eed8422da6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996311"
---
# <a name="override-element"></a>Élément Override

Permet de remplacer la valeur d’un paramètre de manifeste en fonction d’une condition spécifiée. Il existe deux types de conditions :

- Paramètres régionaux Office différents de ceux par défaut.
- Modèle de la prise en charge de l’ensemble de conditions requises qui est différente du modèle par défaut.

Il existe deux types d' `<Override>` éléments : un pour les substitutions de paramètres régionaux, appelé **LocaleTokenOverride** , et l’autre pour les substitutions d’ensemble de conditions requises, appelé **RequirementTokenOverride**. Mais il n’existe aucun `type` paramètre pour l' `<Override>` élément. La différence est déterminée par l’élément parent et le type de l’élément parent. Un `<Override>` élément qui se trouve à l’intérieur d’un `<Token>` élément dont le `xsi:type` est `RequirementToken` , doit être de type **RequirementTokenOverride**. Un `<Override>` élément situé à l’intérieur d’un autre élément parent, ou à l’intérieur d’un `<Override>` élément de type `LocaleToken` , doit être de type **LocaleTokenOverride**. Chaque type est décrit dans des sections distinctes ci-dessous.

## <a name="override-element-of-type-localetokenoverride"></a>Élément override de type LocaleTokenOverride

Un `<Override>` élément exprime un conditionnel et peut être lu sous la forme d’un «if... Then... " résultat. Si l' `<Override>` élément est de type **LocaleTokenOverride** , l' `Locale` attribut est la condition, et l' `Value` attribut est le à la suite. Par exemple, le code suivant est lu « si le paramètre paramètres régionaux Office est fr-fr, le nom complet est «lecteur vidéo ».»

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

**Type de complément :** application de contenu, de volet Office, de messagerie

### <a name="syntax"></a>Syntaxe

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a>Contenu dans

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
|[Jeton](token.md)|

### <a name="attributes"></a>Attributs

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|Paramètres régionaux|string|obligatoire|Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.|
|Valeur|string|obligatoire|Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.|

### <a name="examples"></a>範例

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
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
```

### <a name="see-also"></a>Voir aussi

- [Localisation des compléments Office](../../develop/localization.md)
- [Raccourcis clavier](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a>Élément override de type RequirementTokenOverride

Un `<Override>` élément exprime un conditionnel et peut être lu sous la forme d’un «if... Then... " résultat. Si l' `<Override>` élément est de type **RequirementTokenOverride** , l’élément enfant `<Requirements>` exprime la condition, et l' `Value` attribut est le à la suite. Par exemple, le premier `<Override>` des éléments suivants est lu « si la plateforme actuelle prend en charge la version 1,7 de FeatureOne, puis utilisez la chaîne «oldAddinVersion » à la place du `${token.requirements}` jeton dans l’URL du grand-parent `<ExtendedOverrides>` (au lieu de la chaîne par défaut « mise à niveau »)».

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

**Type de complément :** volet Office

### <a name="syntax"></a>Syntaxe

```XML
<Override Value="string" />
```

### <a name="contained-in"></a>Contenu dans

|Élément|
|:-----|
|[Jeton](token.md)|

### <a name="must-contain"></a>Doit contenir

|Élément|Contenu|Courrier|TaskPane|
|:-----|:-----|:-----|:-----|
|[Configuration requise](requirements.md)|||x|

### <a name="attributes"></a>Attributs

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|Valeur|string|obligatoire|Valeur du jeton de grand-parent lorsque la condition est satisfaite.|

### <a name="example"></a>Exemple

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Définition de l’élément Requirements dans le manifeste](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [Raccourcis clavier](../../design/keyboard-shortcuts.md)
