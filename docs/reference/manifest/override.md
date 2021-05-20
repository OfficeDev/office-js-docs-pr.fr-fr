---
title: Élément Override dans le fichier manifest
description: L’élément Override vous permet de spécifier la valeur d’un paramètre en fonction d’une condition spécifiée.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 131d72883d050038e2df5b7d8bbca033af9e6ee4
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555156"
---
# <a name="override-element"></a>Élément Override

Fournit un moyen de passer outre à la valeur d’un paramètre manifeste en fonction d’une condition spécifiée. Il existe trois types de conditions :

- Un Office local qui est différent de la valeur `LocaleToken` par défaut , appelé **LocaleTokenOverride**.
- Un modèle de support d’ensemble d’exigences différent du `RequirementToken` modèle par défaut, **appelé RequirementTokenOverride**.
- La source est différente de la valeur par `Runtime` défaut , **appelée RuntimeOverride** (actuellement en avant-première).

Un `<Override>` élément qui est à l’intérieur `<Runtime>` d’un élément doit être de type **RuntimeOverride**.

Il n’y a `overrideType` pas d’attribut pour `<Override>` l’élément. La différence est déterminée par l’élément parent et le type de l’élément parent. Un `<Override>` élément qui est à l’intérieur `<Token>` `xsi:type` d’un élément qui est , doit être de type `RequirementToken` **RequirementTokenOverride**. Un `<Override>` élément à l’intérieur de tout autre élément parent, ou à `<Override>` l’intérieur `LocaleToken` d’un élément de type , doit être de type **LocalTokenOverride**. Pour plus d’informations sur l’utilisation de cet élément lorsqu’il s’agit d’un enfant `<Token>` [d’un élément, voir Travail avec des dérogations étendues du manifeste](../../develop/extended-overrides.md).

Chaque type est décrit dans des sections distinctes plus tard dans cet article.

## <a name="override-element-for-localetoken"></a>Élément de remplacement pour `LocaleToken`

Un `<Override>` élément exprime un conditionnel et peut être lu comme un « Si ... puis ... » déclaration. Si `<Override>` l’élément est de type **LocalTokenOverride**, alors `Locale` l’attribut est la condition, et l’attribut `Value` est le conséquent. Par exemple, ce qui suit est lu " Si le paramètre Office local est fr-fr, alors le nom de l’affichage est 'Lecteur vidéo'. »

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

### <a name="examples"></a>Exemples

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

## <a name="override-element-for-requirementtoken"></a>Élément de remplacement pour `RequirementToken`

Un `<Override>` élément exprime un conditionnel et peut être lu comme un « Si ... puis ... » déclaration. Si `<Override>` l’élément est de type **RequirementTokenOverride**, alors l’élément `<Requirements>` enfant exprime la condition, et l’attribut `Value` est le conséquent. Par exemple, le premier élément suivant est lu « Si la plate-forme actuelle `<Override>` prend en charge la version FeatureOne 1.7, utilisez la chaîne « oldAddinVersion » à la place du `${token.requirements}` jeton dans l’URL du grand-parent (au lieu de la `<ExtendedOverrides>` chaîne par défaut « mise à niveau »). »

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
|Valeur|string|obligatoire|Valeur du jeton grand-parent lorsque la condition est remplie.|

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

## <a name="override-element-for-runtime-preview"></a>Élément de remplacement pour `Runtime` (aperçu)

> [!IMPORTANT]
> Cette fonctionnalité n’est prise en [charge que](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) pour un aperçu Outlook sur le web et sur Windows avec un abonnement Microsoft 365 spécial. Pour plus de détails, [consultez Configurez votre Outlook add-in pour l’activation basée sur l’événement](../../outlook/autolaunch.md).
>
> Étant donné que les fonctionnalités d’aperçu sont sujettes à changement sans préavis, elles ne doivent pas être utilisées dans les modules de production.

Un `<Override>` élément exprime un conditionnel et peut être lu comme un « Si ... puis ... » déclaration. Si `<Override>` l’élément est de type **RuntimeOverride**, alors `type` l’attribut est la condition, et l’attribut `resid` est le conséquent. Par exemple, ce qui suit est lu « Si le type est 'javascript', alors `resid` le est 'JSRuntime.Url'. » Outlook Desktop nécessite cet élément pour les [gestionnaires de points d’extension LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview)

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

**Type de complément :** messagerie

### <a name="syntax"></a>Syntaxe

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a>Contenu dans

- [Runtime](runtime.md)

### <a name="attributes"></a>Attributs

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|**type**|string|Oui|Spécifie la langue pour cette substitution. À l’heure `"javascript"` actuelle, est la seule option prise en charge.|
|**resid**|string|Oui|Spécifie l’emplacement de l’URL du fichier JavaScript qui doit passer outre à l’emplacement de l’URL du HTML par défaut défini dans [l’élément runtime](runtime.md) parent `resid` . Le `resid` ne peut pas être plus de 32 caractères et doit correspondre à un attribut `id` d’un `Url` élément dans `Resources` l’élément.|

### <a name="examples"></a>Exemples

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a>Voir aussi

- [Runtime](runtime.md)
- [Configurez votre Outlook add-in pour l’activation basée sur l’événement](../../outlook/autolaunch.md)
