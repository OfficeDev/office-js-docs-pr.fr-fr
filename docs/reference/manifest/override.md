---
title: Élément Override dans le fichier manifest
description: L’élément Override vous permet de spécifier la valeur d’un paramètre en fonction d’une condition spécifiée.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: e4e2ccd9936eec12fd7adb4eca8e46a5f391785f
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222261"
---
# <a name="override-element"></a>Élément Override

Permet de remplacer la valeur d’un paramètre de manifeste en fonction d’une condition spécifiée. Il existe trois types de conditions :

- Un Office qui est différent du paramètre par défaut `LocaleToken` , **appelé LocaleTokenOverride**.
- Modèle de prise en charge de l’ensemble de conditions requises différent du modèle par `RequirementToken` défaut, appelé **RequirementTokenOverride**.
- La source est différente de la valeur par `Runtime` défaut, **appelée RuntimeOverride**.

Un **élément Override** qui se trouve à l’intérieur d’un **élément Runtime** doit être de type **RuntimeOverride**.

Il n’existe `overrideType` aucun attribut pour **l’élément Override.** La différence est déterminée par l’élément parent et le type de l’élément parent. Un **élément Override** qui se trouve à l’intérieur d’un élément **Token** dont , doit être de `xsi:type` type `RequirementToken` **RequirementTokenOverride**. Un **élément Override à** l’intérieur d’un autre élément parent, ou à l’intérieur d’un élément **Override** de type , doit être de type `LocaleToken` **LocaleTokenOverride**. Pour plus d’informations sur l’utilisation de cet élément lorsqu’il s’agit d’un enfant d’un élément **Token,** voir Work [with extended overrides of the manifest](../../develop/extended-overrides.md).

Chaque type est décrit dans des sections distinctes plus loin dans cet article.

## <a name="override-element-for-localetoken"></a>Élément Override pour `LocaleToken`

Un **élément Override** exprime une conditionnel et peut être lu en tant que « If ... then ... » . Si **l’élément Override** est de type **LocaleTokenOverride**, l’attribut est la condition et l’attribut `Locale` en est la `Value` conséquence. Par exemple, l’exemple suivant indique « Si le paramètre Office paramètres régionaux est fr-fr, le nom complet est Lecteur vidéo ».

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

## <a name="override-element-for-requirementtoken"></a>Élément Override pour `RequirementToken`

Un **élément Override** exprime une conditionnel et peut être lu en tant que « If ... then ... » . Si **l’élément Override** est de type **RequirementTokenOverride**, l’élément **Requirements** enfant exprime la condition et l’attribut `Value` en est le résultat. Par exemple,  la première substitution de l’exemple suivant est la suivante : « Si la plateforme actuelle prend en charge FeatureOne version 1.7, utilisez la chaîne « oldAddinVersion » à la place du jeton dans l’URL de l’illustre (au lieu de la chaîne par défaut `${token.requirements}` `<ExtendedOverrides>` « upgrade »).

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
|Valeur|string|obligatoire|Valeur du jeton de preuve lorsque la condition est remplie.|

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
- [Spécifier les Office et les plateformes qui peuvent héberger votre add-in](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)
- [Raccourcis clavier](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime"></a>Élément Override pour `Runtime`

> [!IMPORTANT]
> La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises [mailbox 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) avec la fonctionnalité d’activation basée [sur les événements.](../../outlook/autolaunch.md) Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

Un **élément Override** exprime une conditionnel et peut être lu en tant que « If ... then ... » . Si **l’élément Override** est de type **RuntimeOverride,** l’attribut est la condition et l’attribut `type` en est la `resid` conséquence. Par exemple, l’exemple suivant est « Si le type est « javascript », il `resid` s’agit de « JSRuntime.Url ». Outlook Desktop requiert cet élément pour les handleurs de [point d’extension LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent)

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
|**type**|string|Oui|Spécifie la langue de ce remplacement. Pour l’instant, `"javascript"` il s’agit de la seule option prise en charge.|
|**resid**|string|Oui|Spécifie l’emplacement d’URL du fichier JavaScript qui doit remplacer l’emplacement d’URL du code HTML par défaut défini dans l’élément [Runtime](runtime.md) `resid` parent. Il ne peut pas y avoir plus de 32 caractères et doit correspondre à un `resid` `id` attribut `Url` d’un élément dans `Resources` l’élément.|

### <a name="examples"></a>Exemples

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a>Voir aussi

- [Runtime](runtime.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](../../outlook/autolaunch.md)
