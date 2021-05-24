---
title: Élément Override dans le fichier manifest
description: L’élément Override vous permet de spécifier la valeur d’un paramètre en fonction d’une condition spécifiée.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd270fa19750810238b42c26c2abc35a61c1bac8
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590903"
---
# <a name="override-element"></a><span data-ttu-id="5b232-103">Élément Override</span><span class="sxs-lookup"><span data-stu-id="5b232-103">Override element</span></span>

<span data-ttu-id="5b232-104">Permet de remplacer la valeur d’un paramètre de manifeste en fonction d’une condition spécifiée.</span><span class="sxs-lookup"><span data-stu-id="5b232-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="5b232-105">Il existe trois types de conditions :</span><span class="sxs-lookup"><span data-stu-id="5b232-105">There are three kinds of conditions:</span></span>

- <span data-ttu-id="5b232-106">Un Office qui est différent du paramètre par `LocaleToken` défaut, **appelé LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="5b232-106">An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.</span></span>
- <span data-ttu-id="5b232-107">Modèle de prise en charge de l’ensemble de conditions requises différent du modèle par `RequirementToken` défaut, appelé **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="5b232-107">A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.</span></span>
- <span data-ttu-id="5b232-108">La source est différente de la valeur par `Runtime` défaut, **appelée RuntimeOverride**.</span><span class="sxs-lookup"><span data-stu-id="5b232-108">The source is different from the default `Runtime`, called **RuntimeOverride**.</span></span>

<span data-ttu-id="5b232-109">Un `<Override>` élément qui se trouve à l’intérieur d’un élément doit être de type `<Runtime>` **RuntimeOverride**.</span><span class="sxs-lookup"><span data-stu-id="5b232-109">An `<Override>` element that is inside of a `<Runtime>` element must be of type **RuntimeOverride**.</span></span>

<span data-ttu-id="5b232-110">Il n’existe `overrideType` aucun attribut pour `<Override>` l’élément.</span><span class="sxs-lookup"><span data-stu-id="5b232-110">There is no `overrideType` attribute for the `<Override>` element.</span></span> <span data-ttu-id="5b232-111">La différence est déterminée par l’élément parent et le type de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="5b232-111">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="5b232-112">Un `<Override>` élément qui se trouve à l’intérieur d’un élément dont , doit être de type `<Token>` `xsi:type` `RequirementToken` **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="5b232-112">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="5b232-113">Un élément à l’intérieur d’un autre élément parent, ou à l’intérieur d’un élément de type, doit `<Override>` `<Override>` être de type `LocaleToken` **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="5b232-113">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="5b232-114">Pour plus d’informations sur l’utilisation de cet élément lorsqu’il est enfant d’un élément, voir `<Token>` [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="5b232-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="5b232-115">Chaque type est décrit dans des sections distinctes plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="5b232-115">Each type is described in separate sections later in this article.</span></span>

## <a name="override-element-for-localetoken"></a><span data-ttu-id="5b232-116">Élément Override pour `LocaleToken`</span><span class="sxs-lookup"><span data-stu-id="5b232-116">Override element for `LocaleToken`</span></span>

<span data-ttu-id="5b232-117">Un `<Override>` élément exprime une conditionnel et peut être lu en tant que « If ... alors... .</span><span class="sxs-lookup"><span data-stu-id="5b232-117">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="5b232-118">Si `<Override>` l’élément est de type **LocaleTokenOverride**, l’attribut est la condition et l’attribut `Locale` en est le `Value` résultat.</span><span class="sxs-lookup"><span data-stu-id="5b232-118">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="5b232-119">Par exemple, le texte suivant est lu « Si le paramètre Office paramètres régionaux est fr-fr, le nom complet est Lecteur vidéo ».</span><span class="sxs-lookup"><span data-stu-id="5b232-119">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="5b232-120">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="5b232-120">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="5b232-121">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="5b232-121">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="5b232-122">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="5b232-122">Contained in</span></span>

|<span data-ttu-id="5b232-123">Élément</span><span class="sxs-lookup"><span data-stu-id="5b232-123">Element</span></span>|
|:-----|
|[<span data-ttu-id="5b232-124">CitationText</span><span class="sxs-lookup"><span data-stu-id="5b232-124">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="5b232-125">Description</span><span class="sxs-lookup"><span data-stu-id="5b232-125">Description</span></span>](description.md)|
|[<span data-ttu-id="5b232-126">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="5b232-126">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="5b232-127">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="5b232-127">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="5b232-128">DisplayName</span><span class="sxs-lookup"><span data-stu-id="5b232-128">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="5b232-129">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="5b232-129">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="5b232-130">IconUrl</span><span class="sxs-lookup"><span data-stu-id="5b232-130">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="5b232-131">QueryUri</span><span class="sxs-lookup"><span data-stu-id="5b232-131">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="5b232-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5b232-132">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="5b232-133">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="5b232-133">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="5b232-134">Jeton</span><span class="sxs-lookup"><span data-stu-id="5b232-134">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="5b232-135">Attributs</span><span class="sxs-lookup"><span data-stu-id="5b232-135">Attributes</span></span>

|<span data-ttu-id="5b232-136">Attribut</span><span class="sxs-lookup"><span data-stu-id="5b232-136">Attribute</span></span>|<span data-ttu-id="5b232-137">Type</span><span class="sxs-lookup"><span data-stu-id="5b232-137">Type</span></span>|<span data-ttu-id="5b232-138">Requis</span><span class="sxs-lookup"><span data-stu-id="5b232-138">Required</span></span>|<span data-ttu-id="5b232-139">Description</span><span class="sxs-lookup"><span data-stu-id="5b232-139">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5b232-140">Paramètres régionaux</span><span class="sxs-lookup"><span data-stu-id="5b232-140">Locale</span></span>|<span data-ttu-id="5b232-141">string</span><span class="sxs-lookup"><span data-stu-id="5b232-141">string</span></span>|<span data-ttu-id="5b232-142">obligatoire</span><span class="sxs-lookup"><span data-stu-id="5b232-142">required</span></span>|<span data-ttu-id="5b232-143">Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="5b232-143">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="5b232-144">Valeur</span><span class="sxs-lookup"><span data-stu-id="5b232-144">Value</span></span>|<span data-ttu-id="5b232-145">string</span><span class="sxs-lookup"><span data-stu-id="5b232-145">string</span></span>|<span data-ttu-id="5b232-146">obligatoire</span><span class="sxs-lookup"><span data-stu-id="5b232-146">required</span></span>|<span data-ttu-id="5b232-147">Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.</span><span class="sxs-lookup"><span data-stu-id="5b232-147">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="5b232-148">Exemples</span><span class="sxs-lookup"><span data-stu-id="5b232-148">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="5b232-149">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5b232-149">See also</span></span>

- [<span data-ttu-id="5b232-150">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="5b232-150">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="5b232-151">Raccourcis clavier</span><span class="sxs-lookup"><span data-stu-id="5b232-151">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a><span data-ttu-id="5b232-152">Élément Override pour `RequirementToken`</span><span class="sxs-lookup"><span data-stu-id="5b232-152">Override element for `RequirementToken`</span></span>

<span data-ttu-id="5b232-153">Un `<Override>` élément exprime une conditionnel et peut être lu en tant que « If ... alors... .</span><span class="sxs-lookup"><span data-stu-id="5b232-153">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="5b232-154">Si `<Override>` l’élément est de type **RequirementTokenOverride**, l’élément enfant exprime la condition et l’attribut `<Requirements>` en est le `Value` résultat.</span><span class="sxs-lookup"><span data-stu-id="5b232-154">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="5b232-155">Par exemple, la première partie de ce qui suit est lue « Si la plateforme actuelle prend en charge FeatureOne version 1.7, utilisez la chaîne « oldAddinVersion » à la place du jeton dans l’URL de l’outre-famille (au lieu de la chaîne par défaut `<Override>` `${token.requirements}` « upgrade `<ExtendedOverrides>` »).</span><span class="sxs-lookup"><span data-stu-id="5b232-155">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="5b232-156">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="5b232-156">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="5b232-157">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="5b232-157">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="5b232-158">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="5b232-158">Contained in</span></span>

|<span data-ttu-id="5b232-159">Élément</span><span class="sxs-lookup"><span data-stu-id="5b232-159">Element</span></span>|
|:-----|
|[<span data-ttu-id="5b232-160">Jeton</span><span class="sxs-lookup"><span data-stu-id="5b232-160">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="5b232-161">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="5b232-161">Must contain</span></span>

|<span data-ttu-id="5b232-162">Élément</span><span class="sxs-lookup"><span data-stu-id="5b232-162">Element</span></span>|<span data-ttu-id="5b232-163">Contenu</span><span class="sxs-lookup"><span data-stu-id="5b232-163">Content</span></span>|<span data-ttu-id="5b232-164">Courrier</span><span class="sxs-lookup"><span data-stu-id="5b232-164">Mail</span></span>|<span data-ttu-id="5b232-165">TaskPane</span><span class="sxs-lookup"><span data-stu-id="5b232-165">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="5b232-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5b232-166">Requirements</span></span>](requirements.md)|||<span data-ttu-id="5b232-167">x</span><span class="sxs-lookup"><span data-stu-id="5b232-167">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="5b232-168">Attributs</span><span class="sxs-lookup"><span data-stu-id="5b232-168">Attributes</span></span>

|<span data-ttu-id="5b232-169">Attribut</span><span class="sxs-lookup"><span data-stu-id="5b232-169">Attribute</span></span>|<span data-ttu-id="5b232-170">Type</span><span class="sxs-lookup"><span data-stu-id="5b232-170">Type</span></span>|<span data-ttu-id="5b232-171">Requis</span><span class="sxs-lookup"><span data-stu-id="5b232-171">Required</span></span>|<span data-ttu-id="5b232-172">Description</span><span class="sxs-lookup"><span data-stu-id="5b232-172">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5b232-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="5b232-173">Value</span></span>|<span data-ttu-id="5b232-174">string</span><span class="sxs-lookup"><span data-stu-id="5b232-174">string</span></span>|<span data-ttu-id="5b232-175">obligatoire</span><span class="sxs-lookup"><span data-stu-id="5b232-175">required</span></span>|<span data-ttu-id="5b232-176">Valeur du jeton de preuve lorsque la condition est remplie.</span><span class="sxs-lookup"><span data-stu-id="5b232-176">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="5b232-177">Exemple</span><span class="sxs-lookup"><span data-stu-id="5b232-177">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="5b232-178">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5b232-178">See also</span></span>

- [<span data-ttu-id="5b232-179">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="5b232-179">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="5b232-180">Définition de l’élément Requirements dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="5b232-180">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="5b232-181">Raccourcis clavier</span><span class="sxs-lookup"><span data-stu-id="5b232-181">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime"></a><span data-ttu-id="5b232-182">Élément Override pour `Runtime`</span><span class="sxs-lookup"><span data-stu-id="5b232-182">Override element for `Runtime`</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5b232-183">La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises [mailbox 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) avec la fonctionnalité [d’activation basée sur les événements.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="5b232-183">Support for this element was introduced in [Mailbox requirement set 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) with the [event-based activation feature](../../outlook/autolaunch.md).</span></span> <span data-ttu-id="5b232-184">Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="5b232-184">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="5b232-185">Un `<Override>` élément exprime une conditionnel et peut être lu en tant que « If ... alors... .</span><span class="sxs-lookup"><span data-stu-id="5b232-185">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="5b232-186">Si `<Override>` l’élément est de type **RuntimeOverride**, l’attribut est la `type` condition et `resid` l’attribut en est la conséquence.</span><span class="sxs-lookup"><span data-stu-id="5b232-186">If the `<Override>` element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent.</span></span> <span data-ttu-id="5b232-187">Par exemple, l’exemple suivant est « Si le type est « javascript », il `resid` s’agit de « JSRuntime.Url ». Outlook Desktop requiert cet élément pour les handleurs de [point d’extension LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent)</span><span class="sxs-lookup"><span data-stu-id="5b232-187">For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span>

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

<span data-ttu-id="5b232-188">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="5b232-188">**Add-in type:** Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="5b232-189">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="5b232-189">Syntax</span></span>

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a><span data-ttu-id="5b232-190">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="5b232-190">Contained in</span></span>

- [<span data-ttu-id="5b232-191">Runtime</span><span class="sxs-lookup"><span data-stu-id="5b232-191">Runtime</span></span>](runtime.md)

### <a name="attributes"></a><span data-ttu-id="5b232-192">Attributs</span><span class="sxs-lookup"><span data-stu-id="5b232-192">Attributes</span></span>

|<span data-ttu-id="5b232-193">Attribut</span><span class="sxs-lookup"><span data-stu-id="5b232-193">Attribute</span></span>|<span data-ttu-id="5b232-194">Type</span><span class="sxs-lookup"><span data-stu-id="5b232-194">Type</span></span>|<span data-ttu-id="5b232-195">Requis</span><span class="sxs-lookup"><span data-stu-id="5b232-195">Required</span></span>|<span data-ttu-id="5b232-196">Description</span><span class="sxs-lookup"><span data-stu-id="5b232-196">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5b232-197">**type**</span><span class="sxs-lookup"><span data-stu-id="5b232-197">**type**</span></span>|<span data-ttu-id="5b232-198">string</span><span class="sxs-lookup"><span data-stu-id="5b232-198">string</span></span>|<span data-ttu-id="5b232-199">Oui</span><span class="sxs-lookup"><span data-stu-id="5b232-199">Yes</span></span>|<span data-ttu-id="5b232-200">Spécifie la langue de ce remplacement.</span><span class="sxs-lookup"><span data-stu-id="5b232-200">Specifies the language for this override.</span></span> <span data-ttu-id="5b232-201">Pour l’instant, `"javascript"` c’est la seule option prise en charge.</span><span class="sxs-lookup"><span data-stu-id="5b232-201">At present, `"javascript"` is the only supported option.</span></span>|
|<span data-ttu-id="5b232-202">**resid**</span><span class="sxs-lookup"><span data-stu-id="5b232-202">**resid**</span></span>|<span data-ttu-id="5b232-203">string</span><span class="sxs-lookup"><span data-stu-id="5b232-203">string</span></span>|<span data-ttu-id="5b232-204">Oui</span><span class="sxs-lookup"><span data-stu-id="5b232-204">Yes</span></span>|<span data-ttu-id="5b232-205">Spécifie l’emplacement d’URL du fichier JavaScript qui doit remplacer l’emplacement d’URL du code HTML par défaut défini dans l’élément [Runtime](runtime.md) `resid` parent.</span><span class="sxs-lookup"><span data-stu-id="5b232-205">Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`.</span></span> <span data-ttu-id="5b232-206">Il ne peut pas y avoir plus de 32 caractères et doit correspondre à un `resid` `id` attribut `Url` d’un élément dans `Resources` l’élément.</span><span class="sxs-lookup"><span data-stu-id="5b232-206">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span>|

### <a name="examples"></a><span data-ttu-id="5b232-207">Exemples</span><span class="sxs-lookup"><span data-stu-id="5b232-207">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="5b232-208">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5b232-208">See also</span></span>

- [<span data-ttu-id="5b232-209">Runtime</span><span class="sxs-lookup"><span data-stu-id="5b232-209">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="5b232-210">Configurer votre complément Outlook pour l’activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="5b232-210">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
