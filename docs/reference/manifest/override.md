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
# <a name="override-element"></a><span data-ttu-id="ee85c-103">Élément Override</span><span class="sxs-lookup"><span data-stu-id="ee85c-103">Override element</span></span>

<span data-ttu-id="ee85c-104">Fournit un moyen de passer outre à la valeur d’un paramètre manifeste en fonction d’une condition spécifiée.</span><span class="sxs-lookup"><span data-stu-id="ee85c-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="ee85c-105">Il existe trois types de conditions :</span><span class="sxs-lookup"><span data-stu-id="ee85c-105">There are three kinds of conditions:</span></span>

- <span data-ttu-id="ee85c-106">Un Office local qui est différent de la valeur `LocaleToken` par défaut , appelé **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ee85c-106">An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.</span></span>
- <span data-ttu-id="ee85c-107">Un modèle de support d’ensemble d’exigences différent du `RequirementToken` modèle par défaut, **appelé RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ee85c-107">A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.</span></span>
- <span data-ttu-id="ee85c-108">La source est différente de la valeur par `Runtime` défaut , **appelée RuntimeOverride** (actuellement en avant-première).</span><span class="sxs-lookup"><span data-stu-id="ee85c-108">The source is different from the default `Runtime`, called **RuntimeOverride** (currently in preview).</span></span>

<span data-ttu-id="ee85c-109">Un `<Override>` élément qui est à l’intérieur `<Runtime>` d’un élément doit être de type **RuntimeOverride**.</span><span class="sxs-lookup"><span data-stu-id="ee85c-109">An `<Override>` element that is inside of a `<Runtime>` element must be of type **RuntimeOverride**.</span></span>

<span data-ttu-id="ee85c-110">Il n’y a `overrideType` pas d’attribut pour `<Override>` l’élément.</span><span class="sxs-lookup"><span data-stu-id="ee85c-110">There is no `overrideType` attribute for the `<Override>` element.</span></span> <span data-ttu-id="ee85c-111">La différence est déterminée par l’élément parent et le type de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="ee85c-111">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="ee85c-112">Un `<Override>` élément qui est à l’intérieur `<Token>` `xsi:type` d’un élément qui est , doit être de type `RequirementToken` **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ee85c-112">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="ee85c-113">Un `<Override>` élément à l’intérieur de tout autre élément parent, ou à `<Override>` l’intérieur `LocaleToken` d’un élément de type , doit être de type **LocalTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ee85c-113">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="ee85c-114">Pour plus d’informations sur l’utilisation de cet élément lorsqu’il s’agit d’un enfant `<Token>` [d’un élément, voir Travail avec des dérogations étendues du manifeste](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="ee85c-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="ee85c-115">Chaque type est décrit dans des sections distinctes plus tard dans cet article.</span><span class="sxs-lookup"><span data-stu-id="ee85c-115">Each type is described in separate sections later in this article.</span></span>

## <a name="override-element-for-localetoken"></a><span data-ttu-id="ee85c-116">Élément de remplacement pour `LocaleToken`</span><span class="sxs-lookup"><span data-stu-id="ee85c-116">Override element for `LocaleToken`</span></span>

<span data-ttu-id="ee85c-117">Un `<Override>` élément exprime un conditionnel et peut être lu comme un « Si ... puis ... » déclaration.</span><span class="sxs-lookup"><span data-stu-id="ee85c-117">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="ee85c-118">Si `<Override>` l’élément est de type **LocalTokenOverride**, alors `Locale` l’attribut est la condition, et l’attribut `Value` est le conséquent.</span><span class="sxs-lookup"><span data-stu-id="ee85c-118">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="ee85c-119">Par exemple, ce qui suit est lu " Si le paramètre Office local est fr-fr, alors le nom de l’affichage est 'Lecteur vidéo'. »</span><span class="sxs-lookup"><span data-stu-id="ee85c-119">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="ee85c-120">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="ee85c-120">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="ee85c-121">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ee85c-121">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="ee85c-122">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ee85c-122">Contained in</span></span>

|<span data-ttu-id="ee85c-123">Élément</span><span class="sxs-lookup"><span data-stu-id="ee85c-123">Element</span></span>|
|:-----|
|[<span data-ttu-id="ee85c-124">CitationText</span><span class="sxs-lookup"><span data-stu-id="ee85c-124">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="ee85c-125">Description</span><span class="sxs-lookup"><span data-stu-id="ee85c-125">Description</span></span>](description.md)|
|[<span data-ttu-id="ee85c-126">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="ee85c-126">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="ee85c-127">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="ee85c-127">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="ee85c-128">DisplayName</span><span class="sxs-lookup"><span data-stu-id="ee85c-128">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="ee85c-129">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="ee85c-129">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="ee85c-130">IconUrl</span><span class="sxs-lookup"><span data-stu-id="ee85c-130">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="ee85c-131">QueryUri</span><span class="sxs-lookup"><span data-stu-id="ee85c-131">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="ee85c-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ee85c-132">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="ee85c-133">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="ee85c-133">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="ee85c-134">Jeton</span><span class="sxs-lookup"><span data-stu-id="ee85c-134">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="ee85c-135">Attributs</span><span class="sxs-lookup"><span data-stu-id="ee85c-135">Attributes</span></span>

|<span data-ttu-id="ee85c-136">Attribut</span><span class="sxs-lookup"><span data-stu-id="ee85c-136">Attribute</span></span>|<span data-ttu-id="ee85c-137">Type</span><span class="sxs-lookup"><span data-stu-id="ee85c-137">Type</span></span>|<span data-ttu-id="ee85c-138">Requis</span><span class="sxs-lookup"><span data-stu-id="ee85c-138">Required</span></span>|<span data-ttu-id="ee85c-139">Description</span><span class="sxs-lookup"><span data-stu-id="ee85c-139">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ee85c-140">Paramètres régionaux</span><span class="sxs-lookup"><span data-stu-id="ee85c-140">Locale</span></span>|<span data-ttu-id="ee85c-141">string</span><span class="sxs-lookup"><span data-stu-id="ee85c-141">string</span></span>|<span data-ttu-id="ee85c-142">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ee85c-142">required</span></span>|<span data-ttu-id="ee85c-143">Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="ee85c-143">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="ee85c-144">Valeur</span><span class="sxs-lookup"><span data-stu-id="ee85c-144">Value</span></span>|<span data-ttu-id="ee85c-145">string</span><span class="sxs-lookup"><span data-stu-id="ee85c-145">string</span></span>|<span data-ttu-id="ee85c-146">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ee85c-146">required</span></span>|<span data-ttu-id="ee85c-147">Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.</span><span class="sxs-lookup"><span data-stu-id="ee85c-147">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="ee85c-148">Exemples</span><span class="sxs-lookup"><span data-stu-id="ee85c-148">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="ee85c-149">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ee85c-149">See also</span></span>

- [<span data-ttu-id="ee85c-150">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="ee85c-150">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="ee85c-151">Raccourcis clavier</span><span class="sxs-lookup"><span data-stu-id="ee85c-151">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a><span data-ttu-id="ee85c-152">Élément de remplacement pour `RequirementToken`</span><span class="sxs-lookup"><span data-stu-id="ee85c-152">Override element for `RequirementToken`</span></span>

<span data-ttu-id="ee85c-153">Un `<Override>` élément exprime un conditionnel et peut être lu comme un « Si ... puis ... » déclaration.</span><span class="sxs-lookup"><span data-stu-id="ee85c-153">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="ee85c-154">Si `<Override>` l’élément est de type **RequirementTokenOverride**, alors l’élément `<Requirements>` enfant exprime la condition, et l’attribut `Value` est le conséquent.</span><span class="sxs-lookup"><span data-stu-id="ee85c-154">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="ee85c-155">Par exemple, le premier élément suivant est lu « Si la plate-forme actuelle `<Override>` prend en charge la version FeatureOne 1.7, utilisez la chaîne « oldAddinVersion » à la place du `${token.requirements}` jeton dans l’URL du grand-parent (au lieu de la `<ExtendedOverrides>` chaîne par défaut « mise à niveau »). »</span><span class="sxs-lookup"><span data-stu-id="ee85c-155">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="ee85c-156">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="ee85c-156">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="ee85c-157">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ee85c-157">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="ee85c-158">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ee85c-158">Contained in</span></span>

|<span data-ttu-id="ee85c-159">Élément</span><span class="sxs-lookup"><span data-stu-id="ee85c-159">Element</span></span>|
|:-----|
|[<span data-ttu-id="ee85c-160">Jeton</span><span class="sxs-lookup"><span data-stu-id="ee85c-160">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="ee85c-161">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="ee85c-161">Must contain</span></span>

|<span data-ttu-id="ee85c-162">Élément</span><span class="sxs-lookup"><span data-stu-id="ee85c-162">Element</span></span>|<span data-ttu-id="ee85c-163">Contenu</span><span class="sxs-lookup"><span data-stu-id="ee85c-163">Content</span></span>|<span data-ttu-id="ee85c-164">Courrier</span><span class="sxs-lookup"><span data-stu-id="ee85c-164">Mail</span></span>|<span data-ttu-id="ee85c-165">TaskPane</span><span class="sxs-lookup"><span data-stu-id="ee85c-165">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="ee85c-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ee85c-166">Requirements</span></span>](requirements.md)|||<span data-ttu-id="ee85c-167">x</span><span class="sxs-lookup"><span data-stu-id="ee85c-167">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="ee85c-168">Attributs</span><span class="sxs-lookup"><span data-stu-id="ee85c-168">Attributes</span></span>

|<span data-ttu-id="ee85c-169">Attribut</span><span class="sxs-lookup"><span data-stu-id="ee85c-169">Attribute</span></span>|<span data-ttu-id="ee85c-170">Type</span><span class="sxs-lookup"><span data-stu-id="ee85c-170">Type</span></span>|<span data-ttu-id="ee85c-171">Requis</span><span class="sxs-lookup"><span data-stu-id="ee85c-171">Required</span></span>|<span data-ttu-id="ee85c-172">Description</span><span class="sxs-lookup"><span data-stu-id="ee85c-172">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ee85c-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="ee85c-173">Value</span></span>|<span data-ttu-id="ee85c-174">string</span><span class="sxs-lookup"><span data-stu-id="ee85c-174">string</span></span>|<span data-ttu-id="ee85c-175">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ee85c-175">required</span></span>|<span data-ttu-id="ee85c-176">Valeur du jeton grand-parent lorsque la condition est remplie.</span><span class="sxs-lookup"><span data-stu-id="ee85c-176">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="ee85c-177">Exemple</span><span class="sxs-lookup"><span data-stu-id="ee85c-177">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="ee85c-178">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ee85c-178">See also</span></span>

- [<span data-ttu-id="ee85c-179">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ee85c-179">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ee85c-180">Définition de l’élément Requirements dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="ee85c-180">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="ee85c-181">Raccourcis clavier</span><span class="sxs-lookup"><span data-stu-id="ee85c-181">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime-preview"></a><span data-ttu-id="ee85c-182">Élément de remplacement pour `Runtime` (aperçu)</span><span class="sxs-lookup"><span data-stu-id="ee85c-182">Override element for `Runtime` (preview)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ee85c-183">Cette fonctionnalité n’est prise en [charge que](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) pour un aperçu Outlook sur le web et sur Windows avec un abonnement Microsoft 365 spécial.</span><span class="sxs-lookup"><span data-stu-id="ee85c-183">This feature is only supported for [preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="ee85c-184">Pour plus de détails, [consultez Configurez votre Outlook add-in pour l’activation basée sur l’événement](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="ee85c-184">For more details, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>
>
> <span data-ttu-id="ee85c-185">Étant donné que les fonctionnalités d’aperçu sont sujettes à changement sans préavis, elles ne doivent pas être utilisées dans les modules de production.</span><span class="sxs-lookup"><span data-stu-id="ee85c-185">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

<span data-ttu-id="ee85c-186">Un `<Override>` élément exprime un conditionnel et peut être lu comme un « Si ... puis ... » déclaration.</span><span class="sxs-lookup"><span data-stu-id="ee85c-186">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="ee85c-187">Si `<Override>` l’élément est de type **RuntimeOverride**, alors `type` l’attribut est la condition, et l’attribut `resid` est le conséquent.</span><span class="sxs-lookup"><span data-stu-id="ee85c-187">If the `<Override>` element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent.</span></span> <span data-ttu-id="ee85c-188">Par exemple, ce qui suit est lu « Si le type est 'javascript', alors `resid` le est 'JSRuntime.Url'. » Outlook Desktop nécessite cet élément pour les [gestionnaires de points d’extension LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview)</span><span class="sxs-lookup"><span data-stu-id="ee85c-188">For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent-preview) handlers.</span></span>

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

<span data-ttu-id="ee85c-189">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="ee85c-189">**Add-in type:** Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="ee85c-190">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ee85c-190">Syntax</span></span>

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a><span data-ttu-id="ee85c-191">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ee85c-191">Contained in</span></span>

- [<span data-ttu-id="ee85c-192">Runtime</span><span class="sxs-lookup"><span data-stu-id="ee85c-192">Runtime</span></span>](runtime.md)

### <a name="attributes"></a><span data-ttu-id="ee85c-193">Attributs</span><span class="sxs-lookup"><span data-stu-id="ee85c-193">Attributes</span></span>

|<span data-ttu-id="ee85c-194">Attribut</span><span class="sxs-lookup"><span data-stu-id="ee85c-194">Attribute</span></span>|<span data-ttu-id="ee85c-195">Type</span><span class="sxs-lookup"><span data-stu-id="ee85c-195">Type</span></span>|<span data-ttu-id="ee85c-196">Requis</span><span class="sxs-lookup"><span data-stu-id="ee85c-196">Required</span></span>|<span data-ttu-id="ee85c-197">Description</span><span class="sxs-lookup"><span data-stu-id="ee85c-197">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ee85c-198">**type**</span><span class="sxs-lookup"><span data-stu-id="ee85c-198">**type**</span></span>|<span data-ttu-id="ee85c-199">string</span><span class="sxs-lookup"><span data-stu-id="ee85c-199">string</span></span>|<span data-ttu-id="ee85c-200">Oui</span><span class="sxs-lookup"><span data-stu-id="ee85c-200">Yes</span></span>|<span data-ttu-id="ee85c-201">Spécifie la langue pour cette substitution.</span><span class="sxs-lookup"><span data-stu-id="ee85c-201">Specifies the language for this override.</span></span> <span data-ttu-id="ee85c-202">À l’heure `"javascript"` actuelle, est la seule option prise en charge.</span><span class="sxs-lookup"><span data-stu-id="ee85c-202">At present, `"javascript"` is the only supported option.</span></span>|
|<span data-ttu-id="ee85c-203">**resid**</span><span class="sxs-lookup"><span data-stu-id="ee85c-203">**resid**</span></span>|<span data-ttu-id="ee85c-204">string</span><span class="sxs-lookup"><span data-stu-id="ee85c-204">string</span></span>|<span data-ttu-id="ee85c-205">Oui</span><span class="sxs-lookup"><span data-stu-id="ee85c-205">Yes</span></span>|<span data-ttu-id="ee85c-206">Spécifie l’emplacement de l’URL du fichier JavaScript qui doit passer outre à l’emplacement de l’URL du HTML par défaut défini dans [l’élément runtime](runtime.md) parent `resid` .</span><span class="sxs-lookup"><span data-stu-id="ee85c-206">Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`.</span></span> <span data-ttu-id="ee85c-207">Le `resid` ne peut pas être plus de 32 caractères et doit correspondre à un attribut `id` d’un `Url` élément dans `Resources` l’élément.</span><span class="sxs-lookup"><span data-stu-id="ee85c-207">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span>|

### <a name="examples"></a><span data-ttu-id="ee85c-208">Exemples</span><span class="sxs-lookup"><span data-stu-id="ee85c-208">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="ee85c-209">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ee85c-209">See also</span></span>

- [<span data-ttu-id="ee85c-210">Runtime</span><span class="sxs-lookup"><span data-stu-id="ee85c-210">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="ee85c-211">Configurez votre Outlook add-in pour l’activation basée sur l’événement</span><span class="sxs-lookup"><span data-stu-id="ee85c-211">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
