---
title: Élément Override dans le fichier manifest
description: L’élément Override vous permet de spécifier la valeur d’un paramètre en fonction d’une condition spécifiée.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: d2146cc1f44e829bc78076c8093b2ebf791dc722
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505338"
---
# <a name="override-element"></a><span data-ttu-id="96654-103">Élément Override</span><span class="sxs-lookup"><span data-stu-id="96654-103">Override element</span></span>

<span data-ttu-id="96654-104">Permet de remplacer la valeur d’un paramètre de manifeste en fonction d’une condition spécifiée.</span><span class="sxs-lookup"><span data-stu-id="96654-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="96654-105">Il existe deux types de conditions :</span><span class="sxs-lookup"><span data-stu-id="96654-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="96654-106">Paramètres régionaux Office différents de la valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="96654-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="96654-107">Modèle de prise en charge de l’ensemble de conditions requises différent du modèle par défaut.</span><span class="sxs-lookup"><span data-stu-id="96654-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="96654-108">Il existe deux types d’éléments, l’un pour les remplacements de `<Override>` paramètres régionaux, appelé **LocaleTokenOverride** et l’autre pour les substitutions d’ensembles de conditions requises, appelé **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="96654-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride**, and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="96654-109">Mais il n’existe `type` aucun paramètre pour `<Override>` l’élément.</span><span class="sxs-lookup"><span data-stu-id="96654-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="96654-110">La différence est déterminée par l’élément parent et le type de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="96654-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="96654-111">Un `<Override>` élément qui se trouve à l’intérieur d’un élément dont , doit être de type `<Token>` `xsi:type` `RequirementToken` **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="96654-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="96654-112">Un élément à l’intérieur d’un autre élément parent, ou à l’intérieur d’un élément de type, doit `<Override>` `<Override>` être de type `LocaleToken` **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="96654-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="96654-113">Chaque type est décrit dans des sections distinctes ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="96654-113">Each type is described in separate sections below.</span></span> <span data-ttu-id="96654-114">Pour plus d’informations sur l’utilisation de cet élément lorsqu’il est enfant d’un élément, voir Work `<Token>` [with extended overrides of the manifest](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="96654-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="96654-115">Élément Override de type LocaleTokenOverride</span><span class="sxs-lookup"><span data-stu-id="96654-115">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="96654-116">Un `<Override>` élément exprime une conditionnel et peut être lu en tant que « If ... alors... .</span><span class="sxs-lookup"><span data-stu-id="96654-116">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="96654-117">Si `<Override>` l’élément est de type **LocaleTokenOverride**, l’attribut est la `Locale` condition et `Value` l’attribut en est la conséquence.</span><span class="sxs-lookup"><span data-stu-id="96654-117">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="96654-118">Par exemple, l’exemple suivant indique « Si le paramètre de paramètres régionaux Office est fr-fr, le nom complet est Lecteur vidéo ».</span><span class="sxs-lookup"><span data-stu-id="96654-118">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="96654-119">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="96654-119">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="96654-120">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="96654-120">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="96654-121">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="96654-121">Contained in</span></span>

|<span data-ttu-id="96654-122">Élément</span><span class="sxs-lookup"><span data-stu-id="96654-122">Element</span></span>|
|:-----|
|[<span data-ttu-id="96654-123">CitationText</span><span class="sxs-lookup"><span data-stu-id="96654-123">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="96654-124">Description</span><span class="sxs-lookup"><span data-stu-id="96654-124">Description</span></span>](description.md)|
|[<span data-ttu-id="96654-125">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="96654-125">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="96654-126">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="96654-126">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="96654-127">DisplayName</span><span class="sxs-lookup"><span data-stu-id="96654-127">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="96654-128">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="96654-128">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="96654-129">IconUrl</span><span class="sxs-lookup"><span data-stu-id="96654-129">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="96654-130">QueryUri</span><span class="sxs-lookup"><span data-stu-id="96654-130">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="96654-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="96654-131">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="96654-132">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="96654-132">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="96654-133">Jeton</span><span class="sxs-lookup"><span data-stu-id="96654-133">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="96654-134">Attributs</span><span class="sxs-lookup"><span data-stu-id="96654-134">Attributes</span></span>

|<span data-ttu-id="96654-135">Attribut</span><span class="sxs-lookup"><span data-stu-id="96654-135">Attribute</span></span>|<span data-ttu-id="96654-136">Type</span><span class="sxs-lookup"><span data-stu-id="96654-136">Type</span></span>|<span data-ttu-id="96654-137">Requis</span><span class="sxs-lookup"><span data-stu-id="96654-137">Required</span></span>|<span data-ttu-id="96654-138">Description</span><span class="sxs-lookup"><span data-stu-id="96654-138">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="96654-139">Paramètres régionaux</span><span class="sxs-lookup"><span data-stu-id="96654-139">Locale</span></span>|<span data-ttu-id="96654-140">string</span><span class="sxs-lookup"><span data-stu-id="96654-140">string</span></span>|<span data-ttu-id="96654-141">obligatoire</span><span class="sxs-lookup"><span data-stu-id="96654-141">required</span></span>|<span data-ttu-id="96654-142">Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="96654-142">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="96654-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="96654-143">Value</span></span>|<span data-ttu-id="96654-144">string</span><span class="sxs-lookup"><span data-stu-id="96654-144">string</span></span>|<span data-ttu-id="96654-145">obligatoire</span><span class="sxs-lookup"><span data-stu-id="96654-145">required</span></span>|<span data-ttu-id="96654-146">Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.</span><span class="sxs-lookup"><span data-stu-id="96654-146">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="96654-147">範例</span><span class="sxs-lookup"><span data-stu-id="96654-147">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="96654-148">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="96654-148">See also</span></span>

- [<span data-ttu-id="96654-149">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="96654-149">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="96654-150">Raccourcis clavier</span><span class="sxs-lookup"><span data-stu-id="96654-150">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="96654-151">Élément Override de type RequirementTokenOverride</span><span class="sxs-lookup"><span data-stu-id="96654-151">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="96654-152">Un `<Override>` élément exprime une conditionnel et peut être lu en tant que « If ... then ... » .</span><span class="sxs-lookup"><span data-stu-id="96654-152">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="96654-153">Si `<Override>` l’élément est de type **RequirementTokenOverride**, l’élément enfant exprime la condition et l’attribut `<Requirements>` en est le `Value` résultat.</span><span class="sxs-lookup"><span data-stu-id="96654-153">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="96654-154">Par exemple, la première partie de ce qui suit est lue « Si la plateforme actuelle prend en charge FeatureOne version 1.7, utilisez la chaîne « oldAddinVersion » à la place du jeton dans l’URL de l’enfant (au lieu de la chaîne par défaut `<Override>` `${token.requirements}` « upgrade `<ExtendedOverrides>` »).</span><span class="sxs-lookup"><span data-stu-id="96654-154">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="96654-155">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="96654-155">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="96654-156">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="96654-156">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="96654-157">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="96654-157">Contained in</span></span>

|<span data-ttu-id="96654-158">Élément</span><span class="sxs-lookup"><span data-stu-id="96654-158">Element</span></span>|
|:-----|
|[<span data-ttu-id="96654-159">Jeton</span><span class="sxs-lookup"><span data-stu-id="96654-159">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="96654-160">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="96654-160">Must contain</span></span>

|<span data-ttu-id="96654-161">Élément</span><span class="sxs-lookup"><span data-stu-id="96654-161">Element</span></span>|<span data-ttu-id="96654-162">Contenu</span><span class="sxs-lookup"><span data-stu-id="96654-162">Content</span></span>|<span data-ttu-id="96654-163">Courrier</span><span class="sxs-lookup"><span data-stu-id="96654-163">Mail</span></span>|<span data-ttu-id="96654-164">TaskPane</span><span class="sxs-lookup"><span data-stu-id="96654-164">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="96654-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96654-165">Requirements</span></span>](requirements.md)|||<span data-ttu-id="96654-166">x</span><span class="sxs-lookup"><span data-stu-id="96654-166">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="96654-167">Attributs</span><span class="sxs-lookup"><span data-stu-id="96654-167">Attributes</span></span>

|<span data-ttu-id="96654-168">Attribut</span><span class="sxs-lookup"><span data-stu-id="96654-168">Attribute</span></span>|<span data-ttu-id="96654-169">Type</span><span class="sxs-lookup"><span data-stu-id="96654-169">Type</span></span>|<span data-ttu-id="96654-170">Requis</span><span class="sxs-lookup"><span data-stu-id="96654-170">Required</span></span>|<span data-ttu-id="96654-171">Description</span><span class="sxs-lookup"><span data-stu-id="96654-171">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="96654-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="96654-172">Value</span></span>|<span data-ttu-id="96654-173">string</span><span class="sxs-lookup"><span data-stu-id="96654-173">string</span></span>|<span data-ttu-id="96654-174">obligatoire</span><span class="sxs-lookup"><span data-stu-id="96654-174">required</span></span>|<span data-ttu-id="96654-175">Valeur du jeton de preuve lorsque la condition est remplie.</span><span class="sxs-lookup"><span data-stu-id="96654-175">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="96654-176">Exemple</span><span class="sxs-lookup"><span data-stu-id="96654-176">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="96654-177">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="96654-177">See also</span></span>

- [<span data-ttu-id="96654-178">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="96654-178">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="96654-179">Définition de l’élément Requirements dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="96654-179">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="96654-180">Raccourcis clavier</span><span class="sxs-lookup"><span data-stu-id="96654-180">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
