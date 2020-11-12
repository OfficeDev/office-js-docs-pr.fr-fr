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
# <a name="override-element"></a><span data-ttu-id="ff736-103">Élément Override</span><span class="sxs-lookup"><span data-stu-id="ff736-103">Override element</span></span>

<span data-ttu-id="ff736-104">Permet de remplacer la valeur d’un paramètre de manifeste en fonction d’une condition spécifiée.</span><span class="sxs-lookup"><span data-stu-id="ff736-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="ff736-105">Il existe deux types de conditions :</span><span class="sxs-lookup"><span data-stu-id="ff736-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="ff736-106">Paramètres régionaux Office différents de ceux par défaut.</span><span class="sxs-lookup"><span data-stu-id="ff736-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="ff736-107">Modèle de la prise en charge de l’ensemble de conditions requises qui est différente du modèle par défaut.</span><span class="sxs-lookup"><span data-stu-id="ff736-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="ff736-108">Il existe deux types d' `<Override>` éléments : un pour les substitutions de paramètres régionaux, appelé **LocaleTokenOverride** , et l’autre pour les substitutions d’ensemble de conditions requises, appelé **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ff736-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride** , and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="ff736-109">Mais il n’existe aucun `type` paramètre pour l' `<Override>` élément.</span><span class="sxs-lookup"><span data-stu-id="ff736-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="ff736-110">La différence est déterminée par l’élément parent et le type de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="ff736-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="ff736-111">Un `<Override>` élément qui se trouve à l’intérieur d’un `<Token>` élément dont le `xsi:type` est `RequirementToken` , doit être de type **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ff736-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="ff736-112">Un `<Override>` élément situé à l’intérieur d’un autre élément parent, ou à l’intérieur d’un `<Override>` élément de type `LocaleToken` , doit être de type **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ff736-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="ff736-113">Chaque type est décrit dans des sections distinctes ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="ff736-113">Each type is described in separate sections below.</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="ff736-114">Élément override de type LocaleTokenOverride</span><span class="sxs-lookup"><span data-stu-id="ff736-114">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="ff736-115">Un `<Override>` élément exprime un conditionnel et peut être lu sous la forme d’un «if... Then... " résultat.</span><span class="sxs-lookup"><span data-stu-id="ff736-115">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="ff736-116">Si l' `<Override>` élément est de type **LocaleTokenOverride** , l' `Locale` attribut est la condition, et l' `Value` attribut est le à la suite.</span><span class="sxs-lookup"><span data-stu-id="ff736-116">If the `<Override>` element is of type **LocaleTokenOverride** , then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="ff736-117">Par exemple, le code suivant est lu « si le paramètre paramètres régionaux Office est fr-fr, le nom complet est «lecteur vidéo ».»</span><span class="sxs-lookup"><span data-stu-id="ff736-117">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="ff736-118">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="ff736-118">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="ff736-119">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ff736-119">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="ff736-120">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ff736-120">Contained in</span></span>

|<span data-ttu-id="ff736-121">Élément</span><span class="sxs-lookup"><span data-stu-id="ff736-121">Element</span></span>|
|:-----|
|[<span data-ttu-id="ff736-122">CitationText</span><span class="sxs-lookup"><span data-stu-id="ff736-122">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="ff736-123">Description</span><span class="sxs-lookup"><span data-stu-id="ff736-123">Description</span></span>](description.md)|
|[<span data-ttu-id="ff736-124">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="ff736-124">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="ff736-125">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="ff736-125">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="ff736-126">DisplayName</span><span class="sxs-lookup"><span data-stu-id="ff736-126">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="ff736-127">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="ff736-127">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="ff736-128">IconUrl</span><span class="sxs-lookup"><span data-stu-id="ff736-128">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="ff736-129">QueryUri</span><span class="sxs-lookup"><span data-stu-id="ff736-129">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="ff736-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ff736-130">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="ff736-131">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="ff736-131">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="ff736-132">Jeton</span><span class="sxs-lookup"><span data-stu-id="ff736-132">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="ff736-133">Attributs</span><span class="sxs-lookup"><span data-stu-id="ff736-133">Attributes</span></span>

|<span data-ttu-id="ff736-134">Attribut</span><span class="sxs-lookup"><span data-stu-id="ff736-134">Attribute</span></span>|<span data-ttu-id="ff736-135">Type</span><span class="sxs-lookup"><span data-stu-id="ff736-135">Type</span></span>|<span data-ttu-id="ff736-136">Requis</span><span class="sxs-lookup"><span data-stu-id="ff736-136">Required</span></span>|<span data-ttu-id="ff736-137">Description</span><span class="sxs-lookup"><span data-stu-id="ff736-137">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ff736-138">Paramètres régionaux</span><span class="sxs-lookup"><span data-stu-id="ff736-138">Locale</span></span>|<span data-ttu-id="ff736-139">string</span><span class="sxs-lookup"><span data-stu-id="ff736-139">string</span></span>|<span data-ttu-id="ff736-140">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ff736-140">required</span></span>|<span data-ttu-id="ff736-141">Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="ff736-141">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="ff736-142">Valeur</span><span class="sxs-lookup"><span data-stu-id="ff736-142">Value</span></span>|<span data-ttu-id="ff736-143">string</span><span class="sxs-lookup"><span data-stu-id="ff736-143">string</span></span>|<span data-ttu-id="ff736-144">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ff736-144">required</span></span>|<span data-ttu-id="ff736-145">Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.</span><span class="sxs-lookup"><span data-stu-id="ff736-145">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="ff736-146">範例</span><span class="sxs-lookup"><span data-stu-id="ff736-146">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="ff736-147">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ff736-147">See also</span></span>

- [<span data-ttu-id="ff736-148">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="ff736-148">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="ff736-149">Raccourcis clavier</span><span class="sxs-lookup"><span data-stu-id="ff736-149">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="ff736-150">Élément override de type RequirementTokenOverride</span><span class="sxs-lookup"><span data-stu-id="ff736-150">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="ff736-151">Un `<Override>` élément exprime un conditionnel et peut être lu sous la forme d’un «if... Then... " résultat.</span><span class="sxs-lookup"><span data-stu-id="ff736-151">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="ff736-152">Si l' `<Override>` élément est de type **RequirementTokenOverride** , l’élément enfant `<Requirements>` exprime la condition, et l' `Value` attribut est le à la suite.</span><span class="sxs-lookup"><span data-stu-id="ff736-152">If the `<Override>` element is of type **RequirementTokenOverride** , then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="ff736-153">Par exemple, le premier `<Override>` des éléments suivants est lu « si la plateforme actuelle prend en charge la version 1,7 de FeatureOne, puis utilisez la chaîne «oldAddinVersion » à la place du `${token.requirements}` jeton dans l’URL du grand-parent `<ExtendedOverrides>` (au lieu de la chaîne par défaut « mise à niveau »)».</span><span class="sxs-lookup"><span data-stu-id="ff736-153">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="ff736-154">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="ff736-154">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="ff736-155">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ff736-155">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="ff736-156">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ff736-156">Contained in</span></span>

|<span data-ttu-id="ff736-157">Élément</span><span class="sxs-lookup"><span data-stu-id="ff736-157">Element</span></span>|
|:-----|
|[<span data-ttu-id="ff736-158">Jeton</span><span class="sxs-lookup"><span data-stu-id="ff736-158">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="ff736-159">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="ff736-159">Must contain</span></span>

|<span data-ttu-id="ff736-160">Élément</span><span class="sxs-lookup"><span data-stu-id="ff736-160">Element</span></span>|<span data-ttu-id="ff736-161">Contenu</span><span class="sxs-lookup"><span data-stu-id="ff736-161">Content</span></span>|<span data-ttu-id="ff736-162">Courrier</span><span class="sxs-lookup"><span data-stu-id="ff736-162">Mail</span></span>|<span data-ttu-id="ff736-163">TaskPane</span><span class="sxs-lookup"><span data-stu-id="ff736-163">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="ff736-164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ff736-164">Requirements</span></span>](requirements.md)|||<span data-ttu-id="ff736-165">x</span><span class="sxs-lookup"><span data-stu-id="ff736-165">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="ff736-166">Attributs</span><span class="sxs-lookup"><span data-stu-id="ff736-166">Attributes</span></span>

|<span data-ttu-id="ff736-167">Attribut</span><span class="sxs-lookup"><span data-stu-id="ff736-167">Attribute</span></span>|<span data-ttu-id="ff736-168">Type</span><span class="sxs-lookup"><span data-stu-id="ff736-168">Type</span></span>|<span data-ttu-id="ff736-169">Requis</span><span class="sxs-lookup"><span data-stu-id="ff736-169">Required</span></span>|<span data-ttu-id="ff736-170">Description</span><span class="sxs-lookup"><span data-stu-id="ff736-170">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ff736-171">Valeur</span><span class="sxs-lookup"><span data-stu-id="ff736-171">Value</span></span>|<span data-ttu-id="ff736-172">string</span><span class="sxs-lookup"><span data-stu-id="ff736-172">string</span></span>|<span data-ttu-id="ff736-173">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ff736-173">required</span></span>|<span data-ttu-id="ff736-174">Valeur du jeton de grand-parent lorsque la condition est satisfaite.</span><span class="sxs-lookup"><span data-stu-id="ff736-174">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="ff736-175">Exemple</span><span class="sxs-lookup"><span data-stu-id="ff736-175">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="ff736-176">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ff736-176">See also</span></span>

- [<span data-ttu-id="ff736-177">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ff736-177">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ff736-178">Définition de l’élément Requirements dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="ff736-178">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="ff736-179">Raccourcis clavier</span><span class="sxs-lookup"><span data-stu-id="ff736-179">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
