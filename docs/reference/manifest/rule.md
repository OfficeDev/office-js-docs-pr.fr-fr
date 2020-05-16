---
title: Élément Rule dans le fichier manifeste
description: L’élément rule spécifie les règles d’activation qui doivent être évaluées pour ce complément de messagerie contextuel.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: c1f2be3371333bfd87e0693d02a9a5984c18317b
ms.sourcegitcommit: b634bfe9a946fbd95754e87f070a904ed57586ff
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/15/2020
ms.locfileid: "44253620"
---
# <a name="rule-element"></a><span data-ttu-id="7209a-103">Élément Rule</span><span class="sxs-lookup"><span data-stu-id="7209a-103">Rule element</span></span>

<span data-ttu-id="7209a-104">Spécifie les règles d’activation qui doivent être évaluées pour ce complément de messagerie contextuel.</span><span class="sxs-lookup"><span data-stu-id="7209a-104">Specifies the activation rules that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="7209a-105">**Type de complément :** Courrier (contextuel)</span><span class="sxs-lookup"><span data-stu-id="7209a-105">**Add-in type:** Mail (contextual)</span></span>

## <a name="contained-in"></a><span data-ttu-id="7209a-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="7209a-106">Contained in</span></span>

- [<span data-ttu-id="7209a-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="7209a-107">OfficeApp</span></span>](officeapp.md)
- <span data-ttu-id="7209a-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (déconseillé)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span><span class="sxs-lookup"><span data-stu-id="7209a-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span></span>

## <a name="attributes"></a><span data-ttu-id="7209a-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="7209a-109">Attributes</span></span>

| <span data-ttu-id="7209a-110">Attribut</span><span class="sxs-lookup"><span data-stu-id="7209a-110">Attribute</span></span> | <span data-ttu-id="7209a-111">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7209a-111">Required</span></span> | <span data-ttu-id="7209a-112">Description</span><span class="sxs-lookup"><span data-stu-id="7209a-112">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7209a-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="7209a-113">**xsi:type**</span></span> | <span data-ttu-id="7209a-114">Oui</span><span class="sxs-lookup"><span data-stu-id="7209a-114">Yes</span></span> | <span data-ttu-id="7209a-115">Type de règle en cours de définition.</span><span class="sxs-lookup"><span data-stu-id="7209a-115">The type of rule being defined.</span></span> |

<span data-ttu-id="7209a-116">Le type de règle peut correspondre à l’une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="7209a-116">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="7209a-117">ItemIs</span><span class="sxs-lookup"><span data-stu-id="7209a-117">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="7209a-118">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="7209a-118">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="7209a-119">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="7209a-119">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="7209a-120">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="7209a-120">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="7209a-121">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="7209a-121">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="7209a-122">Règle ItemIs</span><span class="sxs-lookup"><span data-stu-id="7209a-122">ItemIs rule</span></span>

<span data-ttu-id="7209a-123">Définit une règle qui donne la valeur true si l’élément sélectionné est du type spécifié.</span><span class="sxs-lookup"><span data-stu-id="7209a-123">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="7209a-124">Attributs</span><span class="sxs-lookup"><span data-stu-id="7209a-124">Attributes</span></span>

| <span data-ttu-id="7209a-125">Attribut</span><span class="sxs-lookup"><span data-stu-id="7209a-125">Attribute</span></span> | <span data-ttu-id="7209a-126">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7209a-126">Required</span></span> | <span data-ttu-id="7209a-127">Description</span><span class="sxs-lookup"><span data-stu-id="7209a-127">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7209a-128">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="7209a-128">**ItemType**</span></span> | <span data-ttu-id="7209a-129">Oui</span><span class="sxs-lookup"><span data-stu-id="7209a-129">Yes</span></span> | <span data-ttu-id="7209a-p101">Spécifie le type d’élément à mettre en correspondance. Peut être `Message` ou `Appointment`. Le type d’élément `Message` inclut e-mails, demandes de réunion, réponses à une demande de réunion et annulations de réunion.</span><span class="sxs-lookup"><span data-stu-id="7209a-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="7209a-133">**FormType**</span><span class="sxs-lookup"><span data-stu-id="7209a-133">**FormType**</span></span> | <span data-ttu-id="7209a-134">Non (dans [ExtensionPoint](extensionpoint.md)), Oui (dans [App_office](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="7209a-134">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="7209a-p102">Spécifie si l’application doit apparaître dans le formulaire de lecture ou de modification pour l’élément. Peut correspondre à l’une des valeurs suivantes : `Read`, `Edit`, `ReadOrEdit`. Si spécifiée dans un `Rule` dans un `ExtensionPoint`, cette valeur DOIT être `Read`.</span><span class="sxs-lookup"><span data-stu-id="7209a-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="7209a-138">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="7209a-138">**ItemClass**</span></span> | <span data-ttu-id="7209a-139">Non</span><span class="sxs-lookup"><span data-stu-id="7209a-139">No</span></span> | <span data-ttu-id="7209a-p103">Spécifie la classe de message personnalisé à mettre en correspondance. Pour plus d’informations, voir l’article relatif à l’[activation d’un complément de messagerie dans Outlook pour une classe de message spécifique](../../outlook/activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="7209a-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](../../outlook/activation-rules.md).</span></span> |
| <span data-ttu-id="7209a-142">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="7209a-142">**IncludeSubClasses**</span></span> | <span data-ttu-id="7209a-143">Non</span><span class="sxs-lookup"><span data-stu-id="7209a-143">No</span></span> | <span data-ttu-id="7209a-144">Spécifie si la règle doit donner la valeur true si l’élément est une sous-classe de la classe de message spécifiée ; par défaut, la valeur est `false`.</span><span class="sxs-lookup"><span data-stu-id="7209a-144">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="7209a-145">Exemple</span><span class="sxs-lookup"><span data-stu-id="7209a-145">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="7209a-146">Règle ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="7209a-146">ItemHasAttachment rule</span></span>

<span data-ttu-id="7209a-147">Définit une règle qui donne la valeur true si l’élément contient une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="7209a-147">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="7209a-148">Exemple</span><span class="sxs-lookup"><span data-stu-id="7209a-148">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="7209a-149">Règle ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="7209a-149">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="7209a-150">Définit une règle qui donne la valeur true si l’élément contient dans son objet ou son corps du texte correspondant au type d’entité spécifié.</span><span class="sxs-lookup"><span data-stu-id="7209a-150">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="7209a-151">Attributs</span><span class="sxs-lookup"><span data-stu-id="7209a-151">Attributes</span></span>

| <span data-ttu-id="7209a-152">Attribut</span><span class="sxs-lookup"><span data-stu-id="7209a-152">Attribute</span></span> | <span data-ttu-id="7209a-153">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7209a-153">Required</span></span> | <span data-ttu-id="7209a-154">Description</span><span class="sxs-lookup"><span data-stu-id="7209a-154">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7209a-155">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="7209a-155">**EntityType**</span></span> | <span data-ttu-id="7209a-156">Oui</span><span class="sxs-lookup"><span data-stu-id="7209a-156">Yes</span></span> | <span data-ttu-id="7209a-p104">Spécifie le type d’entité à rechercher pour que la règle donne la valeur true. Peut correspondre à l’une des valeurs suivantes : `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` ou `Contact`.</span><span class="sxs-lookup"><span data-stu-id="7209a-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="7209a-159">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="7209a-159">**RegExFilter**</span></span> | <span data-ttu-id="7209a-160">Non</span><span class="sxs-lookup"><span data-stu-id="7209a-160">No</span></span> | <span data-ttu-id="7209a-161">Spécifie une expression régulière à exécuter par rapport à cette entité à des fins d’activation.</span><span class="sxs-lookup"><span data-stu-id="7209a-161">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="7209a-162">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="7209a-162">**FilterName**</span></span> | <span data-ttu-id="7209a-163">Non</span><span class="sxs-lookup"><span data-stu-id="7209a-163">No</span></span> | <span data-ttu-id="7209a-164">Spécifie le nom du filtre d’expression régulière, afin qu’il soit possible par la suite de s’y référer dans le code de votre complément.</span><span class="sxs-lookup"><span data-stu-id="7209a-164">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="7209a-165">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="7209a-165">**IgnoreCase**</span></span> | <span data-ttu-id="7209a-166">Non</span><span class="sxs-lookup"><span data-stu-id="7209a-166">No</span></span> | <span data-ttu-id="7209a-167">Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par l’attribut **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="7209a-167">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="7209a-168">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="7209a-168">**Highlight**</span></span> | <span data-ttu-id="7209a-169">Non</span><span class="sxs-lookup"><span data-stu-id="7209a-169">No</span></span> | <span data-ttu-id="7209a-p105">**Remarque :** cela s’applique uniquement aux éléments **Rule** au sein des éléments **ExtensionPoint**. Spécifie comment le client doit mettre en surbrillance les entités correspondantes. Peut correspondre à l’une des valeurs suivantes : `all` ou `none`. Si non spécifié, la valeur par défaut est `all`.</span><span class="sxs-lookup"><span data-stu-id="7209a-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="7209a-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="7209a-174">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="7209a-175">Règle ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="7209a-175">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="7209a-176">Définit une règle qui donne la valeur true si une correspondance de l’expression régulière spécifiée est trouvée dans la propriété spécifiée de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7209a-176">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="7209a-177">Attributs</span><span class="sxs-lookup"><span data-stu-id="7209a-177">Attributes</span></span>

| <span data-ttu-id="7209a-178">Attribut</span><span class="sxs-lookup"><span data-stu-id="7209a-178">Attribute</span></span> | <span data-ttu-id="7209a-179">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7209a-179">Required</span></span> | <span data-ttu-id="7209a-180">Description</span><span class="sxs-lookup"><span data-stu-id="7209a-180">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7209a-181">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="7209a-181">**RegExName**</span></span> | <span data-ttu-id="7209a-182">Oui</span><span class="sxs-lookup"><span data-stu-id="7209a-182">Yes</span></span> | <span data-ttu-id="7209a-183">Spécifie le nom de l’expression régulière afin que vous puissiez vous référer à l’expression dans le code de votre complément.</span><span class="sxs-lookup"><span data-stu-id="7209a-183">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="7209a-184">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="7209a-184">**RegExValue**</span></span> | <span data-ttu-id="7209a-185">Oui</span><span class="sxs-lookup"><span data-stu-id="7209a-185">Yes</span></span> | <span data-ttu-id="7209a-186">Spécifie l’expression régulière qui sera évaluée pour déterminer si le complément de messagerie doit être affiché.</span><span class="sxs-lookup"><span data-stu-id="7209a-186">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="7209a-187">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="7209a-187">**PropertyName**</span></span> | <span data-ttu-id="7209a-188">Oui</span><span class="sxs-lookup"><span data-stu-id="7209a-188">Yes</span></span> | <span data-ttu-id="7209a-p106">Spécifie le nom de la propriété par rapport à laquelle l’expression sera évaluée. Les options disponibles sont les suivantes : `Subject`, `BodyAsPlaintext`, `BodyAsHTML` ou `SenderSMTPAddress`.</span><span class="sxs-lookup"><span data-stu-id="7209a-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="7209a-191">Si vous spécifiez `BodyAsHTML`, Outlook applique seulement l’expression régulière si le corps de l’élément est du code HTML.</span><span class="sxs-lookup"><span data-stu-id="7209a-191">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="7209a-192">Si ce n’est pas le cas, Outlook ne renvoie aucune correspondance pour cette expression régulière.</span><span class="sxs-lookup"><span data-stu-id="7209a-192">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="7209a-193">Si vous spécifiez `BodyAsPlaintext`, Outlook applique toujours l’expression régulière au corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7209a-193">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="7209a-194">**Remarque :** vous devez donner la valeur `BodyAsPlaintext` à l’attribut **PropertyName** si vous spécifiez l’attribut **Highlight** pour l’élément **Rule**.</span><span class="sxs-lookup"><span data-stu-id="7209a-194">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="7209a-195">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="7209a-195">**IgnoreCase**</span></span> | <span data-ttu-id="7209a-196">Non</span><span class="sxs-lookup"><span data-stu-id="7209a-196">No</span></span> | <span data-ttu-id="7209a-197">Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par l’attribut **RegExName**.</span><span class="sxs-lookup"><span data-stu-id="7209a-197">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="7209a-198">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="7209a-198">**Highlight**</span></span> | <span data-ttu-id="7209a-199">Non</span><span class="sxs-lookup"><span data-stu-id="7209a-199">No</span></span> | <span data-ttu-id="7209a-200">Spécifie comment le client doit mettre en surbrillance le texte correspondant.</span><span class="sxs-lookup"><span data-stu-id="7209a-200">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="7209a-201">Cet attribut ne peut être appliqué qu’aux éléments **Rule** au sein des éléments **ExtensionPoint**.</span><span class="sxs-lookup"><span data-stu-id="7209a-201">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="7209a-202">Peut correspondre à l’une des valeurs suivantes : `all` ou `none`.</span><span class="sxs-lookup"><span data-stu-id="7209a-202">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="7209a-203">Si non spécifié, la valeur par défaut est `all`.</span><span class="sxs-lookup"><span data-stu-id="7209a-203">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="7209a-204">**Remarque :** vous devez donner la valeur `BodyAsPlaintext` à l’attribut **PropertyName** si vous spécifiez l’attribut **Highlight** pour l’élément **Rule**.</span><span class="sxs-lookup"><span data-stu-id="7209a-204">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="7209a-205">Exemple</span><span class="sxs-lookup"><span data-stu-id="7209a-205">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="7209a-206">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="7209a-206">RuleCollection</span></span>

<span data-ttu-id="7209a-207">Définit une collection de règles et l’opérateur logique à utiliser lors de leur évaluation.</span><span class="sxs-lookup"><span data-stu-id="7209a-207">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="7209a-208">Attributs</span><span class="sxs-lookup"><span data-stu-id="7209a-208">Attributes</span></span>

| <span data-ttu-id="7209a-209">Attribut</span><span class="sxs-lookup"><span data-stu-id="7209a-209">Attribute</span></span> | <span data-ttu-id="7209a-210">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7209a-210">Required</span></span> | <span data-ttu-id="7209a-211">Description</span><span class="sxs-lookup"><span data-stu-id="7209a-211">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7209a-212">**Mode**</span><span class="sxs-lookup"><span data-stu-id="7209a-212">**Mode**</span></span> | <span data-ttu-id="7209a-213">Oui</span><span class="sxs-lookup"><span data-stu-id="7209a-213">Yes</span></span> | <span data-ttu-id="7209a-p109">Spécifie l’opérateur logique à utiliser lors de l’évaluation de cette collection de règles. Il peut s’agir des éléments `And` ou `Or`.</span><span class="sxs-lookup"><span data-stu-id="7209a-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="7209a-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="7209a-216">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="7209a-217">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7209a-217">See also</span></span>

- [<span data-ttu-id="7209a-218">Règles d’activation pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="7209a-218">Activation rules for Outlook add-ins</span></span>](../../outlook/activation-rules.md)
- [<span data-ttu-id="7209a-219">Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues</span><span class="sxs-lookup"><span data-stu-id="7209a-219">Match strings in an Outlook item as well-known entities</span></span>](../../outlook/match-strings-in-an-item-as-well-known-entities.md)    
- [<span data-ttu-id="7209a-220">Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="7209a-220">Use regular expression activation rules to show an Outlook add-in</span></span>](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
