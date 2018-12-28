---
title: Élément Rule dans le fichier manifeste
description: ''
ms.date: 11/30/2018
ms.openlocfilehash: 2c5ae07e5d0a3c9c8979abcada3d758c415e2e59
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457452"
---
# <a name="rule-element"></a><span data-ttu-id="61e41-102">Élément Rule</span><span class="sxs-lookup"><span data-stu-id="61e41-102">Rule element</span></span>

<span data-ttu-id="61e41-103">Spécifie les règles d’activation à évaluer pour ce complément de messagerie contextuel.</span><span class="sxs-lookup"><span data-stu-id="61e41-103">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="61e41-104">**Type de complément :** complément de messagerie contextuel</span><span class="sxs-lookup"><span data-stu-id="61e41-104">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="61e41-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="61e41-105">Contained in</span></span>

- [<span data-ttu-id="61e41-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="61e41-106">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="61e41-107">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="61e41-107">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="61e41-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="61e41-108">Attributes</span></span>

| <span data-ttu-id="61e41-109">Attribut</span><span class="sxs-lookup"><span data-stu-id="61e41-109">Attribute</span></span> | <span data-ttu-id="61e41-110">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="61e41-110">Required</span></span> | <span data-ttu-id="61e41-111">Description</span><span class="sxs-lookup"><span data-stu-id="61e41-111">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="61e41-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="61e41-112">**xsi:type**</span></span> | <span data-ttu-id="61e41-113">Oui</span><span class="sxs-lookup"><span data-stu-id="61e41-113">Yes</span></span> | <span data-ttu-id="61e41-114">Type de règle en cours de définition.</span><span class="sxs-lookup"><span data-stu-id="61e41-114">The type of rule being defined.</span></span> |

<span data-ttu-id="61e41-115">Le type de règle peut correspondre à l’une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="61e41-115">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="61e41-116">ItemIs</span><span class="sxs-lookup"><span data-stu-id="61e41-116">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="61e41-117">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="61e41-117">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="61e41-118">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="61e41-118">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="61e41-119">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="61e41-119">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="61e41-120">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="61e41-120">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="61e41-121">Règle ItemIs</span><span class="sxs-lookup"><span data-stu-id="61e41-121">ItemIs rule</span></span>

<span data-ttu-id="61e41-122">Définit une règle qui donne la valeur true si l’élément sélectionné est du type spécifié.</span><span class="sxs-lookup"><span data-stu-id="61e41-122">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="61e41-123">Attributs</span><span class="sxs-lookup"><span data-stu-id="61e41-123">Attributes</span></span>

| <span data-ttu-id="61e41-124">Attribut</span><span class="sxs-lookup"><span data-stu-id="61e41-124">Attribute</span></span> | <span data-ttu-id="61e41-125">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="61e41-125">Required</span></span> | <span data-ttu-id="61e41-126">Description</span><span class="sxs-lookup"><span data-stu-id="61e41-126">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="61e41-127">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="61e41-127">**ItemType**</span></span> | <span data-ttu-id="61e41-128">Oui</span><span class="sxs-lookup"><span data-stu-id="61e41-128">Yes</span></span> | <span data-ttu-id="61e41-p101">Spécifie le type d’élément avec lequel établir une correspondance. Peut être `Message` ou `Appointment`. Le type d’élément `Message` inclue les e-mails, les demandes de réunion, les réponses à une demande de réunion et les annulations de réunion.</span><span class="sxs-lookup"><span data-stu-id="61e41-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="61e41-132">**FormType**</span><span class="sxs-lookup"><span data-stu-id="61e41-132">**FormType**</span></span> | <span data-ttu-id="61e41-133">Non (dans [ExtensionPoint](extensionpoint.md)), Oui (dans [App_office](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="61e41-133">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="61e41-p102">Spécifie si l’application doit apparaître dans le formulaire de lecture ou de modification pour l’élément. Peut correspondre à l’une des valeurs suivantes : `Read`, `Edit`, `ReadOrEdit`. Si spécifiée dans un `Rule` dans un `ExtensionPoint`, cette valeur DOIT être `Read`.</span><span class="sxs-lookup"><span data-stu-id="61e41-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="61e41-137">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="61e41-137">**ItemClass**</span></span> | <span data-ttu-id="61e41-138">Non</span><span class="sxs-lookup"><span data-stu-id="61e41-138">No</span></span> | <span data-ttu-id="61e41-p103">Spécifie la classe de message personnalisé à mettre en correspondance. Pour plus d’informations, voir l’article relatif à l’[activation d’un complément de messagerie dans Outlook pour une classe de message spécifique](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="61e41-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="61e41-141">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="61e41-141">**IncludeSubClasses**</span></span> | <span data-ttu-id="61e41-142">Non</span><span class="sxs-lookup"><span data-stu-id="61e41-142">No</span></span> | <span data-ttu-id="61e41-143">Spécifie si la règle doit donner la valeur true si l’élément est une sous-classe de la classe de message spécifiée ; par défaut, la valeur est `false`.</span><span class="sxs-lookup"><span data-stu-id="61e41-143">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="61e41-144">Exemple</span><span class="sxs-lookup"><span data-stu-id="61e41-144">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="61e41-145">Règle ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="61e41-145">ItemHasAttachment rule</span></span>

<span data-ttu-id="61e41-146">Définit une règle qui donne la valeur true si l’élément contient une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="61e41-146">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="61e41-147">Exemple</span><span class="sxs-lookup"><span data-stu-id="61e41-147">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="61e41-148">Règle ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="61e41-148">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="61e41-149">Définit une règle qui donne la valeur true si l’élément contient dans son objet ou son corps du texte correspondant au type d’entité spécifié.</span><span class="sxs-lookup"><span data-stu-id="61e41-149">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="61e41-150">Attributs</span><span class="sxs-lookup"><span data-stu-id="61e41-150">Attributes</span></span>

| <span data-ttu-id="61e41-151">Attribut</span><span class="sxs-lookup"><span data-stu-id="61e41-151">Attribute</span></span> | <span data-ttu-id="61e41-152">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="61e41-152">Required</span></span> | <span data-ttu-id="61e41-153">Description</span><span class="sxs-lookup"><span data-stu-id="61e41-153">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="61e41-154">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="61e41-154">**EntityType**</span></span> | <span data-ttu-id="61e41-155">Oui</span><span class="sxs-lookup"><span data-stu-id="61e41-155">Yes</span></span> | <span data-ttu-id="61e41-p104">Spécifie le type d’entité à rechercher pour que la règle donne la valeur True. Il peut s’agir de l’un des éléments suivants : `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` ou `Contact`.</span><span class="sxs-lookup"><span data-stu-id="61e41-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="61e41-158">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="61e41-158">**RegExFilter**</span></span> | <span data-ttu-id="61e41-159">Non</span><span class="sxs-lookup"><span data-stu-id="61e41-159">No</span></span> | <span data-ttu-id="61e41-160">Spécifie une expression régulière à exécuter par rapport à cette entité à des fins d’activation.</span><span class="sxs-lookup"><span data-stu-id="61e41-160">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="61e41-161">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="61e41-161">**FilterName**</span></span> | <span data-ttu-id="61e41-162">Non</span><span class="sxs-lookup"><span data-stu-id="61e41-162">No</span></span> | <span data-ttu-id="61e41-163">Spécifie le nom du filtre d’expression régulière, afin qu’il soit possible par la suite de s’y référer dans le code de votre complément.</span><span class="sxs-lookup"><span data-stu-id="61e41-163">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="61e41-164">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="61e41-164">**IgnoreCase**</span></span> | <span data-ttu-id="61e41-165">Non</span><span class="sxs-lookup"><span data-stu-id="61e41-165">No</span></span> | <span data-ttu-id="61e41-166">Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par l’attribut **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="61e41-166">Specifies whether to ignore case when matching the regular expression specified by  **RegExFilter**.</span></span> |
| <span data-ttu-id="61e41-167">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="61e41-167">**Highlight**</span></span> | <span data-ttu-id="61e41-168">Non</span><span class="sxs-lookup"><span data-stu-id="61e41-168">No</span></span> | <span data-ttu-id="61e41-p105">**Remarque :** cela s’applique uniquement aux éléments **Rule** au sein des éléments **ExtensionPoint**. Spécifie comment le client doit mettre en surbrillance les entités correspondantes. Peut correspondre à l’une des valeurs suivantes : `all` ou `none`. Si non spécifié, la valeur par défaut est `all`.</span><span class="sxs-lookup"><span data-stu-id="61e41-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="61e41-173">Exemple</span><span class="sxs-lookup"><span data-stu-id="61e41-173">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="61e41-174">Règle ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="61e41-174">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="61e41-175">Définit une règle qui donne la valeur true si une correspondance de l’expression régulière spécifiée est trouvée dans la propriété spécifiée de l’élément.</span><span class="sxs-lookup"><span data-stu-id="61e41-175">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="61e41-176">Attributs</span><span class="sxs-lookup"><span data-stu-id="61e41-176">Attributes</span></span>

| <span data-ttu-id="61e41-177">Attribut</span><span class="sxs-lookup"><span data-stu-id="61e41-177">Attribute</span></span> | <span data-ttu-id="61e41-178">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="61e41-178">Required</span></span> | <span data-ttu-id="61e41-179">Description</span><span class="sxs-lookup"><span data-stu-id="61e41-179">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="61e41-180">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="61e41-180">**RegExName**</span></span> | <span data-ttu-id="61e41-181">Oui</span><span class="sxs-lookup"><span data-stu-id="61e41-181">Yes</span></span> | <span data-ttu-id="61e41-182">Spécifie le nom de l’expression régulière afin que vous puissiez vous référer à l’expression dans le code de votre complément.</span><span class="sxs-lookup"><span data-stu-id="61e41-182">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="61e41-183">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="61e41-183">**RegExValue**</span></span> | <span data-ttu-id="61e41-184">Oui</span><span class="sxs-lookup"><span data-stu-id="61e41-184">Yes</span></span> | <span data-ttu-id="61e41-185">Spécifie l’expression régulière qui sera évaluée pour déterminer si le complément de messagerie doit être affiché.</span><span class="sxs-lookup"><span data-stu-id="61e41-185">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="61e41-186">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="61e41-186">**PropertyName**</span></span> | <span data-ttu-id="61e41-187">Oui</span><span class="sxs-lookup"><span data-stu-id="61e41-187">Yes</span></span> | <span data-ttu-id="61e41-p106">Spécifie le nom de la propriété par rapport à laquelle l’expression sera évaluée. Les options disponibles sont les suivantes : `Subject`, `BodyAsPlaintext`, `BodyAsHTML` ou `SenderSMTPAddress`.</span><span class="sxs-lookup"><span data-stu-id="61e41-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="61e41-190">Si vous spécifiez `BodyAsHTML`, Outlook applique seulement l’expression régulière si le corps de l’élément est du code HTML.</span><span class="sxs-lookup"><span data-stu-id="61e41-190">If you specify `BodyAsHTML`, Outlook applies the regular expression only if the item body is HTML, and otherwise Outlook returns no matches for that regular expression.</span></span> <span data-ttu-id="61e41-191">Si ce n’est pas le cas, Outlook ne renvoie aucune correspondance pour cette expression régulière.</span><span class="sxs-lookup"><span data-stu-id="61e41-191">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="61e41-192">Si vous spécifiez `BodyAsPlaintext`, Outlook applique toujours l’expression régulière au corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="61e41-192">If you specify  `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="61e41-193">**Remarque :** vous devez donner la valeur `BodyAsPlaintext` à l’attribut **PropertyName** si vous spécifiez l’attribut **Highlight** pour l’élément **Rule**.</span><span class="sxs-lookup"><span data-stu-id="61e41-193">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="61e41-194">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="61e41-194">**IgnoreCase**</span></span> | <span data-ttu-id="61e41-195">Non</span><span class="sxs-lookup"><span data-stu-id="61e41-195">No</span></span> | <span data-ttu-id="61e41-196">Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par l’attribut **RegExName**.</span><span class="sxs-lookup"><span data-stu-id="61e41-196">Specifies whether to ignore case when matching the regular expression specified by  **RegExName**.</span></span> |
| <span data-ttu-id="61e41-197">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="61e41-197">**Highlight**</span></span> | <span data-ttu-id="61e41-198">Non</span><span class="sxs-lookup"><span data-stu-id="61e41-198">No</span></span> | <span data-ttu-id="61e41-199">Spécifie comment le client doit mettre en surbrillance le texte correspondant.</span><span class="sxs-lookup"><span data-stu-id="61e41-199">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="61e41-200">Cet attribut ne peut être appliqué qu’aux éléments **Rule** au sein des éléments **ExtensionPoint**.</span><span class="sxs-lookup"><span data-stu-id="61e41-200">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="61e41-201">Peut correspondre à l’une des valeurs suivantes : `all` ou `none`.</span><span class="sxs-lookup"><span data-stu-id="61e41-201">The returned string can be one of the following values: `all`, , or `none`.</span></span> <span data-ttu-id="61e41-202">Si non spécifié, la valeur par défaut est `all`.</span><span class="sxs-lookup"><span data-stu-id="61e41-202">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="61e41-203">**Remarque :** vous devez donner la valeur `BodyAsPlaintext` à l’attribut **PropertyName** si vous spécifiez l’attribut **Highlight** pour l’élément **Rule**.</span><span class="sxs-lookup"><span data-stu-id="61e41-203">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="61e41-204">Exemple</span><span class="sxs-lookup"><span data-stu-id="61e41-204">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="61e41-205">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="61e41-205">RuleCollection</span></span>

<span data-ttu-id="61e41-206">Définit une collection de règles et l’opérateur logique à utiliser lors de leur évaluation.</span><span class="sxs-lookup"><span data-stu-id="61e41-206">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="61e41-207">Attributs</span><span class="sxs-lookup"><span data-stu-id="61e41-207">Attributes</span></span>

| <span data-ttu-id="61e41-208">Attribut</span><span class="sxs-lookup"><span data-stu-id="61e41-208">Attribute</span></span> | <span data-ttu-id="61e41-209">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="61e41-209">Required</span></span> | <span data-ttu-id="61e41-210">Description</span><span class="sxs-lookup"><span data-stu-id="61e41-210">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="61e41-211">**Mode**</span><span class="sxs-lookup"><span data-stu-id="61e41-211">**Mode**</span></span> | <span data-ttu-id="61e41-212">Oui</span><span class="sxs-lookup"><span data-stu-id="61e41-212">Yes</span></span> | <span data-ttu-id="61e41-p109">Spécifie l’opérateur logique à utiliser lors de l’évaluation de cette collection de règles. Il peut s’agir des éléments `And` ou `Or`.</span><span class="sxs-lookup"><span data-stu-id="61e41-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="61e41-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="61e41-215">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="61e41-216">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="61e41-216">See also</span></span>

- [<span data-ttu-id="61e41-217">Règles d’activation pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="61e41-217">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="61e41-218">Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues</span><span class="sxs-lookup"><span data-stu-id="61e41-218">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="61e41-219">Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="61e41-219">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)