# <a name="rule-element"></a><span data-ttu-id="47a98-101">Élément Rule</span><span class="sxs-lookup"><span data-stu-id="47a98-101">Rule element</span></span>

<span data-ttu-id="47a98-102">Spécifie les règles d’activation à évaluer pour ce complément de messagerie contextuel.</span><span class="sxs-lookup"><span data-stu-id="47a98-102">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="47a98-103">**Type de complément :** complément de messagerie contextuel</span><span class="sxs-lookup"><span data-stu-id="47a98-103">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="47a98-104">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="47a98-104">Contained in:</span></span>

- [<span data-ttu-id="47a98-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="47a98-105">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="47a98-106">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="47a98-106">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="47a98-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="47a98-107">Attributes</span></span>

| <span data-ttu-id="47a98-108">Attribut</span><span class="sxs-lookup"><span data-stu-id="47a98-108">Attribute</span></span> | <span data-ttu-id="47a98-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="47a98-109">Required</span></span> | <span data-ttu-id="47a98-110">Description</span><span class="sxs-lookup"><span data-stu-id="47a98-110">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="47a98-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="47a98-111">**xsi:type**</span></span> | <span data-ttu-id="47a98-112">Oui</span><span class="sxs-lookup"><span data-stu-id="47a98-112">Yes</span></span> | <span data-ttu-id="47a98-113">Type de règle en cours de définition.</span><span class="sxs-lookup"><span data-stu-id="47a98-113">The type of rule being defined.</span></span> |

<span data-ttu-id="47a98-114">Le type de règle peut correspondre à l’une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="47a98-114">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="47a98-115">ItemIs</span><span class="sxs-lookup"><span data-stu-id="47a98-115">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="47a98-116">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="47a98-116">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="47a98-117">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="47a98-117">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="47a98-118">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="47a98-118">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="47a98-119">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="47a98-119">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="47a98-120">Règle ItemIs</span><span class="sxs-lookup"><span data-stu-id="47a98-120">ItemIs rule</span></span>

<span data-ttu-id="47a98-121">Définit une règle qui donne la valeur true si l’élément sélectionné est du type spécifié.</span><span class="sxs-lookup"><span data-stu-id="47a98-121">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="47a98-122">Attributs</span><span class="sxs-lookup"><span data-stu-id="47a98-122">Attributes</span></span>

| <span data-ttu-id="47a98-123">Attribut</span><span class="sxs-lookup"><span data-stu-id="47a98-123">Attribute</span></span> | <span data-ttu-id="47a98-124">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="47a98-124">Required</span></span> | <span data-ttu-id="47a98-125">Description</span><span class="sxs-lookup"><span data-stu-id="47a98-125">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="47a98-126">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="47a98-126">**ItemType**</span></span> | <span data-ttu-id="47a98-127">Oui</span><span class="sxs-lookup"><span data-stu-id="47a98-127">Yes</span></span> | <span data-ttu-id="47a98-p101">Spécifie le type d’élément à mettre en correspondance. Peut être `Message` ou `Appointment`. Le type d’élément `Message` inclut e-mails, demandes de réunion, réponses à une demande de réunion et annulations de réunion.</span><span class="sxs-lookup"><span data-stu-id="47a98-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="47a98-131">**FormType**</span><span class="sxs-lookup"><span data-stu-id="47a98-131">**FormType**</span></span> | <span data-ttu-id="47a98-132">Non (dans [ExtensionPoint](extensionpoint.md)), Oui (dans [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="47a98-132">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="47a98-p102">Spécifie si l’application doit apparaître dans le formulaire de lecture ou de modification pour l’élément. Peut correspondre à l’une des valeurs suivantes : `Read`, `Edit`, `ReadOrEdit`. Si spécifiée dans un `Rule` dans un `ExtensionPoint`, cette valeur DOIT être `Read`.</span><span class="sxs-lookup"><span data-stu-id="47a98-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="47a98-136">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="47a98-136">**ItemClass**</span></span> | <span data-ttu-id="47a98-137">Non</span><span class="sxs-lookup"><span data-stu-id="47a98-137">No</span></span> | <span data-ttu-id="47a98-p103">Spécifie la classe de message personnalisé à mettre en correspondance. Pour plus d’informations, voir l’article relatif à l’[Activation d’un complément de messagerie dans Outlook pour une classe de message spécifique](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="47a98-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="47a98-140">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="47a98-140">**IncludeSubClasses**</span></span> | <span data-ttu-id="47a98-141">Non</span><span class="sxs-lookup"><span data-stu-id="47a98-141">No</span></span> | <span data-ttu-id="47a98-142">Spécifie si la règle doit donner la valeur true si l’élément est une sous-classe de la classe de message spécifiée. Par défaut, la valeur est `false`.</span><span class="sxs-lookup"><span data-stu-id="47a98-142">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="47a98-143">Exemple</span><span class="sxs-lookup"><span data-stu-id="47a98-143">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="47a98-144">Règle ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="47a98-144">ItemHasAttachment rule</span></span>

<span data-ttu-id="47a98-145">Définit une règle qui donne la valeur true si l’élément contient une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="47a98-145">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="47a98-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="47a98-146">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="47a98-147">Règle ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="47a98-147">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="47a98-148">Définit une règle qui donne la valeur true si l’élément contient dans son objet ou son corps du texte correspondant au type d’entité spécifié.</span><span class="sxs-lookup"><span data-stu-id="47a98-148">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="47a98-149">Attributs</span><span class="sxs-lookup"><span data-stu-id="47a98-149">Attributes</span></span>

| <span data-ttu-id="47a98-150">Attribut</span><span class="sxs-lookup"><span data-stu-id="47a98-150">Attribute</span></span> | <span data-ttu-id="47a98-151">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="47a98-151">Required</span></span> | <span data-ttu-id="47a98-152">Description</span><span class="sxs-lookup"><span data-stu-id="47a98-152">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="47a98-153">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="47a98-153">**EntityType**</span></span> | <span data-ttu-id="47a98-154">Oui</span><span class="sxs-lookup"><span data-stu-id="47a98-154">Yes</span></span> | <span data-ttu-id="47a98-p104">Spécifie le type d’entité à rechercher pour que la règle donne la valeur true. Peut correspondre à l’une des valeurs suivantes : `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` ou `Contact`.</span><span class="sxs-lookup"><span data-stu-id="47a98-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="47a98-157">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="47a98-157">**RegExFilter**</span></span> | <span data-ttu-id="47a98-158">Non</span><span class="sxs-lookup"><span data-stu-id="47a98-158">No</span></span> | <span data-ttu-id="47a98-159">Spécifie une expression régulière à exécuter par rapport à cette entité à des fins d’activation.</span><span class="sxs-lookup"><span data-stu-id="47a98-159">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="47a98-160">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="47a98-160">**FilterName**</span></span> | <span data-ttu-id="47a98-161">Non</span><span class="sxs-lookup"><span data-stu-id="47a98-161">No</span></span> | <span data-ttu-id="47a98-162">Spécifie le nom du filtre d’expression régulière, afin qu’il soit possible par la suite de s’y référer dans le code de votre complément.</span><span class="sxs-lookup"><span data-stu-id="47a98-162">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="47a98-163">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="47a98-163">**IgnoreCase**</span></span> | <span data-ttu-id="47a98-164">Non</span><span class="sxs-lookup"><span data-stu-id="47a98-164">No</span></span> | <span data-ttu-id="47a98-165">Indique d’ignorer la casse lors de l’exécution de l’expression régulière spécifiée par l’attribut **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="47a98-165">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="47a98-166">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="47a98-166">**Highlight**</span></span> | <span data-ttu-id="47a98-167">Non</span><span class="sxs-lookup"><span data-stu-id="47a98-167">No</span></span> | <span data-ttu-id="47a98-p105">**Remarque :** cela s’applique uniquement aux éléments **Rule** au sein des éléments **ExtensionPoint**. Spécifie comment le client doit mettre en surbrillance les entités correspondantes. Peut correspondre à l’une des valeurs suivantes : `all` ou `none`. Si non spécifié, la valeur par défaut est `all`.</span><span class="sxs-lookup"><span data-stu-id="47a98-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="47a98-172">Exemple</span><span class="sxs-lookup"><span data-stu-id="47a98-172">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="47a98-173">Règle ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="47a98-173">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="47a98-174">Définit une règle qui donne la valeur true si une correspondance de l’expression régulière spécifiée est trouvée dans la propriété spécifiée de l’élément.</span><span class="sxs-lookup"><span data-stu-id="47a98-174">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="47a98-175">Attributs</span><span class="sxs-lookup"><span data-stu-id="47a98-175">Attributes</span></span>

| <span data-ttu-id="47a98-176">Attribut</span><span class="sxs-lookup"><span data-stu-id="47a98-176">Attribute</span></span> | <span data-ttu-id="47a98-177">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="47a98-177">Required</span></span> | <span data-ttu-id="47a98-178">Description</span><span class="sxs-lookup"><span data-stu-id="47a98-178">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="47a98-179">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="47a98-179">**RegExName**</span></span> | <span data-ttu-id="47a98-180">Oui</span><span class="sxs-lookup"><span data-stu-id="47a98-180">Yes</span></span> | <span data-ttu-id="47a98-181">Spécifie le nom de l’expression régulière afin que vous puissiez vous référer à l’expression dans le code de votre complément.</span><span class="sxs-lookup"><span data-stu-id="47a98-181">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="47a98-182">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="47a98-182">**RegExValue**</span></span> | <span data-ttu-id="47a98-183">Oui</span><span class="sxs-lookup"><span data-stu-id="47a98-183">Yes</span></span> | <span data-ttu-id="47a98-184">Spécifie l’expression régulière qui sera évaluée pour déterminer si le complément de messagerie doit être affiché.</span><span class="sxs-lookup"><span data-stu-id="47a98-184">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="47a98-185">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="47a98-185">**PropertyName**</span></span> | <span data-ttu-id="47a98-186">Oui</span><span class="sxs-lookup"><span data-stu-id="47a98-186">Yes</span></span> | <span data-ttu-id="47a98-p106">Spécifie le nom de la propriété par rapport à laquelle l’expression sera évaluée. Peut correspondre à l’une des valeurs suivantes : `Subject`, `BodyAsPlaintext`, `BodyAsHtml` ou `SenderSTMPAddress`.</span><span class="sxs-lookup"><span data-stu-id="47a98-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHtml`, or `SenderSTMPAddress`.</span></span> |
| <span data-ttu-id="47a98-189">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="47a98-189">**IgnoreCase**</span></span> | <span data-ttu-id="47a98-190">Non</span><span class="sxs-lookup"><span data-stu-id="47a98-190">No</span></span> | <span data-ttu-id="47a98-191">Indique d’ignorer la casse lors de l’exécution de l’expression régulière.</span><span class="sxs-lookup"><span data-stu-id="47a98-191">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="47a98-192">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="47a98-192">**Highlight**</span></span> | <span data-ttu-id="47a98-193">Non</span><span class="sxs-lookup"><span data-stu-id="47a98-193">No</span></span> | <span data-ttu-id="47a98-p107">**Remarque :** cela s’applique uniquement aux éléments **Rule** au sein des éléments **ExtensionPoint**. Spécifie comment le client doit mettre en surbrillance le texte correspondant. Peut correspondre à l’une des valeurs suivantes : `all` ou `none`. Si non spécifié, la valeur par défaut est `all`.</span><span class="sxs-lookup"><span data-stu-id="47a98-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="47a98-198">Exemple</span><span class="sxs-lookup"><span data-stu-id="47a98-198">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="47a98-199">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="47a98-199">RuleCollection</span></span>

<span data-ttu-id="47a98-200">Définit une collection de règles et l’opérateur logique à utiliser lors de leur évaluation.</span><span class="sxs-lookup"><span data-stu-id="47a98-200">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="47a98-201">Attributs</span><span class="sxs-lookup"><span data-stu-id="47a98-201">Attributes</span></span>

| <span data-ttu-id="47a98-202">Attribut</span><span class="sxs-lookup"><span data-stu-id="47a98-202">Attribute</span></span> | <span data-ttu-id="47a98-203">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="47a98-203">Required</span></span> | <span data-ttu-id="47a98-204">Description</span><span class="sxs-lookup"><span data-stu-id="47a98-204">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="47a98-205">**Mode**</span><span class="sxs-lookup"><span data-stu-id="47a98-205">**Mode**</span></span> | <span data-ttu-id="47a98-206">Oui</span><span class="sxs-lookup"><span data-stu-id="47a98-206">Yes</span></span> | <span data-ttu-id="47a98-p108">Spécifie l’opérateur logique à utiliser lors de l’évaluation de cette collection de règles. Peut être `And` ou `Or`.</span><span class="sxs-lookup"><span data-stu-id="47a98-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="47a98-209">Exemple</span><span class="sxs-lookup"><span data-stu-id="47a98-209">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="47a98-210">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="47a98-210">See also</span></span>

- [<span data-ttu-id="47a98-211">Règles d’activation pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="47a98-211">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="47a98-212">Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues</span><span class="sxs-lookup"><span data-stu-id="47a98-212">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="47a98-213">Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="47a98-213">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)