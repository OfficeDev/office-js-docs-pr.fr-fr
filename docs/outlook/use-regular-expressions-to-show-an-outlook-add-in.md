---
title: Utiliser les règles d’activation d’expression régulière afin d’afficher un complément
description: Découvrez comment utiliser les règles d’activation d’expression régulière pour les compléments contextuels Outlook.
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: 4a5507b410ed729f76c3efa0119e87c6a6dbc71a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292474"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a><span data-ttu-id="992c2-103">Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="992c2-103">Use regular expression activation rules to show an Outlook add-in</span></span>

<span data-ttu-id="992c2-104">Vous pouvez spécifier des règles d’expressions régulières pour qu’un [complément contextuel](contextual-outlook-add-ins.md) soit activé lorsqu’une correspondance est trouvée dans les champs spécifiques du message.</span><span class="sxs-lookup"><span data-stu-id="992c2-104">You can specify regular expression rules to have a [contextual add-in](contextual-outlook-add-ins.md) activated when a match is found in specific fields of the message.</span></span> <span data-ttu-id="992c2-105">Les compléments contextuels sont activés uniquement en mode lecture. Outlook n’active pas de compléments contextuels lorsque l’utilisateur compose un élément.</span><span class="sxs-lookup"><span data-stu-id="992c2-105">Contextual add-ins activate only in read mode, Outlook does not activate contextual add-ins when the user is composing an item.</span></span> <span data-ttu-id="992c2-106">Il existe également d’autres scénarios dans lesquels Outlook n’active pas de compléments, par exemple, des éléments signés numériquement.</span><span class="sxs-lookup"><span data-stu-id="992c2-106">There are also other scenarios where Outlook does not activate add-ins, for example, digitally signed items.</span></span> <span data-ttu-id="992c2-107">Pour plus d’informations, reportez-vous à la rubrique [Règles d’activation pour les compléments Outlook](activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="992c2-107">For more information, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

<span data-ttu-id="992c2-108">Vous pouvez spécifier une expression régulière dans le cadre d’une règle [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) ou [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) dans le manifeste XML du complément.</span><span class="sxs-lookup"><span data-stu-id="992c2-108">You can specify a regular expression as part of an [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule or [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule in the add-in XML manifest.</span></span> <span data-ttu-id="992c2-109">Les règles sont spécifiées dans un point d’extension [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity).</span><span class="sxs-lookup"><span data-stu-id="992c2-109">The rules are specified in a [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity) extension point.</span></span>

<span data-ttu-id="992c2-110">Outlook évalue les expressions régulières en fonction des règles définies pour l’interpréteur JavaScript utilisé par le navigateur de l’ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="992c2-110">Outlook evaluates regular expressions based on the rules for the JavaScript interpreter used by the browser on the client computer.</span></span> <span data-ttu-id="992c2-111">Outlook prend en charge la même liste de caractères spéciaux que tous les processeurs XML.</span><span class="sxs-lookup"><span data-stu-id="992c2-111">Outlook supports the same list of special characters that all XML processors also support.</span></span> <span data-ttu-id="992c2-112">Le tableau suivant répertorie ces caractères spéciaux.</span><span class="sxs-lookup"><span data-stu-id="992c2-112">The following table lists these special characters.</span></span> <span data-ttu-id="992c2-113">Vous pouvez les utiliser dans une expression régulière en spécifiant la séquence d’échappement pour le caractère correspondant, comme décrit dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="992c2-113">You can use these characters in a regular expression by specifying the escaped sequence for the corresponding character, as described in the following table.</span></span>

<br/>

|<span data-ttu-id="992c2-114">Caractère</span><span class="sxs-lookup"><span data-stu-id="992c2-114">Character</span></span>|<span data-ttu-id="992c2-115">Description</span><span class="sxs-lookup"><span data-stu-id="992c2-115">Description</span></span>|<span data-ttu-id="992c2-116">Séquence d’échappement à utiliser</span><span class="sxs-lookup"><span data-stu-id="992c2-116">Escape sequence to use</span></span>|
|:-----|:-----|:-----|
|`"`|<span data-ttu-id="992c2-117">Guillemets doubles</span><span class="sxs-lookup"><span data-stu-id="992c2-117">Double quotation mark</span></span>|`&quot;`|
|`&`|<span data-ttu-id="992c2-118">Esperluette</span><span class="sxs-lookup"><span data-stu-id="992c2-118">Ampersand</span></span>|`&amp;`|
|`'`|<span data-ttu-id="992c2-119">Apostrophe</span><span class="sxs-lookup"><span data-stu-id="992c2-119">Apostrophe</span></span>|`&apos;`|
|`<`|<span data-ttu-id="992c2-120">Signe inférieur à</span><span class="sxs-lookup"><span data-stu-id="992c2-120">Less-than sign</span></span>|`&lt;`|
|`>`|<span data-ttu-id="992c2-121">Signe supérieur à</span><span class="sxs-lookup"><span data-stu-id="992c2-121">Greater-than sign</span></span>|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="992c2-122">Règle ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="992c2-122">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="992c2-123">Une règle `ItemHasRegularExpressionMatch` est utile dans le contrôle de l’activation d’un complément basé sur les valeurs spécifiques d’une propriété prise en charge.</span><span class="sxs-lookup"><span data-stu-id="992c2-123">An  `ItemHasRegularExpressionMatch` rule is useful in controlling activation of an add-in based on specific values of a supported property.</span></span> <span data-ttu-id="992c2-124">La règle `ItemHasRegularExpressionMatch` contient les attributs suivants.</span><span class="sxs-lookup"><span data-stu-id="992c2-124">The `ItemHasRegularExpressionMatch` rule has the following attributes.</span></span>

<br/>

|<span data-ttu-id="992c2-125">Nom de l’attribut</span><span class="sxs-lookup"><span data-stu-id="992c2-125">Attribute name</span></span>|<span data-ttu-id="992c2-126">Description</span><span class="sxs-lookup"><span data-stu-id="992c2-126">Description</span></span>|
|:-----|:-----|
|`RegExName`|<span data-ttu-id="992c2-127">Spécifie le nom de l’expression régulière afin que vous puissiez vous référer à l’expression dans le code de votre complément.</span><span class="sxs-lookup"><span data-stu-id="992c2-127">Specifies the name of the regular expression so that you can refer to the expression in the code for your add-in.</span></span>|
|`RegExValue`|<span data-ttu-id="992c2-128">Spécifie l’expression régulière qui sera évaluée pour déterminer si le complément doit être affiché.</span><span class="sxs-lookup"><span data-stu-id="992c2-128">Specifies the regular expression that will be evaluated to determine whether the add-in should be shown.</span></span>|
|`PropertyName`|<span data-ttu-id="992c2-129">Spécifie le nom de la propriété par rapport à laquelle l’expression sera évaluée.</span><span class="sxs-lookup"><span data-stu-id="992c2-129">Specifies the name of the property that the regular expression will be evaluated against.</span></span> <span data-ttu-id="992c2-130">Les valeurs autorisées sont `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress` et `Subject`.</span><span class="sxs-lookup"><span data-stu-id="992c2-130">The allowed values are `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress`, and `Subject`.</span></span><br/><br/><span data-ttu-id="992c2-131">Si vous spécifiez `BodyAsHTML`, Outlook applique seulement l’expression régulière si le corps de l’élément est du code HTML.</span><span class="sxs-lookup"><span data-stu-id="992c2-131">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="992c2-132">Si ce n’est pas le cas, Outlook ne renvoie aucune correspondance pour cette expression régulière.</span><span class="sxs-lookup"><span data-stu-id="992c2-132">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="992c2-133">Si vous spécifiez `BodyAsPlaintext`, Outlook applique toujours l’expression régulière au corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="992c2-133">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="992c2-134">**Remarque :** vous devez définir l’attribut `PropertyName` sur `BodyAsPlaintext` si vous spécifiez l’attribut `Highlight` pour l’élément `Rule`.</span><span class="sxs-lookup"><span data-stu-id="992c2-134">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span>|
|`IgnoreCase`|<span data-ttu-id="992c2-135">Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="992c2-135">Specifies whether to ignore case when matching the regular expression specified by `RegExName`.</span></span>|
| `Highlight` | <span data-ttu-id="992c2-136">Spécifie la façon dont le client doit mettre en évidence le texte correspondant.</span><span class="sxs-lookup"><span data-stu-id="992c2-136">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="992c2-137">Cet élément peut uniquement s’appliquer à des éléments `Rule` au sein d’éléments `ExtensionPoint`.</span><span class="sxs-lookup"><span data-stu-id="992c2-137">This element can only be applied to `Rule` elements within `ExtensionPoint` elements.</span></span> <span data-ttu-id="992c2-138">Peut correspondre à l’une des valeurs suivantes : `all` ou `none`.</span><span class="sxs-lookup"><span data-stu-id="992c2-138">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="992c2-139">Si non spécifié, la valeur par défaut est `all`.</span><span class="sxs-lookup"><span data-stu-id="992c2-139">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="992c2-140">**Remarque :** vous devez définir l’attribut `PropertyName` sur `BodyAsPlaintext` si vous spécifiez l’attribut `Highlight` pour l’élément `Rule`.</span><span class="sxs-lookup"><span data-stu-id="992c2-140">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span> |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a><span data-ttu-id="992c2-141">Meilleures pratiques pour l’utilisation d’expressions régulières dans les règles</span><span class="sxs-lookup"><span data-stu-id="992c2-141">Best practices for using regular expressions in rules</span></span>

<span data-ttu-id="992c2-142">Prêtez une attention particulière aux éléments suivants lorsque vous utilisez des expressions régulières :</span><span class="sxs-lookup"><span data-stu-id="992c2-142">Pay special attention to the following when you use regular expressions:</span></span>

- <span data-ttu-id="992c2-143">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour le corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="992c2-143">If you specify an `ItemHasRegularExpressionMatch` rule on the body of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item.</span></span> <span data-ttu-id="992c2-144">L’utilisation d’une expression régulière telle que `.*` pour essayer d’obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="992c2-144">Using a regular expression such as `.*` to attempt to obtain the entire body of an item does not always return the expected results.</span></span>
- <span data-ttu-id="992c2-145">Le corps en texte brut renvoyé sur un navigateur peut être légèrement différent sur un autre.</span><span class="sxs-lookup"><span data-stu-id="992c2-145">The plain text body returned on one browser can be different in subtle ways on another.</span></span> <span data-ttu-id="992c2-146">Si vous utilisez une règle `ItemHasRegularExpressionMatch` avec `BodyAsPlaintext` comme attribut `PropertyName`, testez votre expression régulière sur tous les navigateurs pris en charge par votre complément.</span><span class="sxs-lookup"><span data-stu-id="992c2-146">If you use an `ItemHasRegularExpressionMatch` rule with `BodyAsPlaintext` as the `PropertyName` attribute, test your regular expression on all the browsers that your add-in supports.</span></span>

    <span data-ttu-id="992c2-147">Comme différents navigateurs utilisent diverses méthodes pour obtenir le corps du texte d’un élément sélectionné, vous devez vous assurer que votre expression régulière prend en charge les fines différences qui peuvent être renvoyées dans le cadre du corps de texte.</span><span class="sxs-lookup"><span data-stu-id="992c2-147">Because different browsers use different ways to obtain the text body of a selected item, you should make sure that your regular expression supports the subtle differences that can be returned as part of the body text.</span></span> <span data-ttu-id="992c2-148">Par exemple, certains navigateurs, comme Internet Explorer 9, utilisent la propriété `innerText` du DOM, tandis que d’autres, comme Firefox, utilisent la méthode `.textContent()` afin d’obtenir le corps du texte d’un élément.</span><span class="sxs-lookup"><span data-stu-id="992c2-148">For example, some browsers such as Internet Explorer 9 uses the `innerText` property of the DOM, and others such as Firefox uses the `.textContent()` method to obtain the text body of an item.</span></span> <span data-ttu-id="992c2-149">En outre, différents navigateurs peuvent renvoyer des sauts de ligne de manière différente : un saut de ligne correspond à `\r\n` sur Internet Explorer, et `\n` dans Firefox et Chrome.</span><span class="sxs-lookup"><span data-stu-id="992c2-149">Also, different browsers may return line breaks differently: a line break is `\r\n` on Internet Explorer, and `\n` on Firefox and Chrome.</span></span> <span data-ttu-id="992c2-150">Pour plus d’informations, consultez la page sur la [compatibilité DOM W3C - HTML](https://quirksmode.org/dom/html/).</span><span class="sxs-lookup"><span data-stu-id="992c2-150">For more information, se [W3C DOM Compatibility - HTML](https://quirksmode.org/dom/html/).</span></span>

- <span data-ttu-id="992c2-151">Le corps HTML d’un élément est légèrement différent entre un client riche Outlook et Outlook sur le web ou Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="992c2-151">The HTML body of an item is slightly different between an Outlook rich client, and Outlook on the web or Outlook mobile.</span></span> <span data-ttu-id="992c2-152">Définissez attentivement vos expressions régulières.</span><span class="sxs-lookup"><span data-stu-id="992c2-152">Define your regular expressions carefully.</span></span>

- <span data-ttu-id="992c2-p112">Selon le client Outlook, le type de périphérique ou la propriété auquel une expression régulière est appliquée, il existe d’autres meilleures pratiques et limites pour chacun des clients que vous devez connaître lors de la conception d’expressions régulières en tant que règles d’activation. Pour plus d’informations, consultez la rubrique [limites pour l’activation et l’API JavaScript pour les compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) .</span><span class="sxs-lookup"><span data-stu-id="992c2-p112">Depending on the Outlook client, type of device, or property that a regular expression is being applied on, there are other best practices and limits for each of the clients that you should be aware of when designing regular expressions as activation rules. See [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) for details.</span></span>

### <a name="examples"></a><span data-ttu-id="992c2-155">Exemples</span><span class="sxs-lookup"><span data-stu-id="992c2-155">Examples</span></span>

<span data-ttu-id="992c2-156">La règle `ItemHasRegularExpressionMatch` suivante active le complément chaque fois que l’adresse de messagerie SMTP de l’expéditeur correspond à `@contoso`, indépendamment des caractères majuscules et minuscules.</span><span class="sxs-lookup"><span data-stu-id="992c2-156">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever the sender's SMTP email address matches `@contoso`, regardless of uppercase or lowercase characters.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

<span data-ttu-id="992c2-157">L’exemple suivant montre une autre manière de spécifier la même expression régulière à l’aide de l’attribut `IgnoreCase`.</span><span class="sxs-lookup"><span data-stu-id="992c2-157">The following is another way to specify the same regular expression using the  `IgnoreCase` attribute.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

<span data-ttu-id="992c2-158">La règle `ItemHasRegularExpressionMatch` suivante active le complément chaque fois qu’un symbole de valeur est inclus dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="992c2-158">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever a stock symbol is included in the body of the current item.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="992c2-159">Règle ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="992c2-159">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="992c2-160">Une règle `ItemHasKnownEntity` active un complément en fonction de l'existence d'une entité dans le sujet ou le corps de l'élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="992c2-160">An `ItemHasKnownEntity` rule activates an add-in based on the existence of an entity in the subject or body of the selected item.</span></span> <span data-ttu-id="992c2-161">Le type [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) définit les entités prises en charge.</span><span class="sxs-lookup"><span data-stu-id="992c2-161">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) type defines the supported entities.</span></span> <span data-ttu-id="992c2-162">L’application d’une expression régulière sur une règle `ItemHasKnownEntity` convient lorsque l’activation est basée sur un sous-ensemble de valeurs pour une entité (par exemple, un ensemble spécifique d’URL, ou des numéros de téléphone avec un certain code régional).</span><span class="sxs-lookup"><span data-stu-id="992c2-162">Applying a regular expression on an `ItemHasKnownEntity` rule provides the convenience where activation is based on a subset of values for an entity (for example, a specific set of URLs, or telephone numbers with a certain area code).</span></span>

> [!NOTE]
> <span data-ttu-id="992c2-163">Outlook peut extraire uniquement des chaînes d’entité en anglais, indépendamment des paramètres régionaux par défaut spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="992c2-163">Outlook can only extract entity strings in English regardless of the default locale specified in the manifest.</span></span> <span data-ttu-id="992c2-164">Seuls les messages prennent en charge le type d’entité `MeetingSuggestion`. Ce n’est pas le cas des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="992c2-164">Only messages support the `MeetingSuggestion` entity type; appointments do not.</span></span> <span data-ttu-id="992c2-165">Vous ne pouvez pas extraire les entités des éléments figurant dans le dossier **Éléments envoyés**, ni utiliser une règle `ItemHasKnownEntity` afin d’activer un complément pour les éléments du dossier **Éléments envoyés**.</span><span class="sxs-lookup"><span data-stu-id="992c2-165">You cannot extract entities from items in the **Sent Items** folder, nor can you use an `ItemHasKnownEntity` rule to activate an add-in for items in the **Sent Items** folder.</span></span>

<span data-ttu-id="992c2-166">La règle `ItemHasKnownEntity` prend en charge les attributs dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="992c2-166">The `ItemHasKnownEntity` rule supports the attributes in the following table.</span></span> <span data-ttu-id="992c2-167">Notez que, bien que la spécification d’une expression régulière soit facultative dans une règle `ItemHasKnownEntity`, si vous choisissez d’utiliser une expression régulière comme filtre d’entité, vous devez spécifier à la fois l’attribut `RegExFilter` et `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="992c2-167">Note that while specifying a regular expression is optional in an `ItemHasKnownEntity` rule, if you choose to use a regular expression as an entity filter, you must specify both the `RegExFilter` and `FilterName` attributes.</span></span>

<br/>

|<span data-ttu-id="992c2-168">Nom de l’attribut</span><span class="sxs-lookup"><span data-stu-id="992c2-168">Attribute name</span></span>|<span data-ttu-id="992c2-169">Description</span><span class="sxs-lookup"><span data-stu-id="992c2-169">Description</span></span>|
|:-----|:-----|
|`EntityType`|<span data-ttu-id="992c2-170">Spécifie le type d’entité à rechercher pour que la règle donne la valeur `true`.</span><span class="sxs-lookup"><span data-stu-id="992c2-170">Specifies the type of entity that must be found for the rule to evaluate to `true`.</span></span> <span data-ttu-id="992c2-171">Utilisez plusieurs règles pour spécifier plusieurs types d’entités.</span><span class="sxs-lookup"><span data-stu-id="992c2-171">Use multiple rules to specify multiple types of entities.</span></span>|
|`RegExFilter`|<span data-ttu-id="992c2-172">Spécifie une expression régulière qui filtre les instances de l’entité spécifiée par `EntityType`.</span><span class="sxs-lookup"><span data-stu-id="992c2-172">Specifies a regular expression that further filters instances of the entity specified by `EntityType`.</span></span>|
|`FilterName`|<span data-ttu-id="992c2-173">Spécifie le nom de l’expression régulière spécifiée par `RegExFilter`, afin qu’il soit possible d’y faire référence ultérieurement par code.</span><span class="sxs-lookup"><span data-stu-id="992c2-173">Specifies the name of the regular expression specified by `RegExFilter`, so that it is subsequently possible to refer to it by code.</span></span>|
|`IgnoreCase`|<span data-ttu-id="992c2-174">Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par `RegExFilter`.</span><span class="sxs-lookup"><span data-stu-id="992c2-174">Specifies whether to ignore case when matching the regular expression specified by `RegExFilter`.</span></span>|

### <a name="examples"></a><span data-ttu-id="992c2-175">Exemples</span><span class="sxs-lookup"><span data-stu-id="992c2-175">Examples</span></span>

<span data-ttu-id="992c2-176">La règle `ItemHasKnownEntity` suivante active le complément chaque fois qu’une URL se trouve dans l’objet ou le corps de l’élément actuel, et qu’elle contient la chaîne `youtube`, indépendamment de la casse de cette chaîne.</span><span class="sxs-lookup"><span data-stu-id="992c2-176">The following `ItemHasKnownEntity` rule activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string `youtube`, regardless of the case of the string.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a><span data-ttu-id="992c2-177">Utilisation des résultats d’expressions régulières dans le code</span><span class="sxs-lookup"><span data-stu-id="992c2-177">Using regular expression results in code</span></span>

<span data-ttu-id="992c2-178">Vous pouvez obtenir des correspondances avec une expression régulière en utilisant les méthodes suivantes sur l’élément actif :</span><span class="sxs-lookup"><span data-stu-id="992c2-178">You can obtain matches to a regular expression by using the following methods on the current item:</span></span>

- <span data-ttu-id="992c2-179">[getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) renvoie les correspondances dans l’élément actuel pour toutes les expressions régulières spécifiées dans les règles `ItemHasRegularExpressionMatch` et `ItemHasKnownEntity` du complément.</span><span class="sxs-lookup"><span data-stu-id="992c2-179">[getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for all regular expressions specified in `ItemHasRegularExpressionMatch` and `ItemHasKnownEntity` rules of the add-in.</span></span>

- <span data-ttu-id="992c2-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) renvoie les correspondances dans l’élément actuel pour l’expression régulière identifiée, spécifiée dans une règle `ItemHasRegularExpressionMatch` du complément.</span><span class="sxs-lookup"><span data-stu-id="992c2-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for the identified regular expression specified in an `ItemHasRegularExpressionMatch` rule of the add-in.</span></span>

- <span data-ttu-id="992c2-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) renvoie les instances complètes des entités qui contiennent des correspondances avec l’expression régulière identifiée, spécifiée dans une règle `ItemHasKnownEntity` du complément.</span><span class="sxs-lookup"><span data-stu-id="992c2-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns entire instances of entities that contain matches for the identified regular expression specified in an `ItemHasKnownEntity` rule of the add-in.</span></span>

<span data-ttu-id="992c2-182">Lorsque les expressions régulières sont évaluées, les correspondances sont renvoyées vers votre complément dans un objet tableau.</span><span class="sxs-lookup"><span data-stu-id="992c2-182">When the regular expressions are evaluated, the matches are returned to your add-in in an array object.</span></span> <span data-ttu-id="992c2-183">Pour `getRegExMatches`, cet objet a un identifiant correspondant au nom de l’expression régulière.</span><span class="sxs-lookup"><span data-stu-id="992c2-183">For `getRegExMatches`, that object has the identifier of the name of the regular expression.</span></span>

> [!NOTE]
> <span data-ttu-id="992c2-184">Les correspondances renvoyées par Outlook ne sont pas classées dans un ordre particulier dans le tableau.</span><span class="sxs-lookup"><span data-stu-id="992c2-184">Outlook does not return matches in any particular order in the array.</span></span> <span data-ttu-id="992c2-185">Par ailleurs, vous ne devez pas supposer que les correspondances sont renvoyées dans le même ordre dans ce tableau, même lorsque vous exécutez le même complément sur chacun de ces clients sur le même élément de la même boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="992c2-185">Also, you should not assume that matches are returned in the same order in this array even when you run the same add-in on each of these clients on the same item in the same mailbox.</span></span>

### <a name="examples"></a><span data-ttu-id="992c2-186">Exemples</span><span class="sxs-lookup"><span data-stu-id="992c2-186">Examples</span></span>

<span data-ttu-id="992c2-187">L’exemple suivant montre un regroupement de règles qui contient une règle `ItemHasRegularExpressionMatch` avec une expression régulière nommée `videoURL`.</span><span class="sxs-lookup"><span data-stu-id="992c2-187">The following is an example of a rule collection that contains an  `ItemHasRegularExpressionMatch` rule with a regular expression named `videoURL`.</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

<span data-ttu-id="992c2-188">L’exemple suivant utilise `getRegExMatches` dans l’élément actuel pour définir une variable `videos` pour les résultats de la règle `ItemHasRegularExpressionMatch` précédente.</span><span class="sxs-lookup"><span data-stu-id="992c2-188">The following example uses `getRegExMatches` of the current item to set a variable `videos` to the results of the preceding `ItemHasRegularExpressionMatch` rule.</span></span>

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

<span data-ttu-id="992c2-p119">Plusieurs correspondances sont stockées comme éléments d’un tableau dans cet objet. L’exemple de code suivant montre comment réaliser une itération sur les correspondances pour une expression régulière nommée  `reg1` pour construire une chaîne à afficher sous la forme HTML.</span><span class="sxs-lookup"><span data-stu-id="992c2-p119">Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.</span></span>

```js
function initDialer()
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

<br/>

<span data-ttu-id="992c2-191">Voici un exemple de règle `ItemHasKnownEntity` qui spécifie l’entité `MeetingSuggestion` et une expression régulière nommée `CampSuggestion`.</span><span class="sxs-lookup"><span data-stu-id="992c2-191">The following is an example of an `ItemHasKnownEntity` rule that specifies the `MeetingSuggestion` entity and a regular expression named `CampSuggestion`.</span></span> <span data-ttu-id="992c2-192">Outlook active le complément s’il détecte que l’élément sélectionné contient une suggestion de réunion, et que l’objet ou le corps contient le terme `WonderCamp`.</span><span class="sxs-lookup"><span data-stu-id="992c2-192">Outlook activates the add-in if it detects that the currently selected item contains a meeting suggestion, and the subject or body contains the term `WonderCamp`.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

<span data-ttu-id="992c2-193">L’exemple de code suivant utilise `getFilteredEntitiesByName` sur l’élément actuel pour définir une variable `suggestions` pour un tableau des suggestions de réunion détectées pour la règle `ItemHasKnownEntity` précédente.</span><span class="sxs-lookup"><span data-stu-id="992c2-193">The following code example uses `getFilteredEntitiesByName` on the current item to set a variable `suggestions` to an array of detected meeting suggestions for the preceding `ItemHasKnownEntity` rule.</span></span>

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a><span data-ttu-id="992c2-194">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="992c2-194">See also</span></span>

- <span data-ttu-id="992c2-195">[Complément Outlook : numéro de commande Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - Exemple de complément contextuel qui est activé en fonction d’une correspondance d’expression régulière.</span><span class="sxs-lookup"><span data-stu-id="992c2-195">[Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - A sample contextual add-in that activates based on a regular expression match.</span></span>
- [<span data-ttu-id="992c2-196">Créer des compléments Outlook pour des formulaires de lecture</span><span class="sxs-lookup"><span data-stu-id="992c2-196">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="992c2-197">Règles d’activation pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="992c2-197">Activation rules for Outlook add-ins</span></span>](activation-rules.md)
- [<span data-ttu-id="992c2-198">Limites pour l’activation et l’API JavaScript pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="992c2-198">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="992c2-199">Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues</span><span class="sxs-lookup"><span data-stu-id="992c2-199">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="992c2-200">Meilleures pratiques pour les expressions régulières dans .NET Framework</span><span class="sxs-lookup"><span data-stu-id="992c2-200">Best Practices for Regular Expressions in the .NET Framework</span></span>](/dotnet/standard/base-types/best-practices)
