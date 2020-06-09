---
title: Utilisation d’options de recherche pour trouver du texte dans votre complément Word
description: Apprendre à utiliser les options de recherche dans votre complément Word
ms.date: 09/27/2019
localization_priority: Normal
ms.openlocfilehash: 197ed066d7eece584e8e14febd0ca8f5b4f0c5d0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608571"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="5475d-103">Utilisation d’options de recherche pour trouver du texte dans votre complément Word</span><span class="sxs-lookup"><span data-stu-id="5475d-103">Use search options to find text in your Word add-in</span></span>

<span data-ttu-id="5475d-104">Les compléments doivent fréquemment agir en fonction du texte d’un document.</span><span class="sxs-lookup"><span data-stu-id="5475d-104">Add-ins frequently need to act based on the text of a document.</span></span>
<span data-ttu-id="5475d-105">Une fonction de recherche est exposée par contrôle de contenu (cela inclut [Body](/javascript/api/word/word.body), [Paragraph](/javascript/api/word/word.paragraph), [Range](/javascript/api/word/word.range), [Table](/javascript/api/word/word.table), [TableRow](/javascript/api/word/word.tablerow) et l’objet de base [ContentControl](/javascript/api/word/word.contentcontrol)).</span><span class="sxs-lookup"><span data-stu-id="5475d-105">A search function is exposed by every content control (this includes [Body](/javascript/api/word/word.body), [Paragraph](/javascript/api/word/word.paragraph), [Range](/javascript/api/word/word.range), [Table](/javascript/api/word/word.table), [TableRow](/javascript/api/word/word.tablerow), and the base [ContentControl](/javascript/api/word/word.contentcontrol) object).</span></span> <span data-ttu-id="5475d-106">Cette fonction utilise une chaîne (ou une expression générique) représentant le texte que vous recherchez et un objet [SearchOptions](/javascript/api/word/word.searchoptions).</span><span class="sxs-lookup"><span data-stu-id="5475d-106">This function takes in a string (or wildcard expression) representing the text you are searching for and a [SearchOptions](/javascript/api/word/word.searchoptions) object.</span></span> <span data-ttu-id="5475d-107">Elle renvoie une collection de plages correspondant au texte de recherche.</span><span class="sxs-lookup"><span data-stu-id="5475d-107">It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="5475d-108">Options de recherche</span><span class="sxs-lookup"><span data-stu-id="5475d-108">Search options</span></span>

<span data-ttu-id="5475d-109">Les options de recherche sont une collection de valeurs booléennes qui définissent comment le paramètre de recherche doit être traité.</span><span class="sxs-lookup"><span data-stu-id="5475d-109">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span>

| <span data-ttu-id="5475d-110">Propriété</span><span class="sxs-lookup"><span data-stu-id="5475d-110">Property</span></span>       | <span data-ttu-id="5475d-111">Description</span><span class="sxs-lookup"><span data-stu-id="5475d-111">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="5475d-112">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="5475d-112">ignorePunct</span></span>|<span data-ttu-id="5475d-113">Obtient ou définit une valeur indiquant s’il faut ignorer tous les caractères de ponctuation entre les mots.</span><span class="sxs-lookup"><span data-stu-id="5475d-113">Gets or sets a value indicating whether to ignore all punctuation characters between words.</span></span> <span data-ttu-id="5475d-114">Correspond à la case à cocher « Ignorer les caractères de ponctuation » dans la boîte de dialogue Rechercher et remplacer.</span><span class="sxs-lookup"><span data-stu-id="5475d-114">Corresponds to the "Ignore punctuation characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="5475d-115">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="5475d-115">ignoreSpace</span></span>|<span data-ttu-id="5475d-116">Obtient ou définit une valeur indiquant s’il faut ignorer tous les espaces entre les mots.</span><span class="sxs-lookup"><span data-stu-id="5475d-116">Gets or sets a value indicating whether to ignore all whitespace between words.</span></span> <span data-ttu-id="5475d-117">Correspond à la case à cocher « Ignorer les caractères d’espacement » dans la boîte de dialogue Rechercher et remplacer.</span><span class="sxs-lookup"><span data-stu-id="5475d-117">Corresponds to the "Ignore white-space characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="5475d-118">matchCase</span><span class="sxs-lookup"><span data-stu-id="5475d-118">matchCase</span></span>|<span data-ttu-id="5475d-119">Obtient ou définit une valeur indiquant si la recherche respecte la casse.</span><span class="sxs-lookup"><span data-stu-id="5475d-119">Gets or sets a value indicating whether to perform a case sensitive search.</span></span> <span data-ttu-id="5475d-120">Correspond à la case à cocher « Respecter la casse » dans la boîte de dialogue Rechercher et remplacer.</span><span class="sxs-lookup"><span data-stu-id="5475d-120">Corresponds to the "Match case" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="5475d-121">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="5475d-121">matchPrefix</span></span>|<span data-ttu-id="5475d-122">Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui commencent par la chaîne entrée.</span><span class="sxs-lookup"><span data-stu-id="5475d-122">Gets or sets a value indicating whether to match words that begin with the search string.</span></span> <span data-ttu-id="5475d-123">Correspond à la case à cocher « Préfixe » dans la boîte de dialogue Rechercher et remplacer.</span><span class="sxs-lookup"><span data-stu-id="5475d-123">Corresponds to the "Match prefix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="5475d-124">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="5475d-124">matchSuffix</span></span>|<span data-ttu-id="5475d-125">Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui se terminent par la chaîne entrée.</span><span class="sxs-lookup"><span data-stu-id="5475d-125">Gets or sets a value indicating whether to match words that end with the search string.</span></span> <span data-ttu-id="5475d-126">Correspond à la case à cocher « Suffixe » dans la boîte de dialogue Rechercher et remplacer.</span><span class="sxs-lookup"><span data-stu-id="5475d-126">Corresponds to the "Match suffix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="5475d-127">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="5475d-127">matchWholeWord</span></span>|<span data-ttu-id="5475d-128">Obtient ou définit une valeur indiquant si la recherche doit porter uniquement sur les mots complets et non pas sur du texte qui fait partie d’un mot plus long.</span><span class="sxs-lookup"><span data-stu-id="5475d-128">Gets or sets a value indicating whether to find operation only entire words, not text that is part of a larger word.</span></span> <span data-ttu-id="5475d-129">Correspond à la case à cocher « Mot entier » dans la boîte de dialogue Rechercher et remplacer.</span><span class="sxs-lookup"><span data-stu-id="5475d-129">Corresponds to the "Find whole words only" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="5475d-130">matchWildCards</span><span class="sxs-lookup"><span data-stu-id="5475d-130">matchWildcards</span></span>|<span data-ttu-id="5475d-131">Obtient ou définit une valeur indiquant si la recherche est effectuée à l’aide d’opérateurs de recherche spéciaux.</span><span class="sxs-lookup"><span data-stu-id="5475d-131">Gets or sets a value indicating whether the search will be performed using special search operators.</span></span> <span data-ttu-id="5475d-132">Correspond à la case « Utiliser les caractères génériques » dans la boîte de dialogue Rechercher et remplacer.</span><span class="sxs-lookup"><span data-stu-id="5475d-132">Corresponds to the "Use wildcards" check box in the Find and Replace dialog box.</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="5475d-133">Aide concernant les caractères génériques</span><span class="sxs-lookup"><span data-stu-id="5475d-133">Wildcard guidance</span></span>

<span data-ttu-id="5475d-134">Le tableau suivant fournit une aide concernant les caractères génériques de l’API JavaScript pour Word.</span><span class="sxs-lookup"><span data-stu-id="5475d-134">The following table provides guidance around the Word JavaScript API's search wildcards.</span></span>

| <span data-ttu-id="5475d-135">Pour trouver :</span><span class="sxs-lookup"><span data-stu-id="5475d-135">To find:</span></span>         | <span data-ttu-id="5475d-136">Caractère générique</span><span class="sxs-lookup"><span data-stu-id="5475d-136">Wildcard</span></span> |  <span data-ttu-id="5475d-137">Exemple</span><span class="sxs-lookup"><span data-stu-id="5475d-137">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="5475d-138">Un seul caractère</span><span class="sxs-lookup"><span data-stu-id="5475d-138">Any single character</span></span>| <span data-ttu-id="5475d-139">?</span><span class="sxs-lookup"><span data-stu-id="5475d-139">?</span></span> |<span data-ttu-id="5475d-140">s?t trouve sot et set.</span><span class="sxs-lookup"><span data-stu-id="5475d-140">s?t finds sat and set.</span></span> |
|<span data-ttu-id="5475d-141">Une chaîne de caractères</span><span class="sxs-lookup"><span data-stu-id="5475d-141">Any string of characters</span></span>| * |<span data-ttu-id="5475d-142">s\*n son et solution.</span><span class="sxs-lookup"><span data-stu-id="5475d-142">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="5475d-143">Début d’un mot</span><span class="sxs-lookup"><span data-stu-id="5475d-143">The beginning of a word</span></span>|< |<span data-ttu-id="5475d-144"><(intér) trouve intéressant et intérieur, mais pas désintéressé.</span><span class="sxs-lookup"><span data-stu-id="5475d-144"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="5475d-145">Fin d’un mot</span><span class="sxs-lookup"><span data-stu-id="5475d-145">The end of a word</span></span> |> |<span data-ttu-id="5475d-146">(in)> trouve fin et besoin, mais pas origine.</span><span class="sxs-lookup"><span data-stu-id="5475d-146">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="5475d-147">Un des caractères spécifiés</span><span class="sxs-lookup"><span data-stu-id="5475d-147">One of the specified characters</span></span>|<span data-ttu-id="5475d-148">[ ]</span><span class="sxs-lookup"><span data-stu-id="5475d-148">[ ]</span></span> |<span data-ttu-id="5475d-149">l[ea]s trouve les et las.</span><span class="sxs-lookup"><span data-stu-id="5475d-149">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="5475d-150">Tout caractère de cette plage</span><span class="sxs-lookup"><span data-stu-id="5475d-150">Any single character in this range</span></span>| <span data-ttu-id="5475d-151">[-]</span><span class="sxs-lookup"><span data-stu-id="5475d-151">[-]</span></span> |<span data-ttu-id="5475d-p109">[b-d]arder trouve barder, carder et darder. Les plages doivent être définies dans l’ordre alphabétique ou croissant.</span><span class="sxs-lookup"><span data-stu-id="5475d-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="5475d-154">Tout caractère à l’exception de ceux de la plage entre les crochets</span><span class="sxs-lookup"><span data-stu-id="5475d-154">Any single character except the characters in the range inside the brackets</span></span>|[!x-z] |<span data-ttu-id="5475d-156">p[!a-m]re trouve pore et pure, mais pas pare et pire.</span><span class="sxs-lookup"><span data-stu-id="5475d-156">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="5475d-157">Exactement n occurrences de l’expression ou du caractère précédent</span><span class="sxs-lookup"><span data-stu-id="5475d-157">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="5475d-158">n</span><span class="sxs-lookup"><span data-stu-id="5475d-158">{n}</span></span> |<span data-ttu-id="5475d-159">bal{2}ade trouve ballade mais pas balade.</span><span class="sxs-lookup"><span data-stu-id="5475d-159">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="5475d-160">Au moins n occurrences de l’expression ou du caractère précédent</span><span class="sxs-lookup"><span data-stu-id="5475d-160">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="5475d-161">{n,}</span><span class="sxs-lookup"><span data-stu-id="5475d-161">{n,}</span></span> |<span data-ttu-id="5475d-162">bal{1,}ade recherche balade et ballade.</span><span class="sxs-lookup"><span data-stu-id="5475d-162">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="5475d-163">Entre n et m occurrences de l’expression ou du caractère précédent</span><span class="sxs-lookup"><span data-stu-id="5475d-163">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="5475d-164">{n,m}</span><span class="sxs-lookup"><span data-stu-id="5475d-164">{n,m}</span></span> |<span data-ttu-id="5475d-165">10{1,3} trouve 10, 100 et 1 000.</span><span class="sxs-lookup"><span data-stu-id="5475d-165">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="5475d-166">Une ou plusieurs occurrences de l’expression ou du caractère précédent</span><span class="sxs-lookup"><span data-stu-id="5475d-166">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="5475d-167">mar@e trouve mare et marre.</span><span class="sxs-lookup"><span data-stu-id="5475d-167">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="5475d-168">Échappement des caractères spéciaux</span><span class="sxs-lookup"><span data-stu-id="5475d-168">Escaping the special characters</span></span>

<span data-ttu-id="5475d-p110">La recherche avec des caractères génériques est essentiellement la même que la recherche sur une expression régulière. Il existe des caractères spéciaux dans les expressions régulières, notamment « [ », « ] », « ( »,« ) », « { », « } », « \* », « ? », « < », « > », « ! » et « @ ». Si l’un de ces caractères fait partie de la chaîne littérale que recherche le code, il doit être échappé, afin que Word sache qu’il faut le traiter littéralement et non dans le cadre de la logique de l’expression régulière. Pour échapper un caractère dans la fonction de recherche de l’interface utilisateur de Word, faites-le précéder d’un « \' », mais pour un échappement par programme, placez-le entre les caractères « [] ». Par exemple, « [\*]\* » recherche une chaîne qui commence par « \* », suivie d’autres caractères.</span><span class="sxs-lookup"><span data-stu-id="5475d-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="5475d-174">Exemples</span><span class="sxs-lookup"><span data-stu-id="5475d-174">Examples</span></span>

<span data-ttu-id="5475d-175">Les exemples suivants illustrent des scénarios courants.</span><span class="sxs-lookup"><span data-stu-id="5475d-175">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="5475d-176">Ignorer les signes de ponctuation dans la recherche</span><span class="sxs-lookup"><span data-stu-id="5475d-176">Ignore punctuation search</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document and ignore punctuation.
    var searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="5475d-177">Effectuer une recherche de préfixe</span><span class="sxs-lookup"><span data-stu-id="5475d-177">Search based on a prefix</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document based on a prefix.
    var searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="5475d-178">Effectuer une recherche de suffixe</span><span class="sxs-lookup"><span data-stu-id="5475d-178">Search based on a suffix</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document for any string of characters after 'ly'.
    var searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'orange';
            searchResults.items[i].font.highlightColor = 'black';
            searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-using-a-wildcard"></a><span data-ttu-id="5475d-179">Effectuer une recherche à l’aide d’un caractère générique</span><span class="sxs-lookup"><span data-stu-id="5475d-179">Search using a wildcard</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    var searchResults = context.document.body.search('to*n', {matchWildcards: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = 'pink';
            searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

<span data-ttu-id="5475d-180">Vous trouverez plus d’informations dans l’[API JavaScript de référence pour Word](../reference/overview/word-add-ins-reference-overview.md).</span><span class="sxs-lookup"><span data-stu-id="5475d-180">More information can be found in the [Word JavaScript Reference API](../reference/overview/word-add-ins-reference-overview.md).</span></span>
