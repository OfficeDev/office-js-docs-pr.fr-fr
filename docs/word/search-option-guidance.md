---
title: Utiliser les options de recherche pour trouver du texte dans votre complément Word
description: ''
ms.date: 7/20/2018
ms.openlocfilehash: 9dcd5e42de9cc0816797a4a14b40a0e3e376f158
ms.sourcegitcommit: eea7f2b1679cf9a209d35880b906e311bdf1359c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/26/2018
ms.locfileid: "21254860"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="1d5b0-102">Utilisez les options de recherche pour trouver du texte dans votre complément Word</span><span class="sxs-lookup"><span data-stu-id="1d5b0-102">Use search options to find text in your Word add-in</span></span> 

<span data-ttu-id="1d5b0-103">Les compléments doivent fréquemment agir sur la base du texte d'un document.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-103">Add-ins frequently need to act based on the text of a document.</span></span>
<span data-ttu-id="1d5b0-104">Une fonction de recherche est exposée par chaque contrôle de contenu (ceci inclut [le corps](https://dev.office.com/reference/add-ins/word/body), [le paragraphe](https://dev.office.com/reference/add-ins/word/paragraph), [la plage](https://dev.office.com/reference/add-ins/word/range), [la table](https://dev.office.com/reference/add-ins/word/table), [TableRow](https://dev.office.com/reference/add-ins/word/tablerow)et l'objet [ContentControl](https://dev.office.com/reference/add-ins/word/contentcontrol) de base).</span><span class="sxs-lookup"><span data-stu-id="1d5b0-104">A search function is exposed by every content control (this includes [Body](https://dev.office.com/reference/add-ins/word/body), [Paragraph](https://dev.office.com/reference/add-ins/word/paragraph), [Range](https://dev.office.com/reference/add-ins/word/range), [Table](https://dev.office.com/reference/add-ins/word/table), [TableRow](https://dev.office.com/reference/add-ins/word/tablerow), and the base [ContentControl](https://dev.office.com/reference/add-ins/word/contentcontrol) object).</span></span> <span data-ttu-id="1d5b0-105">Cette fonction prend une chaîne (ou une expression en caractère générique) représentant le texte que vous recherchez et [un objet](https://dev.office.com/reference/add-ins/word/searchoptions) SearchOptions.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-105">This function takes in a string (or wldcard expression) representing the text you are searching for and a [SearchOptions](https://dev.office.com/reference/add-ins/word/searchoptions) object.</span></span> <span data-ttu-id="1d5b0-106">Elle renvoie une collection de plages correspondant au texte recherché.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-106">It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="1d5b0-107">Options de recherche</span><span class="sxs-lookup"><span data-stu-id="1d5b0-107">Search options</span></span>
<span data-ttu-id="1d5b0-108">Les options de recherche sont une collection de valeurs booléennes définissant comment le paramètre de recherche doit être traité.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-108">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span> 

| <span data-ttu-id="1d5b0-109">Propriété</span><span class="sxs-lookup"><span data-stu-id="1d5b0-109">Property</span></span>     | <span data-ttu-id="1d5b0-110">Description</span><span class="sxs-lookup"><span data-stu-id="1d5b0-110">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="1d5b0-111">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="1d5b0-111">ignorePunct</span></span>|<span data-ttu-id="1d5b0-112">Obtient ou définit une valeur indiquant s'il faut ignorer tous les signes de ponctuation entre les mots.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-112">Gets or sets a value indicating whether to ignore all punctuation characters between words.</span></span> <span data-ttu-id="1d5b0-113">Correspond à la case « Ignore les signes de ponctuation » dans la boîte de dialogue Rechercher et remplacer.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-113">True ignores all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="1d5b0-114">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="1d5b0-114">ignoreSpace</span></span>|<span data-ttu-id="1d5b0-115">Obtient ou définit une valeur indiquant s'il faut ignorer toutes les espaces entre les mots.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-115">Gets or sets a value indicating whether to ignore all whitespace between words.</span></span> <span data-ttu-id="1d5b0-116">Correspond à la case à cocher « Ignorez les espaces blanc » dans la boîte de dialogue Recherchez et Remplacez.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-116">Corresponds to the Ignore white-space characters check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="1d5b0-117">matchCase</span><span class="sxs-lookup"><span data-stu-id="1d5b0-117">matchCase</span></span>|<span data-ttu-id="1d5b0-118">Obtient ou définit une valeur indiquant s'il faut effectuer une recherche sensible à la casse.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-118">Gets or sets a value indicating whether to perform a case sensitive search.</span></span> <span data-ttu-id="1d5b0-119">Correspond à la case « Respecter la casse » dans la boîte de dialogue Rechercher et Remplacer.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-119">Corresponds to the Match case check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="1d5b0-120">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="1d5b0-120">matchPrefix</span></span>|<span data-ttu-id="1d5b0-121">Obtient ou définit une valeur indiquant s'il faut faire correspondre les mots qui commencent par la chaîne de recherche.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-121">Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="1d5b0-122">Correspond à la case à cocher « faites correspondre au préfixe » dans la boîte de dialogue Recherchez et remplacez.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-122">Corresponds to the Match prefix check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="1d5b0-123">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="1d5b0-123">matchSuffix</span></span>|<span data-ttu-id="1d5b0-124">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="1d5b0-124">Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="1d5b0-125">Correspond à la case à cocher « Faire correspondre le suffixe » dans la boîte de dialogue Rechercher et Remplacer.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-125">Corresponds to the Match suffix check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="1d5b0-126">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="1d5b0-126">matchWholeWord</span></span>|<span data-ttu-id="1d5b0-127">Obtient ou définit une valeur indiquant si l'opération doit trouver uniquement des phrases entières, et non un texte faisant partie d'un ensemble de mots.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-127">Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="1d5b0-128">Correspond à la case à cocher « Rechercher des phrases entières » dans la boîte de dialogue Recherchez et remplacez.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-128">Corresponds to the Find whole words only check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="1d5b0-129">matchWildCards</span><span class="sxs-lookup"><span data-stu-id="1d5b0-129">matchWildcards</span></span>|<span data-ttu-id="1d5b0-130">Obtient ou définit une valeur indiquant si la recherche sera effectuée à l'aide d'opérateurs de recherche spéciaux.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-130">Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="1d5b0-131">Correspond à la case « Utilisez des caractères génériques » dans la boîte de dialogue Rechercher et remplacer.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-131">Corresponds to the Use wildcards check box in the Find and Replace dialog box.</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="1d5b0-132">Aide concernant les caractères génériques</span><span class="sxs-lookup"><span data-stu-id="1d5b0-132">Wildcard Guidance</span></span>
<span data-ttu-id="1d5b0-133">Le tableau suivant fournit des indications sur les caractères génériques de recherche de l'API JavaScript Word.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-133">The following table provides guidance around the Word JavaScript API’s search wildcards.</span></span>

| <span data-ttu-id="1d5b0-134">Pour rechercher :</span><span class="sxs-lookup"><span data-stu-id="1d5b0-134">To find:</span></span>         | <span data-ttu-id="1d5b0-135">Caractère générique</span><span class="sxs-lookup"><span data-stu-id="1d5b0-135">Wildcard</span></span> |  <span data-ttu-id="1d5b0-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d5b0-136">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="1d5b0-137">Un seul caractère</span><span class="sxs-lookup"><span data-stu-id="1d5b0-137">Any single character</span></span>| <span data-ttu-id="1d5b0-138">?</span><span class="sxs-lookup"><span data-stu-id="1d5b0-138"></span></span> |<span data-ttu-id="1d5b0-139">s?t trouve sot et set.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-139">s?t finds sat and set.</span></span> |
|<span data-ttu-id="1d5b0-140">Une chaîne de caractères</span><span class="sxs-lookup"><span data-stu-id="1d5b0-140">Any string of characters</span></span>| * |<span data-ttu-id="1d5b0-141">s\*n son et solution.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-141">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="1d5b0-142">Début d’un mot</span><span class="sxs-lookup"><span data-stu-id="1d5b0-142">The beginning of a word</span></span>|< |<span data-ttu-id="1d5b0-143"><(intér) trouve intéressant et intérieur, mais pas désintéressé.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-143"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="1d5b0-144">Fin d’un mot</span><span class="sxs-lookup"><span data-stu-id="1d5b0-144">The end of a word</span></span> |> |<span data-ttu-id="1d5b0-145">(in)> trouve fin et besoin, mais pas origine.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-145">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="1d5b0-146">Un des caractères spécifiés</span><span class="sxs-lookup"><span data-stu-id="1d5b0-146">One of the specified characters</span></span>|<span data-ttu-id="1d5b0-147">[ ]</span><span class="sxs-lookup"><span data-stu-id="1d5b0-147"></span></span> |<span data-ttu-id="1d5b0-148">l[ea]s trouve les et las.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-148">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="1d5b0-149">Tout caractère de cette plage</span><span class="sxs-lookup"><span data-stu-id="1d5b0-149">Any single character in this range</span></span>| <span data-ttu-id="1d5b0-150">[-]</span><span class="sxs-lookup"><span data-stu-id="1d5b0-150"></span></span> |<span data-ttu-id="1d5b0-p109">[b-d]arder trouve barder, carder et darder. Les plages doivent être définies dans l’ordre alphabétique ou croissant.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="1d5b0-153">Tout caractère à l’exception de ceux de la plage entre les crochets</span><span class="sxs-lookup"><span data-stu-id="1d5b0-153">Any single character except the characters in the range inside the brackets</span></span>|[!x-z] |<span data-ttu-id="1d5b0-155">p[!a-m]re trouve pore et pure, mais pas pare et pire.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-155">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="1d5b0-156">Exactement n occurrences de l’expression ou du caractère précédent</span><span class="sxs-lookup"><span data-stu-id="1d5b0-156">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="1d5b0-157">{n}</span><span class="sxs-lookup"><span data-stu-id="1d5b0-157">{n}</span></span> |<span data-ttu-id="1d5b0-158">fe{2}d recherche feed et non fed.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-158">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="1d5b0-159">Au moins n occurrences du caractère ou de l’expression précédente</span><span class="sxs-lookup"><span data-stu-id="1d5b0-159">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="1d5b0-160">{n,}</span><span class="sxs-lookup"><span data-stu-id="1d5b0-160">{n,}</span></span> |<span data-ttu-id="1d5b0-161">fe{1,}d recherche fed et non feed.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-161">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="1d5b0-162">Entre n et m occurrences de l’expression ou du caractère précédent</span><span class="sxs-lookup"><span data-stu-id="1d5b0-162">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="1d5b0-163">{n,m}</span><span class="sxs-lookup"><span data-stu-id="1d5b0-163">{n,m}</span></span> |<span data-ttu-id="1d5b0-164">10{1,3} trouve 10, 100 et 1000.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-164">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="1d5b0-165">Une ou plusieurs occurrences de l’expression ou du caractère précédent</span><span class="sxs-lookup"><span data-stu-id="1d5b0-165">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="1d5b0-166">mar@e trouve mare et marre.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-166">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="1d5b0-167">Échappement des caractères spéciaux</span><span class="sxs-lookup"><span data-stu-id="1d5b0-167">Escaping the special characters</span></span>

<span data-ttu-id="1d5b0-p110">La recherche avec des caractères génériques est essentiellement la même que la recherche sur une expression régulière. Il existe des caractères spéciaux dans les expressions régulières, notamment « [ », « ] », « ( »,« ) », « { », « } », « \* », « ? », « < », « > », « ! » et « @ ». Si l’un de ces caractères fait partie de la chaîne littérale que recherche le code, il doit être échappé, afin que Word sache qu’il faut le traiter littéralement et non dans le cadre de la logique de l’expression régulière. Pour échapper un caractère dans la fonction de recherche de l’interface utilisateur de Word, faites-le précéder d’un « \' », mais pour un échappement par programme, placez-le entre les caractères « [] ». Par exemple, « [\*]\* » recherche une chaîne qui commence par « \* », suivie d’autres caractères.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="1d5b0-173">Exemples</span><span class="sxs-lookup"><span data-stu-id="1d5b0-173">Examples</span></span>
<span data-ttu-id="1d5b0-174">Les exemples suivants illustrent des scénarios courants.</span><span class="sxs-lookup"><span data-stu-id="1d5b0-174">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="1d5b0-175">Ignorer les signes de ponctuation dans la recherche</span><span class="sxs-lookup"><span data-stu-id="1d5b0-175">Ignore punctuation search</span></span>

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

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="1d5b0-176">Effectuer une recherche de préfixe</span><span class="sxs-lookup"><span data-stu-id="1d5b0-176">Search based on a prefix</span></span>

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

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="1d5b0-177">Effectuer une recherche de suffixe</span><span class="sxs-lookup"><span data-stu-id="1d5b0-177">Search based on a suffix</span></span>

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

### <a name="search-using-a-wildcard"></a><span data-ttu-id="1d5b0-178">Effectuer une recherche à l’aide d’un caractère générique</span><span class="sxs-lookup"><span data-stu-id="1d5b0-178">Search using a wildcard</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    var searchResults = context.document.body.search('to*n', {matchWildCards: true});

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

<span data-ttu-id="1d5b0-179">Pour en savoir plus, allez dans [l'API de référence JavaScript Word](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="1d5b0-179">More information can be found in the [Word JavaScript Reference API](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview).</span></span>