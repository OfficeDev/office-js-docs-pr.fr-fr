---
title: Utiliser les options de recherche dans votre complément Word pour rechercher du texte
description: Découvrez comment utiliser les options de recherche dans votre complément Word.
ms.date: 02/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 64ffd3b32329dae98f869abaabcb3218e57a4a34
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958999"
---
# <a name="use-search-options-in-your-word-add-in-to-find-text"></a>Utiliser les options de recherche dans votre complément Word pour rechercher du texte

Les compléments doivent fréquemment agir en fonction du texte d’un document. Une méthode de recherche est exposée par chaque contrôle de contenu (cela inclut [Body](/javascript/api/word/word.body), [Paragraph](/javascript/api/word/word.paragraph), [Range](/javascript/api/word/word.range), [Table](/javascript/api/word/word.table), [TableRow](/javascript/api/word/word.tablerow) et l’objet [ContentControl](/javascript/api/word/word.contentcontrol) de base). Cette méthode accepte une chaîne (ou une expression générique) représentant le texte que vous recherchez et un objet [SearchOptions](/javascript/api/word/word.searchoptions) . Elle renvoie une collection de plages correspondant au texte de recherche.

## <a name="search-options"></a>Options de recherche

Les options de recherche sont une collection de valeurs booléennes qui définissent comment le paramètre de recherche doit être traité.

| Propriété       | Description|
|:---------------|:----|
|ignorePunct|Obtient ou définit une valeur indiquant s’il faut ignorer tous les caractères de ponctuation entre les mots. Correspond à la case à cocher « Ignorer les caractères de ponctuation » dans la boîte de dialogue **Rechercher et remplacer** .|
|ignoreSpace|Obtient ou définit une valeur indiquant s’il faut ignorer tous les espaces entre les mots. Correspond à la case à cocher « Ignorer les espaces blancs » dans la boîte de dialogue **Rechercher et remplacer** .|
|matchCase|Obtient ou définit une valeur indiquant s’il faut effectuer une recherche respectant la casse. Correspond à la case à cocher « Cas de correspondance » dans la boîte de dialogue **Rechercher et remplacer** .|
|matchPrefix|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui commencent par la chaîne entrée. Correspond à la case à cocher « Mettre en correspondance le préfixe » dans la boîte de dialogue **Rechercher et remplacer** .|
|matchSuffix|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui se terminent par la chaîne entrée. Correspond à la case à cocher « Mettre en correspondance le suffixe » dans la boîte de dialogue **Rechercher et remplacer** .|
|matchWholeWord|Obtient ou définit une valeur indiquant si la recherche doit porter uniquement sur les mots complets et non pas sur du texte qui fait partie d’un mot plus long. Correspond à la case à cocher « Rechercher des mots entiers uniquement » dans la boîte de dialogue **Rechercher et remplacer** .|
|matchWildCards|Obtient ou définit une valeur indiquant si la recherche est effectuée à l’aide d’opérateurs de recherche spéciaux. Correspond à la case à cocher « Utiliser des caractères génériques » dans la boîte de dialogue **Rechercher et remplacer** .|

## <a name="wildcard-guidance"></a>Aide concernant les caractères génériques

Le tableau suivant fournit une aide concernant les caractères génériques de l’API JavaScript pour Word.

| Pour trouver :         | Caractère générique |  Exemple |
|:-----------------|:--------|:----------|
|Un seul caractère| ? |s?t trouve sot et set. |
|Une chaîne de caractères| * |s*n son et solution.|
|Début d’un mot|< |<(intér) trouve intéressant et intérieur, mais pas désintéressé.|
|Fin d’un mot |> |(in)> trouve fin et besoin, mais pas origine.|
|Un des caractères spécifiés|[ ] |l[ea]s trouve les et las.|
|Tout caractère de cette plage| [-] |[b-d]arder trouve barder, carder et darder. Les plages doivent être définies dans l’ordre alphabétique ou croissant.|
|Tout caractère à l’exception de ceux de la plage entre les crochets|[!x-z] |p[!a-m]re trouve pore et pure, mais pas pare et pire.|
|Exactement *n* occurrences du caractère ou de l’expression précédent|n |bal{2}ade trouve ballade mais pas balade.|
|Au moins *n* occurrences du caractère ou de l’expression précédent|{n,} |bal{1,}ade recherche balade et ballade.|
|De *n* à *m* occurrences du caractère ou de l’expression précédent|{n,m} |10{1,3} trouve 10, 100 et 1 000.|
|Une ou plusieurs occurrences de l’expression ou du caractère précédent|@ |mar@e trouve mare et marre.|

### <a name="escaping-special-characters"></a>Échappement de caractères spéciaux

La recherche par caractères génériques est essentiellement la même que la recherche sur une expression régulière. Il existe des caractères spéciaux dans les expressions régulières, notamment « [ », « ] », « ( », « ( », « { », « } », « »,\* « ? », « < », « > », « ! » et « @ ». Si l’un de ces caractères fait partie de la chaîne littérale que le code recherche, il doit être placé dans une séquence d’échappement, afin que Word sache qu’il doit être traité littéralement et non dans le cadre de la logique de l’expression régulière. Pour échapper un caractère dans la recherche de l’interface utilisateur Word, vous le précédez d’un caractère de barre oblique inverse («\\ ») mais pour l’échapper par programmation, placez-le entre des caractères « [] ». Par exemple, « [\*]\* » recherche une chaîne commençant par un «\* » suivi d’un nombre quelconque d’autres caractères.

## <a name="examples"></a>Exemples

Les exemples suivants illustrent des scénarios courants.

### <a name="ignore-punctuation-search"></a>Ignorer les signes de ponctuation dans la recherche

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document and ignore punctuation.
    const searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-prefix"></a>Effectuer une recherche de préfixe

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document based on a prefix.
    const searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-suffix"></a>Effectuer une recherche de suffixe

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document for any string of characters after 'ly'.
    const searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'orange';
        searchResults.items[i].font.highlightColor = 'black';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-using-a-wildcard"></a>Effectuer une recherche à l’aide d’un caractère générique

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    const searchResults = context.document.body.search('to*n', {matchWildcards: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = 'pink';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

Vous trouverez plus d’informations dans l’[API JavaScript de référence pour Word](../reference/overview/word-add-ins-reference-overview.md).
