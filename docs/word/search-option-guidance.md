---
title: Utiliser les options de recherche pour trouver du texte dans votre complément Word
description: ''
ms.date: 7/20/2018
ms.openlocfilehash: d81ffdcec49d59c175c3e5ecdf82ad1f796fdb3e
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944099"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>Utilisez les options de recherche pour trouver du texte dans votre complément Word 

Les compléments doivent fréquemment agir sur la base du texte d'un document.
Une fonction de recherche est exposée par chaque contrôle de contenu (ceci inclut [le corps](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js), [le paragraphe](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js), [la plage](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js), [la table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js)et l'objet [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) de base). Cette fonction prend une chaîne (ou une expression en caractère générique) représentant le texte que vous recherchez et [un objet](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) SearchOptions. Elle renvoie une collection de plages correspondant au texte recherché.

## <a name="search-options"></a>Options de recherche
Les options de recherche sont une collection de valeurs booléennes définissant comment le paramètre de recherche doit être traité. 

| Propriété     | Description|
|:---------------|:----|
|ignorePunct|Obtient ou définit une valeur indiquant s'il faut ignorer tous les signes de ponctuation entre les mots. Correspond à la case à cocher « Ignorer les signes de ponctuation » dans la boîte de dialogue Rechercher et remplacer.|
|ignoreSpace|Obtient ou définit une valeur indiquant s'il faut ignorer toutes les espaces entre les mots. Correspond à la case à cocher « Ignorer les espaces blancs » dans la boîte de dialogue Rechercher et remplacer.|
|matchCase|Obtient ou définit une valeur indiquant s'il faut effectuer une recherche sensible à la casse. Correspond à la case « Respecter la casse » dans la boîte de dialogue Rechercher et remplacer.|
|matchPrefix|Obtient ou définit une valeur indiquant s'il faut faire correspondre les mots qui commencent par la chaîne de recherche. Correspond à la case à cocher « faites correspondre au préfixe » dans la boîte de dialogue Recherchez et remplacez.|
|matchSuffix|matchSuffix Correspond à la case à cocher « Faire correspondre le suffixe » dans la boîte de dialogue Rechercher et Remplacer.|
|matchWholeWord|Obtient ou définit une valeur indiquant si l'opération doit trouver uniquement des phrases entières, et non un texte faisant partie d'un ensemble de mots. Correspond à la case à cocher « Ne rechercher que des mots entiers » dans la boîte de dialogue Rechercher et remplacer.|
|matchWildCards|Obtient ou définit une valeur indiquant si la recherche sera effectuée à l'aide d'opérateurs de recherche spéciaux. Correspond à la case « Utilisez des caractères génériques » dans la boîte de dialogue Rechercher et remplacer.|

## <a name="wildcard-guidance"></a>Aide concernant les caractères génériques
Le tableau suivant fournit des indications sur les caractères génériques de recherche de l'API JavaScript Word.

| Pour trouver :         | Caractère générique |  Exemple |
|:-----------------|:--------|:----------|
| Un seul caractère| ? |s?t trouve sot et set. |
|Une chaîne de caractères| * |s*n son et solution.|
|Début d’un mot|< |<(intér) trouve intéressant et intérieur, mais pas désintéressé.|
|Fin d’un mot |> |(in)> trouve fin et besoin, mais pas origine.|
|Un des caractères spécifiés|[ ] |l[ea]s trouve les et las.|
|Tout caractère de cette plage| [-] |[b-d]arder trouve barder, carder et darder. Les plages doivent être définies dans l’ordre alphabétique ou croissant.|
|Tout caractère à l’exception de ceux de la plage entre les crochets|[!x-z] |p[!a-m]re trouve pore et pure, mais pas pare et pire.|
|Exactement n occurrences de l’expression ou du caractère précédent|{n} |fe{2}d trouve feed et non fed.|
|Au moins n occurrences de l’expression ou du caractère qui précéde|{n,} |fe{1,}d trouve fed et feed.|
|De n à m occurrences de l’expression ou du caractère qui précéde|{n,m} |10{1,3} trouve 10, 100 et 1000.|
|Une ou plusieurs occurrences de l’expression ou du caractère qui précéde|@ |mar@e trouve mare et marre.|

### <a name="escaping-the-special-characters"></a>Échappement des caractères spéciaux

La recherche avec des caractères génériques est essentiellement la même que la recherche sur une expression régulière. Il existe des caractères spéciaux dans les expressions régulières, notamment « [ », « ] », « ( »,« ) », « { », « } », « \* », « ? », « < », « > », « ! » et « @ ». Si l’un de ces caractères fait partie de la chaîne littérale que recherche le code, il doit être échappé, afin que Word sache qu’il faut le traiter littéralement et non dans le cadre de la logique de l’expression régulière. Pour échapper un caractère dans la fonction de recherche de l’interface utilisateur de Word, faites-le précéder d’un « \' », mais pour un échappement par programme, placez-le entre les caractères « [] ». Par exemple, « [\*]\* » recherche une chaîne qui commence par « \* », suivie d’autres caractères. 

## <a name="examples"></a>Exemples
Les exemples suivants illustrent des scénarios courants.

### <a name="ignore-punctuation-search"></a>Ignorer les signes de ponctuation dans la recherche

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

### <a name="search-based-on-a-prefix"></a>Effectuer une recherche de préfixe

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

### <a name="search-based-on-a-suffix"></a>Effectuer une recherche de suffixe

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

### <a name="search-using-a-wildcard"></a>Effectuer une recherche à l’aide d’un caractère générique

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

Pour en savoir plus, allez dans [l'API de référence JavaScript Word](https://docs.microsoft.com/javascript/office/overview/word-add-ins-reference-overview?view=office-js).