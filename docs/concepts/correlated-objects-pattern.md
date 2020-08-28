---
title: Éviter d’utiliser la méthode context.sync dans des boucles
description: Découvrez comment utiliser les modèles de boucle de fractionnement et d’objets corrélés pour éviter d’appeler Context. Sync dans une boucle.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: d3628400ef783035cf6a816144dbd5cfb30582ee
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292994"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>Éviter d’utiliser la méthode context.sync dans des boucles

> [!NOTE]
> Cet article suppose que vous êtes au-delà de la phase de démarrage de l’utilisation d’au moins l’une des quatre API JavaScript Office spécifiques &mdash; aux applications pour Excel, Word, OneNote et Visio &mdash; qui utilisent un système de traitement par lots pour interagir avec le document Office. En particulier, vous devez savoir ce qu’est un appel de `context.sync` et savoir ce qu’est un objet de la collection. Si vous n’êtes pas à ce stade, commencez par [comprendre l’API JavaScript pour Office](../develop/understanding-the-javascript-api-for-office.md) et la documentation liée à la section « propre à l’application » dans cet article.

Pour certains scénarios de programmation dans les compléments Office qui utilisent l’un des modèles d’API propres aux applications (pour Excel, Word, OneNote et Visio), votre code doit lire, écrire ou traiter certaines propriétés à partir de chaque membre d’un objet collection. Par exemple, un complément Excel qui doit obtenir les valeurs de chaque cellule d’une colonne de table particulière ou d’un complément Word qui doit mettre en surbrillance chaque instance d’une chaîne dans le document. Vous devez effectuer une itération sur les membres dans la `items` propriété de l’objet de collection ; Toutefois, pour des raisons de performances, vous devez éviter d’appeler `context.sync` dans chaque itération de la boucle. Chaque appel de `context.sync` est un aller-retour entre le complément et le document Office. Les allers-retours répétés ont un impact sur les performances, en particulier si le complément est exécuté dans Office sur le Web, car les allers-retours sont effectués sur Internet.

> [!NOTE]
> Tous les exemples de cet article utilisent `for` des boucles, mais les pratiques décrites s’appliquent à toutes les instructions Loop qui peuvent parcourir un tableau, notamment les suivantes :
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> Elles s’appliquent également à toute méthode Array à laquelle une fonction est passée et appliquée aux éléments du tableau, notamment les suivantes :
>
> - `Array.every`
> - `Array.forEach`
> - `Array.filter`
> - `Array.find`
> - `Array.findIndex`
> - `Array.map`
> - `Array.reduce`
> - `Array.reduceRight`
> - `Array.some`

## <a name="writing-to-the-document"></a>Écriture dans le document

Dans le cas le plus simple, vous écrivez uniquement aux membres d’un objet de collection, et non à la lecture de leurs propriétés. Par exemple, le code suivant met en évidence en jaune toutes les occurrences de « The » dans un document Word.

> [!NOTE]
> Il est généralement recommandé de placer un final `context.sync` juste avant le caractère « } » de la méthode d’application `run` (par exemple `Excel.run` , `Word.run` , etc.). Cela est dû au fait que la `run` méthode effectue un appel masqué de `context.sync` la dernière chose qu’elle effectue si, et seulement si, il existe des commandes en file d’attente qui n’ont pas encore été synchronisées. Le fait que cet appel soit masqué peut prêter à confusion, c’est pourquoi nous vous recommandons généralement d’ajouter le explicite `context.sync` . Toutefois, étant donné que cet article concerne la réduction des appels de `context.sync` , il est en fait plus déroutant d’ajouter une final entièrement inutile `context.sync` . Par conséquent, dans cet article, nous les laissons quand il n’y a pas de commandes non synchronisées à la fin du `run` .

```javascript
Word.run(async function (context) {
    let startTime, endTime;
    const docBody = context.document.body;

    // search() returns an array of Ranges.
    const searchResults = docBody.search('the', { matchWholeWord: true });
    context.load(searchResults, 'items');
    await context.sync();

    // Record the system time.
    startTime = performance.now();

    for (var i = 0; i < searchResults.items.length; i++) {
      searchResults.items[i].font.highlightColor = '#FFFF00';

      await context.sync(); // SYNCHRONIZE IN EACH ITERATION
    }
    
    // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

    // Record the system time again then calculate how long the operation took.
    endTime = performance.now();
    console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
  })
}
```

Le code précédent prenait une seconde complète dans un document avec 200 instances de « The » dans Word sur Windows. Toutefois, lorsque la `await context.sync();` ligne à l’intérieur de la boucle est commentée et que la même ligne juste après le commentaire de la boucle, l’opération a duré seulement 1/10 de seconde. Dans Word sur le Web (avec le serveur Edge en tant que navigateur), la synchronisation à l’intérieur de la boucle a duré 3 secondes, et seulement 6/10 à la synchronisation après la boucle, environ cinq fois plus vite. Dans un document avec 2000 instances de "The", il a fallu (dans Word sur le Web) 80 secondes avec la synchronisation à l’intérieur de la boucle et seulement 4 secondes avec la synchronisation après la boucle, environ 20 fois plus rapide.

> [!NOTE]
> Il est utile de demander si la version synchronisée à l’intérieur de la boucle s’exécute plus rapidement si les synchronisations ont été exécutées simultanément, ce qui peut être effectué en supprimant simplement le `await` mot clé à l’avant du `context.sync()` . Cela entraînerait le lancement de la synchronisation par le runtime, puis démarrera immédiatement l’itération suivante de la boucle sans attendre la fin de la synchronisation. Toutefois, cette solution n’est pas aussi intéressante que le fait `context.sync` de sortir entièrement de la boucle pour les raisons suivantes :
>
> - Tout comme les commandes d’un traitement par lots de synchronisation sont mises en file d’attente, les traitements par lots sont mis en file d’attente dans Office, mais Office ne prend pas en charge plus de 50 traitements par lots dans la file d’attente. Les autres déclenchent des erreurs. Par conséquent, si une boucle comporte plus de 50 itérations, il est possible que la taille de la file d’attente soit dépassée. Plus le nombre d’itérations est élevé, plus le risque de se produire est élevé. 
> - « Simultanément » ne signifie pas simultanément. Il faudra toujours plus de temps pour exécuter plusieurs opérations de synchronisation plutôt que d’en exécuter une.
> - Il n’est pas garanti que les opérations simultanées se terminent dans l’ordre dans lequel elles ont démarré. Dans l’exemple précédent, l’ordre dans lequel le mot « the » est mis en surbrillance n’a pas d’importance, mais il est important que les éléments de la collection soient traités dans l’ordre.

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a>Lecture de valeurs à partir du document avec le motif de boucle de fractionnement

`context.sync`Il est plus difficile d’éviter les s à l’intérieur d’une boucle lorsque le code doit *lire* une propriété des éléments de la collection à mesure qu’il traite chacun d’eux. Supposons que votre code doive itérer tous les contrôles de contenu d’un document Word et enregistrer le texte du premier paragraphe associé à chaque contrôle. Vos instincts de programmation peuvent vous amener à parcourir les contrôles, charger la `text` propriété de chaque paragraphe (premier), appeler `context.sync` pour remplir l’objet de paragraphe de proxy avec le texte du document, puis le consigner. Voici un exemple.

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load('items');
    await context.sync();

    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      await context.sync();
      console.log(paragraph.text);
    }
});
```

Dans ce scénario, pour éviter d’avoir une `context.sync` boucle, vous devez utiliser un modèle appelé modèle de **boucle de fractionnement** . Examinons un exemple concret du modèle avant d’obtenir une description formelle. Voici comment le modèle de boucle de fractionnement peut être appliqué à l’extrait de code précédent. Tenez compte des informations suivantes à propos de ce code :

- Il y a maintenant deux boucles et l' `context.sync` intervient entre elles, de sorte qu’il n’y a pas `context.sync` à l’intérieur de l’une ou l’autre boucle.
- La première boucle parcourt les éléments de l’objet de collection et charge la `text` propriété tout comme la boucle d’origine, mais la première boucle ne peut pas consigner le texte du paragraphe, car elle ne contient plus de `context.sync` pour remplir la `text` propriété de l' `paragraph` objet proxy. Au lieu de cela, il ajoute l' `paragraph` objet à un tableau.
- La deuxième boucle se répète dans le tableau qui a été créé par la première boucle et enregistre l' `text` `paragraph` élément. Cela est possible, car les `context.sync` deux boucles ont rempli toutes les `text` Propriétés.

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load("items");
    await context.sync();

    const firstParagraphsOfCCs = [];
    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      firstParagraphsOfCCs.push(paragraph);
    }

    await context.sync();

    for (let i = 0; i < firstParagraphsOfCCs.length; i++) {
      console.log(firstParagraphsOfCCs[i].text);
    }
});
```

L’exemple précédent suggère la procédure suivante pour activer une boucle qui contient un `context.sync` dans le motif de boucle de fractionnement : 

1. Remplacez la boucle par deux boucles.
2. Créer une première boucle pour effectuer une itération sur la collection et ajouter chaque élément à un tableau tout en chargeant également toute propriété de l’élément que votre code doit lire. 
3. À la suite de la première boucle, appelez `context.sync` pour remplir les objets proxy avec n’importe quelle propriété chargée. 
4. Suivez la `context.sync` boucle avec une seconde pour effectuer une itération sur le tableau créé dans la première boucle et lire les propriétés chargées.

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a>Traitement des objets dans le document avec le modèle d’objets corrélés

Examinons un scénario plus complexe où le traitement des éléments dans la collection nécessite des données qui ne se trouvent pas dans les éléments eux-mêmes. Le scénario prévisionne un complément Word qui opère sur des documents créés à partir d’un modèle avec du texte réutilisable. Éparpillés dans le texte sont une ou plusieurs instances des chaînes d’espace réservé suivantes : « {Coordinator} », « {adjoint} » et « {Manager} ». Le complément remplace chaque espace réservé par le nom de la personne. L’interface utilisateur du complément n’est pas importante dans cet article. Par exemple, il peut contenir un volet de tâches comportant trois zones de texte, chacune étiquetée avec l’un des espaces réservés. L’utilisateur entre un nom dans chaque zone de texte et appuie sur un bouton **remplacer** . Le gestionnaire du bouton crée un tableau qui mappe les noms aux espaces réservés, puis remplace chaque espace réservé par le nom attribué. 

Vous n’avez pas besoin de produire réellement un complément avec cette interface utilisateur pour tester le code. Vous pouvez utiliser l' [outil script Lab](../overview/explore-with-script-lab.md) pour prototyper le code important. Utilisez l’instruction d’affectation suivante pour créer le tableau de mappage.

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

Le code suivant montre comment remplacer chaque espace réservé par son nom attribué si vous avez utilisé `context.sync` à l’intérieur de boucles.

```javascript
Word.run(async (context) => {

    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');

      await context.sync(); 

      for (let j = 0; j < searchResults.items.length; j++) {
        searchResults.items[j].insertText(jobMapping[i].person, Word.InsertLocation.replace);

        await context.sync();
      }
    }
});
```

Dans le code précédent, il existe une boucle externe et une boucle interne. Chacun d’eux contient un `context.sync` . En fonction du tout premier extrait de code de cet article, vous verrez probablement que le `context.sync` dans la boucle interne peut simplement être déplacé après la boucle interne. Mais cela laisserait toujours le code avec un `context.sync` (deux d’entre eux) dans la boucle externe. Le code suivant montre comment vous pouvez supprimer `context.sync` des boucles. Nous abordons le code ci-dessous.

```javascript
Word.run(async (context) => {

    const allSearchResults = [];
    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');
      let correlatedSearchResult = {
        rangesMatchingJob: searchResults,
        personAssignedToJob: jobMapping[i].person
      }
      allSearchResults.push(correlatedSearchResult);
    }

    await context.sync()

    for (let i = 0; i < allSearchResults.length; i++) {
      let correlatedObject = allSearchResults[i];

      for (let j = 0; j < correlatedObject.rangesMatchingJob.items.length; j++) {
        let targetRange = correlatedObject.rangesMatchingJob.items[j];
        let name = correlatedObject.personAssignedToJob;
        targetRange.insertText(name, Word.InsertLocation.replace);
      }
    }

    await context.sync();
});
```

Remarque le code utilise le modèle de boucle de fractionnement :

- La boucle externe de l’exemple précédent a été divisée en deux. (La deuxième boucle possède une boucle interne, qui est attendue car le code se répète sur un ensemble de travaux (ou espaces réservés) et, dans cette définition, elle se répète sur les plages correspondantes.)
- Il y a une `context.sync` boucle après chaque boucle principale, mais pas `context.sync` à l’intérieur d’une boucle.
- La deuxième boucle majeure effectue une itération dans un tableau créé dans la première boucle.

Toutefois, le tableau créé dans la première boucle ne contient *pas* uniquement un objet Office comme première boucle dans la section [Reading values from the document with the Split Loop pattern](#reading-values-from-the-document-with-the-split-loop-pattern). Cela est dû au fait que certaines des informations nécessaires au traitement des objets de plage Word ne se trouvent pas dans les objets Range eux-mêmes, mais qu’elles proviennent du `jobMapping` tableau.

Par conséquent, les objets dans le tableau créé dans la première boucle sont des objets personnalisés ayant deux propriétés. Le premier est un tableau de plages de mots qui correspondent à une fonction spécifique (c’est-à-dire une chaîne d’espace réservé) et le deuxième est une chaîne qui fournit le nom de la personne affectée au travail. Cela rend la boucle finale facile à écrire et facile à lire, car toutes les informations nécessaires au traitement d’une plage donnée sont contenues dans le même objet personnalisé qui contient la plage. Le nom qui doit remplacer _ **correlatedObject**. rangesMatchingJob. Items [j]_ est l’autre propriété du même objet : _ **correlatedObject**. personAssignedToJob_.

Nous appelons cette variante du modèle d' **objets corrélé** . L’idée générale est que la première boucle crée un tableau d’objets personnalisés. Chaque objet possède une propriété dont la valeur est l’un des éléments d’un objet de collection Office (ou un tableau de ces éléments). L’objet personnalisé possède d’autres propriétés, chacune fournissant les informations nécessaires pour traiter les objets Office dans la boucle finale. Voir la section [autres exemples de ces modèles](#other-examples-of-these-patterns) pour un lien vers un exemple dans lequel l’objet de corrélation personnalisé comporte plus de deux propriétés.

Une autre restriction : parfois, il faut plus d’une boucle pour créer le tableau des objets corrélés personnalisés. Cela peut se produire si vous avez besoin de lire une propriété de chaque membre d’un objet de collection Office uniquement pour collecter des informations qui seront utilisées pour traiter un autre objet de collection. (Par exemple, votre code doit lire les titres de toutes les colonnes d’un tableau Excel, car votre complément va appliquer un format numérique aux cellules de certaines colonnes en fonction du titre de cette colonne.) Toutefois, vous pouvez toujours conserver les `context.sync` s entre les boucles, plutôt que dans une boucle. Consultez la section [autres exemples de ces modèles](#other-examples-of-these-patterns) pour obtenir un exemple.

## <a name="other-examples-of-these-patterns"></a>Autres exemples de ces modèles

- Pour un exemple très simple pour Excel qui utilise des `Array.forEach` boucles, consultez la question relative à la réponse acceptée à ce débordement de pile : [est-il possible de prendre en file d’attente plusieurs Context. Load avant Context. Sync ?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- Pour un exemple simple pour Word qui utilise des `Array.forEach` boucles et n’utilise pas de `async` / `await` syntaxe, voir la réponse acceptée à cette question de dépassement de pile : [itération sur tous les paragraphes avec des contrôles de contenu avec l’API JavaScript pour Office](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).
- Pour obtenir un exemple pour Word écrit en écriture manuscrite, consultez l’exemple de [Vérificateur de style Angular2 de complément Word](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), en particulier le fichier [word.document. service. TS](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts). Elle comporte une combinaison de `for` et de `Array.forEach` boucles.
- Pour un exemple de mot avancé, [importez-le dans](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) l' [outil script Lab](../overview/explore-with-script-lab.md). Pour le contexte dans l’utilisation du fichier d’aide à la pile, consultez la réponse acceptée sur le document de la question de débordement de pile [non synchronisé après le texte de remplacement](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text). Cet exemple crée un type d’objet corrélé personnalisé qui a trois propriétés. Il utilise un total de trois boucles pour construire le tableau des objets corrélés, et deux boucles supplémentaires pour effectuer le traitement final. Il existe un mélange de `for` boucles et de `Array.forEach` boucles.
- Bien que ce ne soit pas seulement un exemple des modèles de boucle de fractionnement ou d’objets corrélés, il existe un exemple Excel avancé qui montre comment convertir un ensemble de valeurs de cellule en d’autres devises avec un seul `context.sync` . Pour l’essayer, ouvrez l' [outil script Lab](../overview/explore-with-script-lab.md) et accédez à l’exemple **convertisseur de devise** .

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>Quand *devez-vous utiliser* les modèles de cet article ?

Excel ne peut pas lire plus de 5 Mo de données dans un appel de `context.sync` . Si cette limite est dépassée, une erreur est générée. (Pour plus d’informations, reportez-vous à la section « compléments Excel » de [limitations de ressources et d’optimisation des performances pour les compléments Office](resource-limits-and-performance-optimization.md#excel-add-ins) .) Il est très rare que cette limite soit proche, mais si cela peut se produire avec votre complément, votre code *ne doit pas* charger toutes les données dans une seule boucle et suivre la boucle avec un `context.sync` . Toutefois, vous devez toujours éviter d’avoir une `context.sync` boucle dans chaque itération d’une boucle sur un objet de collection. Au lieu de cela, définissez des sous-ensembles des éléments de la collection et faites une boucle sur chaque sous-ensemble, avec un `context.sync` entre les boucles. Vous pouvez la structurer avec une boucle externe qui itère sur les sous-ensembles et contient le `context.sync` dans chacune de ces itérations externes.
