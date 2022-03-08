---
title: Éviter d’utiliser la méthode context.sync dans des boucles
description: Découvrez comment utiliser la boucle fractionée et les modèles d’objets corrélés pour éviter d’appeler context.sync dans une boucle.
ms.date: 02/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: 735d82ce276b3ba6c6afe8d50229beaf55829884
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340322"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>Éviter d’utiliser la méthode context.sync dans des boucles

> [!NOTE]
> Cet article suppose que vous n’êtes pas au début de l’utilisation d’au moins l’une des quatre API JavaScript&mdash; Office spécifiques à l’application pour Excel, Word, OneNote et Visio&mdash; si utiliser un système de traitement par lots pour interagir avec le document Office. En particulier, vous devez savoir de quoi fait `context.sync` un appel et ce qu’est un objet de collection. Si vous n’en êtes pas à ce stade, commencez par comprendre [l’API JavaScript Office](../develop/understanding-the-javascript-api-for-office.md) et la documentation liée sous « spécifique à l’application » dans cet article.

Pour certains scénarios de programmation dans les applications Office qui utilisent l’un des modèles d’API propres à l’application (pour Excel, Word, OneNote et Visio), votre code doit lire, écrire ou traiter certaines propriétés de chaque membre d’un objet de collection. Par exemple, un Excel qui a besoin d’obtenir les valeurs de chaque cellule d’une colonne de tableau particulière ou d’un add-in Word qui doit mettre en surbrillement chaque instance d’une chaîne dans le document. Vous devez itérer `items` sur les membres de la propriété de l’objet de collection, mais, pour des raisons de performances, `context.sync` vous devez éviter d’appeler chaque itération de la boucle. Chaque appel est `context.sync` un aller-retour entre le add-in et le Office document. Les allers-retours répétés nuit aux performances, en particulier si le Office sur le Web est en cours d’exécution, car les allers-retours traversent Internet.

> [!NOTE]
> Tous les exemples de cet article `for` utilisent des boucles, mais les pratiques décrites s’appliquent à n’importe quelle instruction de boucle qui peut itérer dans un tableau, notamment les suivantes :
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> Elles s’appliquent également à toute méthode de tableau à laquelle une fonction est passée et appliquée aux éléments du tableau, notamment :
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

Dans le cas le plus simple, vous écrivez uniquement aux membres d’un objet de collection, et non à la lecture de leurs propriétés. Par exemple, le code suivant met en sur évidence en jaune chaque instance de « the » dans un document Word.

> [!NOTE]
> Il est généralement pratique `context.sync` de placer une finale juste avant le caractère « } » de fermeture de la méthode d’application `run` ( `Excel.run`par exemple, , `Word.run`etc.). Cela est dû au `run` fait que la méthode effectue un appel masqué en tant que dernière chose qu’elle fait si, et uniquement si, il existe des commandes en file d’attente qui n’ont `context.sync` pas encore été synchronisées. Le fait que cet appel soit masqué peut prêter à confusion, c’est pourquoi nous vous recommandons généralement d’ajouter l’appel explicite `context.sync`. Toutefois, étant donné que cet article est sur la réduction des appels de `context.sync`, il est en fait plus déroutant d’ajouter une finale entièrement inutile `context.sync`. Ainsi, dans cet article, nous ne le faisons pas lorsqu’il n’y a aucune commande nonynchronisée à la fin du `run`.

```javascript
await Word.run(async function (context) {
  let startTime, endTime;
  const docBody = context.document.body;

  // search() returns an array of Ranges.
  const searchResults = docBody.search('the', { matchWholeWord: true });
  searchResults.load('font');
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
```

Le code précédent a pris 1 seconde complète pour se terminer dans un document avec 200 instances de « the » dans Word sur Windows. `await context.sync();` Toutefois, lorsque la ligne à l’intérieur de la boucle est commentée et que la même ligne est décompressée, l’opération n’a pris qu’un 1/10e de seconde. Dans Word sur le web (avec Edge comme navigateur), la synchronisation à l’intérieur de la boucle a pris 3 secondes complètes et seulement 6/10e de seconde avec la synchronisation après la boucle, environ cinq fois plus rapidement. Dans un document avec 2 000 instances de « the », il a fallu (en Word sur le web) 80 secondes avec la synchronisation à l’intérieur de la boucle et seulement 4 secondes avec la synchronisation après la boucle, environ 20 fois plus rapide.

> [!NOTE]
> Il est intéressant de se demander si la version synchronisée à l’intérieur de la boucle s’exécuterait plus rapidement si les synchronisations s’exécutaient simultanément, `await` `context.sync()`ce qui pourrait être fait en supprimant simplement le mot clé à l’avant de la boucle. Cela entraînerait le démarrage de la synchronisation par le runtime, puis le démarrage immédiat de l’itération suivante de la boucle sans attendre la fin de la synchronisation. Toutefois, il ne s’agit pas d’une solution `context.sync` aussi adaptée que de sortir complètement de la boucle pour ces raisons.
>
> - Tout comme les commandes d’un travail de traitement par lots de synchronisation sont en file d’attente, les travaux de lots eux-mêmes sont mis en file d’attente dans Office, mais Office ne prend pas en charge plus de 50 travaux par lots dans la file d’attente. Toute autre erreur se déclenche. Ainsi, s’il y a plus de 50 itérations dans une boucle, il est possible que la taille de la file d’attente soit dépassée. Plus le nombre d’itérations est élevé, plus le risque est grand. 
> - « Simultanément » ne signifie pas simultanément. L’exécution de plusieurs opérations de synchronisation peut encore prendre plus de temps que d’en exécuter une.
> - Il n’est pas garanti que les opérations simultanées se terminent dans l’ordre dans lequel elles ont démarré. Dans l’exemple précédent, peu importe l’ordre dans lequel le mot « the » est mis en surbrillant, mais dans certains cas, il est important que les éléments de la collection soient traitées dans l’ordre.

## <a name="read-values-from-the-document-with-the-split-loop-pattern"></a>Lire les valeurs du document avec le modèle de boucle fractionner

L’évitement `context.sync`dans une boucle devient plus difficile lorsque le code doit *lire* une propriété des éléments de la collection lorsqu’il traite chacun d’eux. Supposons que votre code doit itérer tous les contrôles de contenu dans un document Word et consigner le texte du premier paragraphe associé à chaque contrôle. Vos programmes peuvent vous amener à faire une boucle sur les contrôles, `text` à charger la propriété de chaque (premier) paragraphe, `context.sync` à appeler pour remplir l’objet de paragraphe proxy avec le texte du document, puis à le journaliser. Voici un exemple.

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

Dans ce scénario, pour éviter d’avoir une `context.sync` boucle en boucle, vous devez utiliser un modèle que nous appelons le **modèle de boucle fractionner** . Voyons un exemple concret du modèle avant d’en obtenir une description formelle. Voici comment le modèle de boucle fractionner peut être appliqué à l’extrait de code précédent. Notez ce qui suit à propos de ce code.

- Il existe maintenant deux boucles qui s’entrent `context.sync` entre elles, il n’y a donc aucune `context.sync` boucle à l’intérieur de l’une ou l’autre.
- La première boucle par itérera les éléments de l’objet de collection `text` et charge la propriété de la même façon que la boucle d’origine, mais la première boucle ne peut pas journaliser le texte du paragraphe, car elle ne contient plus de `context.sync` `text` `paragraph` valeur pour remplir la propriété de l’objet proxy. Au lieu de cela, il ajoute l’objet `paragraph` à un tableau.
- La seconde boucle par itérera dans le tableau créé par la première boucle et `text` enregistre le journal de chaque `paragraph` élément. Cela est possible car la `context.sync` boucle entre les deux boucles remplit toutes les propriétés `text` .

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

L’exemple précédent suggère la procédure suivante pour transformer une boucle qui contient une `context.sync` boucle en modèle de boucle fractionner.

1. Remplacez la boucle par deux boucles.
2. Créez une première boucle pour itérer sur la collection et ajoutez chaque élément à un tableau tout en chargeant n’importe quelle propriété de l’élément que votre code doit lire.
3. Après la première boucle, appelez `context.sync` pour remplir les objets proxy avec les propriétés chargées.
4. Suivez la `context.sync` deuxième boucle pour itérer sur le tableau créé dans la première boucle et lire les propriétés chargées.

## <a name="process-objects-in-the-document-with-the-correlated-objects-pattern"></a>Traiter les objets du document avec le modèle d’objets corrélés

Considérons un scénario plus complexe dans lequel le traitement des éléments de la collection nécessite des données qui ne se trouvent pas dans les éléments eux-mêmes. Le scénario envisage un add-in Word qui fonctionne sur des documents créés à partir d’un modèle avec du texte réutilisable. Une ou plusieurs instances des chaînes d’espaces réservé suivantes sont dispersées dans le texte : « {Coordinator} », « {Coordinator} » et « {Manager} ». Le add-in remplace chaque espace réservé par le nom d’une personne. L’interface utilisateur du add-in n’est pas importante pour cet article. Par exemple, il peut avoir un volet Des tâches avec trois zones de texte, chacune étiquetée avec l’un des espaces réservé. L’utilisateur entre un nom dans chaque zone de texte, puis appuie sur un **bouton** Remplacer. Le responsable du bouton crée un tableau qui met les noms sur les espaces réservé, puis remplace chaque espace réservé par le nom attribué.

Vous n’avez pas besoin de produire réellement un add-in avec cette interface utilisateur pour expérimenter le code. Vous pouvez utiliser [l’outil Script Lab pour](../overview/explore-with-script-lab.md) prototyper le code important. Utilisez l’instruction d’affectation suivante pour créer le tableau de mappage.

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

Le code suivant montre comment vous pouvez remplacer chaque espace réservé par son nom attribué si vous avez utilisé des `context.sync` boucles intérieures.

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

Dans le code précédent, il existe une boucle externe et une boucle interne. Chacun d’eux contient un `context.sync`. Selon le tout premier extrait de code de cet article, `context.sync` vous verrez probablement que la boucle interne peut simplement être déplacée après la boucle interne. Toutefois, cela laisserait le code avec un `context.sync` (deux d’entre eux en fait) dans la boucle externe. Le code suivant montre comment supprimer des `context.sync` boucles. Nous abordons le code ci-dessous.

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

Notez que le code utilise le modèle de boucle fractionner.

- La boucle externe de l’exemple précédent a été divisée en deux. (La deuxième boucle possède une boucle interne, ce qui est attendu car le code se itère sur un ensemble de travaux (ou espaces réservé) et dans ce jeu, il itère sur les plages correspondantes.)
- Il existe une boucle après `context.sync` chaque boucle principale, mais pas `context.sync` à l’intérieur d’une boucle.
- La deuxième boucle principale par itérera dans un tableau créé dans la première boucle.

Toutefois, le tableau créé dans la première  boucle ne contient pas seulement un objet Office comme la première boucle l’a fait dans la section Valeurs de lecture du document avec le modèle de boucle [fractionner](#read-values-from-the-document-with-the-split-loop-pattern). Cela est dû au fait que certaines des informations nécessaires au traitement des objets de plage word ne se trouvent pas dans les objets Range eux-mêmes, mais proviennent plutôt du `jobMapping` tableau.

Ainsi, les objets du tableau créé dans la première boucle sont des objets personnalisés qui ont deux propriétés. La première est un tableau de plages de mots qui correspondent à une fonction spécifique (c’est-à-dire, une chaîne d’espace réservé) et la seconde est une chaîne qui fournit le nom de la personne affectée au travail. Cela facilite l’écriture et la lecture de la boucle finale, car toutes les informations nécessaires au traitement d’une plage donnée sont contenues dans le même objet personnalisé qui contient la plage. Le nom qui doit remplacer _correlatedObject.rangesMatchingJob.items[j]_ est l’autre propriété du même objet : _**correlatedObject.personAssignedToJob**_.

Nous appelons cette variante du modèle de boucle fractionée le **modèle d’objets corrélés** . L’idée générale est que la première boucle crée un tableau d’objets personnalisés. Chaque objet possède une propriété dont la valeur est l’un des éléments d’un Office collection d’objets (ou d’un tableau de ces éléments). L’objet personnalisé possède d’autres propriétés, chacune d’elles fournit les informations nécessaires pour traiter Office objets dans la boucle finale. Consultez la section [Autres exemples de ces modèles](#other-examples-of-these-patterns) pour obtenir un lien vers un exemple où l’objet de corrélation personnalisé possède plus de deux propriétés.

Autre mise en garde : il faut parfois plusieurs boucles simplement pour créer le tableau d’objets de mise en corrélation personnalisés. Cela peut se produire si vous devez lire une propriété de chaque membre d’un objet de collection Office uniquement pour collecter des informations qui seront utilisées pour traiter un autre objet de collection. (Par exemple, votre code doit lire les titres de toutes les colonnes d’un tableau Excel, car votre application va appliquer un format numérique aux cellules de certaines colonnes en fonction du titre de cette colonne.) Mais vous pouvez toujours conserver les `context.sync`s entre les boucles, plutôt que dans une boucle. Voir la section [Autres exemples de ces modèles](#other-examples-of-these-patterns) pour obtenir un exemple.

## <a name="other-examples-of-these-patterns"></a>Autres exemples de ces modèles

- Pour obtenir un exemple très simple pour Excel `Array.forEach` qui utilise des boucles, consultez la réponse acceptée à cette question stack overflow : est-il possible de mettre en file d’attente plusieurs [context.load avant context.sync ?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- Pour obtenir un exemple simple pour Word `Array.forEach` qui utilise des boucles et n’utilise`await` `async`/pas de syntaxe, voir la réponse acceptée à cette question stack overflow : Itérant sur tous les paragraphes avec des contrôles de contenu avec Office [API JavaScript](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).
- Pour obtenir un exemple pour Word écrit en TypeScript, voir l’exemple de contrôle de style du [add-in Word Angular2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), en particulier le fichier [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts). Il possède une combinaison de boucles `for` `Array.forEach` et de boucles.
- Pour obtenir un exemple Word avancé, [importez ce gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) dans l [Script Lab’outil](../overview/explore-with-script-lab.md). Pour obtenir le contexte de l’utilisation du gist, voir la réponse acceptée à la question Stack Overflow [Document non synchronisée après le remplacement du texte](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text). Cet exemple crée un type d’objet de corrélation personnalisé qui possède trois propriétés. Il utilise un total de trois boucles pour construire le tableau d’objets corrélés, et deux autres boucles pour réaliser le traitement final. Il existe une combinaison de boucles `for` `Array.forEach` et de boucles.
- Bien qu’il ne s’agit pas strictement d’un exemple de boucle fractionner ou de modèles d’objets corrélés, il existe un exemple de Excel `context.sync`avancé qui montre comment convertir un ensemble de valeurs de cellule en d’autres devises avec une seule . Pour l’essayer, ouvrez [Script Lab’outil et](../overview/explore-with-script-lab.md) accédez à l’exemple **de convertisseur de** devise.

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>Quand ne *devez-vous* pas utiliser les modèles de cet article ?

Excel ne peut pas lire plus de 5 Mo de données dans un appel donné de `context.sync`. Si cette limite est dépassée, une erreur est lancée. (Pour plus d’informations, voir la section Excel sur les [limites](resource-limits-and-performance-optimization.md#excel-add-ins) de ressources et l’optimisation des performances pour les Office de recherche.) Il est très rare que cette limite soit approche, mais s’il est possible que cela se produise avec votre add-in, votre code  `context.sync`ne doit pas charger toutes les données en une seule boucle et suivre la boucle avec un . Toutefois, évitez d’avoir une boucle `context.sync` dans chaque itération d’un objet de collection. Définissez plutôt les sous-ensembles des éléments de la collection et bouclez sur chaque sous-ensemble à tour de tour, `context.sync` avec un entre les boucles. Vous pouvez structurer cela avec une boucle externe qui s’itération sur les sous-ensembles `context.sync` et contient les dans chacune de ces itérations extérieures.
