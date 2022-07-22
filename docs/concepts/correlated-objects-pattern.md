---
title: Éviter d’utiliser la méthode context.sync dans des boucles
description: Découvrez comment utiliser la boucle fractionnée et les modèles d’objets corrélés pour éviter d’appeler context.sync dans une boucle.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6b0239e05a597949160afbb2604143f3d6626462
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958698"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>Éviter d’utiliser la méthode context.sync dans des boucles

> [!NOTE]
> Cet article part du principe que vous n’êtes pas au début de l’utilisation d’au moins l’une des quatre API&mdash;JavaScript Office spécifiques à l’application pour Excel, Word, OneNote et Visio&mdash;qui utilisent un système de traitement par lots pour interagir avec le document Office. En particulier, vous devez savoir ce qu’un appel `context.sync` fait et vous devez savoir ce qu’est un objet de collection. Si vous n’en êtes pas à ce stade, commencez par [comprendre l’API JavaScript Office](../develop/understanding-the-javascript-api-for-office.md) et la documentation liée sous « application spécifique » dans cet article.

Pour certains scénarios de programmation dans les compléments Office qui utilisent l’un des modèles d’API spécifiques à l’application (pour Excel, Word, OneNote et Visio), votre code doit lire, écrire ou traiter une propriété de chaque membre d’un objet de collection. Par exemple, un complément Excel qui doit obtenir les valeurs de chaque cellule d’une colonne de tableau particulière ou d’un complément Word qui doit mettre en surbrillance chaque instance d’une chaîne dans le document. Vous devez itérer sur les membres de la `items` propriété de l’objet collection, mais, pour des raisons de performances, vous devez éviter d’appeler `context.sync` chaque itération de la boucle. Chaque appel est `context.sync` un aller-retour entre le complément et le document Office. Les allers-retours répétés nuisent aux performances, en particulier si le complément est en cours d’exécution dans Office sur le Web parce que les allers-retours vont sur Internet.

> [!NOTE]
> Tous les exemples de cet article utilisent `for` des boucles, mais les pratiques décrites s’appliquent à toute instruction de boucle qui peut itérer dans un tableau, y compris les suivantes :
>
> - `for`
> - `for of`
> - `while`
> - `do while`
>
> Ils s’appliquent également à toute méthode de tableau à laquelle une fonction est passée et appliquée aux éléments du tableau, y compris les éléments suivants :
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

Dans le cas le plus simple, vous écrivez uniquement aux membres d’un objet de collection, sans lire leurs propriétés. Par exemple, le code suivant met en évidence en jaune chaque instance de « the » dans un document Word.

> [!NOTE]
> Il est généralement recommandé d’avoir une finale `context.sync` juste avant le caractère « } » fermant de la fonction d’application `run` (par `Excel.run`exemple, , `Word.run`etc.). Cela est dû au fait que la `run` fonction effectue un appel masqué comme `context.sync` dernière chose qu’elle fait si, et seulement si, il existe des commandes en file d’attente qui n’ont pas encore été synchronisées. Le fait que cet appel soit masqué peut prêter à confusion, donc nous vous recommandons généralement d’ajouter l’explicite `context.sync`. Cependant, étant donné que cet article est sur la réduction des appels de `context.sync`, il est en fait plus déroutant d’ajouter une finale `context.sync`entièrement inutile . Ainsi, dans cet article, nous laissons de l’extérieur quand il n’y a pas de commandes non synchronisées à la fin de la `run`.

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

  for (let i = 0; i < searchResults.items.length; i++) {
    searchResults.items[i].font.highlightColor = '#FFFF00';

    await context.sync(); // SYNCHRONIZE IN EACH ITERATION
  }
  
  // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

  // Record the system time again then calculate how long the operation took.
  endTime = performance.now();
  console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
})
```

Le code précédent a pris 1 seconde complète pour se terminer dans un document avec 200 instances de « the » dans Word sur Windows. Toutefois, lorsque la `await context.sync();` ligne à l’intérieur de la boucle est commentée et que la même ligne juste après la boucle est décompressée, l’opération n’a pris qu’un 1/10e de seconde. Dans Word sur le web (avec Edge comme navigateur), il a fallu 3 secondes complètes avec la synchronisation à l’intérieur de la boucle et seulement 6/10èmes de seconde avec la synchronisation après la boucle, environ cinq fois plus rapide. Dans un document contenant 2 000 instances de « the », il a fallu (dans Word sur le web) 80 secondes avec la synchronisation à l’intérieur de la boucle et seulement 4 secondes avec la synchronisation après la boucle, environ 20 fois plus rapide.

> [!NOTE]
> Il est intéressant de se demander si la version de synchronisation à l’intérieur de la boucle s’exécuterait plus rapidement si les synchronisations s’exécutaient simultanément, ce qui peut être fait en supprimant simplement le `await` mot clé du début du `context.sync()`. Le runtime lance alors la synchronisation, puis démarre immédiatement l’itération suivante de la boucle sans attendre la fin de la synchronisation. Toutefois, il ne s’agit pas d’une solution aussi bonne que le déplacement de la `context.sync` boucle entièrement pour ces raisons.
>
> - Tout comme les commandes d’un travail de synchronisation par lots sont mises en file d’attente, les travaux par lots eux-mêmes sont mis en file d’attente dans Office, mais Office ne prend pas en charge plus de 50 travaux par lots dans la file d’attente. D’autres déclencheurs d’erreurs. Par conséquent, s’il y a plus de 50 itérations dans une boucle, il est possible que la taille de la file d’attente soit dépassée. Plus le nombre d’itérations est élevé, plus la probabilité de ce problème est grande.
> - « Simultanément » ne signifie pas simultanément. Il faudrait encore plus de temps pour exécuter plusieurs opérations de synchronisation que pour en exécuter une.
> - Il n’est pas garanti que les opérations simultanées se terminent dans l’ordre dans lequel elles ont démarré. Dans l’exemple précédent, peu importe l’ordre dans lequel le mot « the » est mis en surbrillance, mais il existe des scénarios où il est important que les éléments de la collection soient traités dans l’ordre.

## <a name="read-values-from-the-document-with-the-split-loop-pattern"></a>Lire les valeurs du document avec le modèle de boucle fractionnée

L’évitement `context.sync`de s à l’intérieur d’une boucle devient plus difficile lorsque le code doit *lire* une propriété des éléments de collection au fur et à mesure qu’il traite chacun d’eux. Supposons que votre code doit itérer tous les contrôles de contenu dans un document Word et consigner le texte du premier paragraphe associé à chaque contrôle. Vos instincts de programmation peuvent vous amener à effectuer une boucle sur les contrôles, charger la `text` propriété de chaque (premier) paragraphe, appeler `context.sync` pour remplir l’objet de paragraphe proxy avec le texte du document, puis le journaliser. Voici un exemple.

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

Dans ce scénario, pour éviter d’avoir une `context.sync` boucle dans une boucle, vous devez utiliser un modèle que nous appelons le modèle de **boucle fractionnée** . Voyons un exemple concret du modèle avant d’en obtenir une description formelle. Voici comment appliquer le modèle de boucle fractionnée à l’extrait de code précédent. Notez ce qui suit à propos de ce code.

- Il y a maintenant deux boucles et le `context.sync` vient entre elles, il n’y a donc pas `context.sync` à l’intérieur de l’une ou l’autre boucle.
- La première boucle itère dans les éléments de l’objet de collection et charge la `text` propriété comme la boucle d’origine, mais la première boucle ne peut pas journaliser le texte du paragraphe, car elle ne contient plus un `context.sync` objet pour remplir la `text` propriété de l’objet `paragraph` proxy. Au lieu de cela, il ajoute l’objet `paragraph` à un tableau.
- La deuxième boucle itère dans le tableau créé par la première boucle et consigne chaque `text` `paragraph` élément. Cela est possible, car le `context.sync` contenu entre les deux boucles a rempli toutes les `text` propriétés.

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

L’exemple précédent suggère la procédure suivante pour transformer une boucle qui contient une `context.sync` boucle dans le modèle de boucle fractionnée.

1. Remplacez la boucle par deux boucles.
2. Créez une première boucle pour itérer sur la collection et ajouter chaque élément à un tableau tout en chargeant toutes les propriétés de l’élément que votre code doit lire.
3. Après la première boucle, appelez `context.sync` pour remplir les objets proxy avec toutes les propriétés chargées.
4. Suivez la `context.sync` deuxième boucle pour itérer sur le tableau créé dans la première boucle et lire les propriétés chargées.

## <a name="process-objects-in-the-document-with-the-correlated-objects-pattern"></a>Traiter des objets dans le document avec le modèle d’objets corrélés

Prenons un scénario plus complexe où le traitement des éléments de la collection nécessite des données qui ne se trouvent pas dans les éléments eux-mêmes. Le scénario envisage un complément Word qui fonctionne sur des documents créés à partir d’un modèle avec du texte réutilisable. Le texte contient une ou plusieurs instances des chaînes d’espace réservé suivantes : « {Coordinator} », « {Deputy} » et « {Manager} ». Le complément remplace chaque espace réservé par le nom d’une personne. L’interface utilisateur du complément n’est pas importante pour cet article. Par exemple, il peut avoir un volet Office avec trois zones de texte, chacune étiquetée avec l’un des espaces réservés. L’utilisateur entre un nom dans chaque zone de texte, puis appuie sur un bouton **Remplacer** . Le gestionnaire du bouton crée un tableau qui mappe les noms aux espaces réservés, puis remplace chaque espace réservé par le nom attribué.

Vous n’avez pas besoin de produire un complément avec cette interface utilisateur pour expérimenter le code. Vous pouvez utiliser [l’outil Script Lab](../overview/explore-with-script-lab.md) pour prototyper le code important. Utilisez l’instruction d’affectation suivante pour créer le tableau de mappages.

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

Le code suivant montre comment remplacer chaque espace réservé par son nom attribué si vous l’avez utilisé `context.sync` à l’intérieur des boucles.

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

Dans le code précédent, il existe une boucle externe et une boucle interne. Chacun d’eux contient un `context.sync`. En fonction du tout premier extrait de code de cet article, vous voyez probablement que la `context.sync` boucle interne peut simplement être déplacée après la boucle interne. Mais cela laisserait toujours le code avec un `context.sync` (deux d’entre eux en fait) dans la boucle externe. Le code suivant montre comment supprimer `context.sync` des boucles. Nous abordons le code ci-dessous.

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

Notez que le code utilise le modèle de boucle fractionnée.

- La boucle externe de l’exemple précédent a été divisée en deux. (La deuxième boucle a une boucle interne, ce qui est attendu, car le code itère sur un ensemble de travaux (ou d’espaces réservés) et dans cet ensemble, il itère sur les plages correspondantes.)
- Il y a une `context.sync` boucle après chaque boucle principale, mais aucune boucle n’est `context.sync` à l’intérieur d’une boucle.
- La deuxième boucle majeure itère dans un tableau créé dans la première boucle.

Toutefois, le tableau créé dans la première boucle ne contient *pas* seulement un objet Office, comme la première boucle l’a fait dans la section [Lecture des valeurs du document avec le modèle de boucle fractionnée](#read-values-from-the-document-with-the-split-loop-pattern). Cela est dû au fait que certaines des informations nécessaires pour traiter les objets de plage Word ne sont pas dans les objets Range eux-mêmes, mais proviennent plutôt du `jobMapping` tableau.

Ainsi, les objets du tableau créé dans la première boucle sont des objets personnalisés qui ont deux propriétés. Le premier est un tableau de plages de mots qui correspondent à un titre de travail spécifique (autrement dit, une chaîne d’espace réservé) et le second est une chaîne qui fournit le nom de la personne affectée au travail. Cela facilite l’écriture et la lecture de la boucle finale, car toutes les informations nécessaires au traitement d’une plage donnée sont contenues dans le même objet personnalisé que celui qui contient la plage. Le nom qui doit remplacer _correlatedObject.rangesMatchingJob.items[j]_ est l’autre propriété du même objet : _**correlatedObject.personAssignedToJob**_.

Nous appelons cette variante du modèle de boucle fractionnée le modèle **d’objets corrélés** . L’idée générale est que la première boucle crée un tableau d’objets personnalisés. Chaque objet a une propriété dont la valeur est l’un des éléments d’un objet de collection Office (ou un tableau de ces éléments). L’objet personnalisé a d’autres propriétés, chacune fournissant les informations nécessaires pour traiter les objets Office dans la boucle finale. Consultez la section [Autres exemples de ces modèles](#other-examples-of-these-patterns) pour obtenir un lien vers un exemple où l’objet de corrélation personnalisé a plus de deux propriétés.

Une autre mise en garde : il faut parfois plusieurs boucles uniquement pour créer le tableau d’objets de corrélation personnalisés. Cela peut se produire si vous devez lire une propriété de chaque membre d’un objet de collection Office uniquement pour collecter des informations qui seront utilisées pour traiter un autre objet de collection. (Par exemple, votre code doit lire les titres de toutes les colonnes d’un tableau Excel, car votre complément va appliquer un format numérique aux cellules de certaines colonnes en fonction du titre de cette colonne.) Mais vous pouvez toujours conserver les `context.sync`s entre les boucles, plutôt que dans une boucle. Pour obtenir un exemple, consultez la section [Autres exemples de ces modèles](#other-examples-of-these-patterns) .

## <a name="other-examples-of-these-patterns"></a>Autres exemples de ces modèles

- Pour obtenir un exemple très simple pour Excel qui utilise des boucles `Array.forEach` , consultez la réponse acceptée à cette question Stack Overflow : [Est-il possible de mettre en file d’attente plusieurs context.load avant context.sync ?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- Pour obtenir un exemple simple pour Word qui utilise des boucles `Array.forEach` et n’utilise `async`/`await` pas de syntaxe, consultez la réponse acceptée à cette question Stack Overflow : [Itération sur tous les paragraphes avec des contrôles de contenu avec l’API JavaScript Office](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).
- Pour obtenir un exemple de Word écrit en TypeScript, consultez l’exemple de vérificateur de [style Angular2 du complément Word](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), en particulier le fichier [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts). Il a un mélange de boucles et `Array.forEach` de `for` boucles.
- Pour un exemple Word avancé, importez [ce gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) dans [l’outil Script Lab](../overview/explore-with-script-lab.md). Pour connaître le contexte dans l’utilisation du gist, consultez la réponse acceptée au document de question Stack Overflow [qui n’est pas synchronisé après le remplacement du texte](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text). Cet exemple crée un type d’objet de corrélation personnalisé qui a trois propriétés. Il utilise un total de trois boucles pour construire le tableau d’objets corrélés, et deux autres boucles pour effectuer le traitement final. Il y a un mélange de boucles et `Array.forEach` de `for` boucles.
- Bien qu’il ne s’agit pas strictement d’un exemple de boucle fractionnée ou de modèles d’objets corrélés, il existe un exemple Excel avancé qui montre comment convertir un ensemble de valeurs de cellule en d’autres devises avec une seule .`context.sync` Pour l’essayer, ouvrez [l’outil Script Lab](../overview/explore-with-script-lab.md) et accédez à l’exemple **currency converter**.

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>Quand *ne devez-vous pas* utiliser les modèles de cet article ?

Excel ne peut pas lire plus de 5 Mo de données dans un appel donné de `context.sync`. Si cette limite est dépassée, une erreur est levée. (Pour plus d’informations, consultez la section « Compléments Excel » des [limites des ressources et de l’optimisation des performances pour les compléments Office](resource-limits-and-performance-optimization.md#excel-add-ins) .) Il est très rare que cette limite soit atteinte, mais s’il y a une chance que cela se produise avec votre complément, votre code *ne doit pas* charger toutes les données dans une seule boucle et suivre la boucle avec un `context.sync`. Toutefois, vous devez toujours éviter d’avoir une `context.sync` itération dans chaque itération d’une boucle sur un objet de collection. Au lieu de cela, définissez des sous-ensembles des éléments de la collection et effectuez une boucle sur chaque sous-ensemble à tour de rôle, avec un `context.sync` entre les boucles. Vous pouvez structurer cette opération avec une boucle externe qui itère sur les sous-ensembles et contient le `context.sync` contenu de chacune de ces itérations externes.
