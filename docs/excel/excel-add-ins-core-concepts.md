---
title: Concepts fondamentaux de programmation avec l’API JavaScript pour Excel
description: Utilisez l’API JavaScript pour Excel afin de créer des compléments pour Excel.
ms.date: 11/29/2018
ms.openlocfilehash: c92006fe8b1b51daa7cac420c63c5926a2414b61
ms.sourcegitcommit: 026437bd3819f4e9cd4153ebe60c98ab04e18f4e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/30/2018
ms.locfileid: "26992218"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a>Concepts fondamentaux de programmation avec l’API JavaScript pour Excel
 
Cet article décrit comment utiliser l’[API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) afin de créer des compléments pour Excel 2016 ou versions ultérieures. Il présente les concepts fondamentaux de l’utilisation des API et fournit des conseils pour effectuer des tâches spécifiques, comme la lecture ou l’écriture d’une grande plage, la mise à jour de toutes les cellules d’une plage, et bien plus encore.

## <a name="asynchronous-nature-of-excel-apis"></a>Nature asynchrone des API Excel

Les compléments Excel web s’exécutent dans un conteneur de navigateurs qui est incorporé dans l’application Office sur les plateformes basées sur un bureau, comme Office pour Windows, et s’exécute à l’intérieur d’un fichier iFrame HTML dans Office Online. En raison de problèmes de performances, il n’est pas possible d’activer l’API Office.js afin d’interagir de manière synchrone avec l’hôte Excel sur toutes les plateformes prises en charge. Par conséquent, l’appel de l’API **sync()** dans Office.js renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) qui est résolue lorsque l’application Excel termine les actions de lecture ou d’écriture demandées. En outre, vous pouvez mettre en file d’attente plusieurs actions, comme la définition des propriétés ou l’appel de méthodes, et les exécuter en tant que lot de commandes avec un seul appel à **sync()**, au lieu d’envoyer une demande distincte pour chaque action. Les sections suivantes décrivent la façon d’y parvenir à l’aide des API **Excel.run()** et **sync()**.
 
## <a name="excelrun"></a>Excel.run
 
**Excel.Run** exécute une fonction dans laquelle vous spécifiez les actions à effectuer concernant le modèle objet Excel. **Excel.Run** crée automatiquement un contexte de la demande que vous pouvez utiliser pour interagir avec des objets Excel. Lorsque l’API **Excel.run** a fini, une promesse est résolue, et tous les objets alloués lors de l’exécution sont automatiquement publiés.
 
L’exemple suivant montre comment utiliser **Excel.run**. L’instruction catch capture et enregistre les erreurs qui se produisent au sein de **Excel.run**.
 
```js
Excel.run(function (context) {
  // You can use the Excel JavaScript API here in the batch function
  // to execute actions on the Excel object model.
  console.log('Your code goes here.');
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```

### <a name="run-options"></a>Options d’exécution

**Excel.Run** est associé à une surcharge liée à un objet [RunOptions](/javascript/api/excel/excel.runoptions). Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution. La propriété suivante est actuellement prise en charge :

 - `delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **true**, la demande de lot est différée et s’exécute lorsque l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **false**, la demande de lot échoue automatiquement si l’utilisateur est en mode de modification de cellule (entraînant une erreur de contact de l’utilisateur). Le comportement par défaut sans propriété `delayForCellEdit` spécifiée est identique au comportement lorsque la valeur est **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="request-context"></a>Contexte de demande
 
Excel et votre complément sont exécutés dans deux processus distincts. Dans la mesure où ils utilisent des environnements d’exécution différents, les compléments Excel nécessitent un objet **RequestContext** afin de connecter votre complément aux objets dans Excel, tels que les feuilles de calcul, les plages, les graphiques et les tableaux.
 
## <a name="proxy-objects"></a>Objets de proxy
 
Les objets JavaScript pour Excel que vous déclarez et utilisez dans un complément sont des objets proxy. Les méthodes que vous appelez ou les propriétés que vous définissez ou chargez sur les objets proxy sont simplement ajoutées à une file d’attente de commandes en attente. Lorsque vous appelez la méthode **sync()** sur le contexte de demande (par exemple, `context.sync()`), les commandes en attente sont envoyées vers Excel et sont exécutées. L’API JavaScript pour Excel est fondamentalement centrée sur les lots. Vous pouvez mettre en file d’attente autant de modifications que vous le souhaitez dans le contexte de la demande, puis appeler la méthode **sync()** pour exécuter le lot de commandes mises en file d’attente.
 
Par exemple, l’extrait de code suivant déclare l’objet JavaScript local **selectedRange** pour référencer une plage sélectionnée dans le document Excel, puis définit des propriétés sur cet objet. L’objet **selectedRange** est un objet proxy. Les propriétés définies et la méthode appelée sur cet objet ne seront pas répercutées dans le document Excel tant que votre complément n’a pas appelé **context.sync()**.
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a>Sync
 
Tout appel de la méthode **sync()** concernant le contexte de demande synchronise l’état entre les objets proxy et les objets du document Excel. La méthode **sync()** exécute les commandes mises en file d’attente concernant le contexte de demande et récupère des valeurs pour les propriétés qui doivent être chargées dans les objets proxy. La méthode **sync()** est exécutée de façon asynchrone et renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), qui est résolue lorsque la méthode **sync()** est terminée.
 
L’exemple suivant montre une fonction de traitement par lot qui définit un objet proxy JavaScript local (**selectedRange**), charge une propriété de cet objet et utilise ensuite le modèle de promesses JavaScript pour appeler **context.sync()** afin de synchroniser l’état entre les objets proxy et les objets du document Excel.
 
```js
Excel.run(function (context) {
  const selectedRange = context.workbook.getSelectedRange();
  selectedRange.load('address');
  return context.sync()
    .then(function () {
      console.log('The selected range is: ' + selectedRange.address);
  });
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
Dans l’exemple précédent, l’objet **selectedRange** est défini et sa propriété **address** est chargée quand l’élément **context.sync()** est appelé.
 
Étant donné que **sync()** est une opération asynchrone qui renvoie une promesse, vous devez toujours **renvoyer** la promesse (dans JavaScript). Cela garantit que l’opération **sync()** se termine avant que le script continue à s’exécuter. Pour plus d’informations sur l’optimisation des performances avec **sync()**, consultez la rubrique [Optimisation des performances à l’aide de l’API JavaScript d’Excel](https://docs.microsoft.com/office/dev/add-ins/excel/performance).
 
### <a name="load"></a>load()
 
Avant que vous puissiez lire les propriétés d’un objet proxy, vous devez charger explicitement les propriétés pour remplir l’objet proxy avec des données à partir du document Excel, puis appeler **context.sync()**. Par exemple, si vous créez un objet proxy pour référencer une plage sélectionnée, puis que vous voulez lire la propriété **address** de la plage sélectionnée, vous devez charger la propriété **address** avant de pouvoir la lire. Pour demander le chargement de propriétés d’un objet, appelez la méthode **load()** sur l’objet et spécifiez les propriétés à charger. 

> [!NOTE]
> Si vous appelez uniquement des méthodes ou définissez des propriétés sur un objet proxy, il est inutile d’appeler la méthode **load()**. La méthode **load()** n’est nécessaire que lorsque vous souhaitez lire les propriétés sur un objet proxy.
 
À l’instar des demandes de définition de propriétés ou d’appel de méthodes sur des objets proxy, des demandes de chargement de propriétés sur des objets proxy sont ajoutées à la file d’attente des commandes sur le contexte de demande, qui s’exécutera la prochaine fois que vous appellerez la méthode **sync()**. Vous pouvez mettre en file d’attente autant d’appels **load()** sur le contexte de la demande que nécessaire.
 
Dans l’exemple suivant, seules les propriétés spécifiques de la plage sont chargées.
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:B2';
  const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
 
  myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);
 
  return context.sync()
    .then(function () {
      console.log (myRange.address);              // ok
      console.log (myRange.format.wrapText);      // ok
      console.log (myRange.format.fill.color);    // ok
      //console.log (myRange.format.font.color);  // not ok as it was not loaded
  });
}).then(function () {
  console.log('done');
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
Comme `format/font` n’est pas spécifié dans l’appel à **myRange.load()**, la propriété `format.font.color` ne peut pas être lue dans l’exemple précédent.

Pour optimiser les performances, vous devez spécifier explicitement les propriétés et les relations à charger lorsque vous utilisez la méthode **load()** sur un objet, comme abordé dans la rubrique [Optimisation des performances à l’aide de l’API JavaScript d’Excel](performance.md). Pour plus d’informations sur la méthode **load()**, consultez la rubrique [Concepts avancés de programmation avec l’API JavaScript Excel](excel-add-ins-advanced-concepts.md).

## <a name="null-or-blank-property-values"></a>valeurs de propriété null ou vides
 
### <a name="null-input-in-2-d-array"></a>entrée de valeurs null dans un tableau 2D
 
Dans Excel, une plage est représentée par un tableau 2D, où les lignes représentent la première dimension et les colonnes la deuxième. Pour définir des valeurs, un format de nombre ou une formule uniquement pour des cellules spécifiques dans une plage, spécifiez des valeurs, un format de nombre ou une formule pour ces cellules dans le tableau 2D, et indiquez `null` pour toutes les autres cellules du tableau 2D.
 
Par exemple, pour mettre à jour le format de nombre pour une seule cellule dans une plage et conserver le format de nombre existant pour toutes les autres cellules de la plage, spécifiez le nouveau format de nombre de la cellule à mettre à jour, puis spécifiez `null` pour toutes les autres cellules. L’extrait de code suivant définit un nouveau format de nombre pour la quatrième cellule de la plage et ne modifie pas le format de nombre pour les trois premières cellules de la plage.
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a>Entrée null pour une propriété
 
`null` n’est pas une entrée valide pour une propriété unique. Par exemple, l’extrait de code suivant n’est pas valide, car la propriété **values** de la plage ne peut pas être définie sur `null`.
 
```js
range.values = null;
```
 
De même, l’extrait de code suivant n’est pas valide, car `null` n’est pas une valeur valide pour la propriété **color**.
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a>valeurs de la propriété Null dans la réponse
 
Les propriétés de mise en forme comme `size` et `color` contiendront des valeurs `null` dans la réponse lorsque différentes valeurs existent dans la plage spécifiée. Par exemple, si vous récupérez une plage et chargez sa propriété `format.font.color` :
 
* Si toutes les cellules de la plage ont la même couleur de police, `range.format.font.color` spécifie cette couleur.
* Si plusieurs couleurs de police sont présentes dans la plage, `range.format.font.color` est `null`.
 
### <a name="blank-input-for-a-property"></a>Entrée vide pour une propriété
 
Lorsque vous spécifiez une valeur vide pour une propriété (c’est-à-dire deux guillemets droits sans espace entre `''`), cela est interprété comme une instruction d’effacement ou de réinitialisation de la propriété. Par exemple :
 
* Si vous spécifiez une valeur vide pour la propriété `values` d’une plage, le contenu de la plage est effacé.
 
* Si vous spécifiez une valeur vide pour la propriété `numberFormat`, le format de nombre est réinitialisé sur `General`.
 
* Si vous spécifiez une valeur vide pour les propriétés `formula` et `formulaLocale`, les valeurs de la formule sont effacées.
 
### <a name="blank-property-values-in-the-response"></a>Valeurs de propriété vides dans la réponse
 
Pour les opérations de lecture, une valeur de propriété vide dans la réponse (c'est-à-dire, deux guillemets droits sans espace entre `''`) indique que la cellule ne contient pas de donnée ni de valeur. Dans le premier exemple ci-dessous, la première et la dernière cellules de la plage ne contiennent pas de donnée. Dans le deuxième exemple, les deux premières cellules de la plage ne contiennent pas de formule.
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a>Lire ou écrire dans une plage non liée
 
### <a name="read-an-unbounded-range"></a>Lire une plage non liée
 
Une adresse de plage non liée est une adresse de plage qui spécifie des colonnes entières ou des lignes entières. Par exemple :
 
* Adresses de plage composées de colonnes entières :<ul><li>`C:C`</li><li>`A:F`</li></ul>
* Adresses de plage composées de lignes entières :<ul><li>`2:2`</li><li>`1:4`</li></ul>
 
Lorsque l’API effectue une demande de récupération d’une plage non liée (par exemple, `getRange('C:C')`), la réponse contient des valeurs `null` pour les propriétés définies au niveau des cellules, telles que `values`, `text`, `numberFormat` et `formula`. Les autres propriétés de la plage, telles que `address` et `cellCount`, contiennent des valeurs valides pour la plage non liée.
 
### <a name="write-to-an-unbounded-range"></a>Écrire dans une plage non liée
 
Vous ne pouvez pas définir des propriétés au niveau de la cellule telles que `values`, `numberFormat`, et `formula` sur plage non liée, car la demande d’entrée  est trop volumineuse. Par exemple, l’extrait de code suivant n’est pas valide, car il tente de spécifier `values` pour une plage non liée. L’API renvoie une erreur si vous tentez de définir des propriétés au niveau de la cellule pour une plage non liée.
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a>Lire ou écrire dans une grande plage
 
Si une plage contient un grand nombre de cellules, de valeurs, de formats de nombre et/ou de formules, il n’est peut-être pas possible d’exécuter des opérations d’API sur cette plage. L’API essaie toujours d’exécuter au mieux l’opération demandée sur une plage (par exemple, pour extraire ou écrire des données spécifiées), mais essayer d’effectuer des opérations de lecture ou d’écriture pour une grande plage peut provoquer une erreur d’API en raison de l’utilisation des ressources excessive. Pour éviter ces erreurs, nous vous recommandons d’exécuter des opérations de lecture ou d’écriture distinctes pour des sous-ensembles plus petits d’une grande plage, au lieu d’essayer d’exécuter une seule opération de lecture ou d’écriture sur une grande plage.
 
## <a name="update-all-cells-in-a-range"></a>Mettre à jour toutes les cellules d’une plage
 
Pour appliquer la même mise à jour à toutes les cellules d’une plage, (par exemple, pour remplir toutes les cellules avec la même valeur, définir le même format de nombre ou renseigner toutes les cellules avec la même formule), définissez la propriété correspondante dans l’objet **range** sur la valeur (unique) de votre choix.
 
L’exemple suivant obtient une plage qui contient 20 cellules, puis définit le format de nombre et remplit toutes les cellules de la plage avec la valeur **3/11/2015**.
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:A20';
  const worksheet = context.workbook.worksheets.getItem(sheetName);
 
  const range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
 
  return context.sync()
    .then(function () {
      console.log(range.text);
  });
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
##<a name="handle-errors"></a>Gestion des erreurs

Lorsqu’une erreur d’API se produit, l’API renvoie un objet **error** qui contient un code et un message. Pour plus d’informations sur la gestion des erreurs, notamment la liste des erreurs d’API, consultez la rubrique [Gestion des erreurs](excel-add-ins-error-handling.md).
 
## <a name="see-also"></a>Voir aussi
 
* [Prise en main des compléments Excel](excel-add-ins-get-started-overview.md)
* [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [Concepts avancés de programmation avec l’API JavaScript Excel](excel-add-ins-advanced-concepts.md)
* [Optimisation des performances à l’aide de l’API JavaScript d’Excel](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [Référence de l’API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
