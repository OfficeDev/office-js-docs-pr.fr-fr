---
title: Concepts de base de l?API JavaScript Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 1582268a3bdac2b7fe63c4b0a48cf1a19f85bd31
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="excel-javascript-api-core-concepts"></a>Concepts de base de l?API JavaScript pour Excel
 
Cet article d?crit comment utiliser l?[API JavaScript pour Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) afin de cr?er des compl?ments pour Excel 2016. Il pr?sente les concepts fondamentaux de l?utilisation des API et fournit des conseils pour effectuer des t?ches sp?cifiques, comme la lecture ou l??criture d?une grande plage, la mise ? jour de toutes les cellules d?une plage, et bien plus encore.

## <a name="asynchronous-nature-of-excel-apis"></a>Nature asynchrone des API Excel

Les compl?ments Excel web s?ex?cutent dans un conteneur de navigateurs qui est incorpor? dans l?application Office sur les plateformes bas?es sur un bureau, comme Office pour Windows, et s?ex?cute ? l?int?rieur d?un fichier iFrame HTML dans Office Online. En raison de probl?mes de performances, il n?est pas possible d?activer l?API Office.js afin d?interagir de mani?re synchrone avec l?h?te Excel sur toutes les plateformes prises en charge. Par cons?quent, l?appel de l?API **sync()** dans Office.js renvoie une [promesse](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) qui est r?solue lorsque l?application Excel termine les actions de lecture ou d??criture demand?es. En outre, vous pouvez mettre en file d?attente plusieurs actions, comme la d?finition des propri?t?s ou l?appel de m?thodes, et les ex?cuter en tant que lot de commandes avec un seul appel ? **sync()**, au lieu d?envoyer une demande distincte pour chaque action. Les sections suivantes d?crivent la fa?on d?y parvenir ? l?aide des API **Excel.run()** et **sync()**.
 
## <a name="excelrun"></a>Excel.run
 
**Excel.Run** ex?cute une fonction dans laquelle vous sp?cifiez les actions ? effectuer concernant le mod?le objet Excel. **Excel.Run** cr?e automatiquement un contexte de la demande que vous pouvez utiliser pour interagir avec des objets Excel. Lorsque l?API **Excel.run** a fini, une promesse est r?solue, et tous les objets allou?s lors de l?ex?cution sont automatiquement publi?s.
 
L?exemple suivant montre comment utiliser **Excel.run**. L?instruction catch capture et enregistre les erreurs qui se produisent au sein de **Excel.run**.
 
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

## <a name="request-context"></a>Contexte de demande
 
Excel et votre compl?ment sont ex?cut?s dans deux processus distincts. Dans la mesure o? ils utilisent des environnements d?ex?cution diff?rents, les compl?ments Excel n?cessitent un objet **RequestContext** afin de connecter votre compl?ment aux objets dans Excel, tels que les feuilles de calcul, les plages, les graphiques et les tableaux.
 
## <a name="proxy-objects"></a>Objets de proxy
 
Les objets JavaScript pour Excel que vous d?clarez et utilisez dans un compl?ment sont des objets proxy. Les m?thodes que vous appelez ou les propri?t?s que vous d?finissez ou chargez sur les objets proxy sont simplement ajout?es ? une file d?attente de commandes en attente. Lorsque vous appelez la m?thode **sync()** sur le contexte de demande (par exemple, `context.sync()`), les commandes en attente sont envoy?es vers Excel et sont ex?cut?es. L?API JavaScript pour Excel est fondamentalement centr?e sur les lots. Vous pouvez mettre en file d?attente autant de modifications que vous le souhaitez dans le contexte de la demande, puis appeler la m?thode **sync()** pour ex?cuter le lot de commandes mises en file d?attente.
 
Par exemple, l?extrait de code suivant d?clare l?objet JavaScript local **selectedRange** pour r?f?rencer une plage s?lectionn?e dans le document Excel, puis d?finit des propri?t?s sur cet objet. L?objet **selectedRange** est un objet proxy. Les propri?t?s d?finies et la m?thode appel?e sur cet objet ne seront pas r?percut?es dans le document Excel tant que votre compl?ment n?a pas appel? **context.sync()**.
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a>Sync
 
Tout appel de la m?thode **sync()** concernant le contexte de demande synchronise l??tat entre les objets proxy et les objets du document Excel. La m?thode **sync()** ex?cute les commandes mises en file d?attente concernant le contexte de demande et r?cup?re des valeurs pour les propri?t?s qui doivent ?tre charg?es dans les objets proxy. La m?thode **sync()** est ex?cut?e de fa?on asynchrone et renvoie une [promesse](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise), qui est r?solue lorsque la m?thode **sync()** est termin?e.
 
L?exemple suivant montre une fonction de traitement par lot qui d?finit un objet proxy JavaScript local (**selectedRange**), charge une propri?t? de cet objet et utilise ensuite le mod?le de promesses JavaScript pour appeler **context.sync()** afin de synchroniser l??tat entre les objets proxy et les objets du document Excel.
 
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
 
Dans l?exemple pr?c?dent, l?objet **selectedRange** est d?fini et sa propri?t? **address** est charg?e quand l??l?ment **context.sync()** est appel?.
 
?tant donn? que **sync()** est une op?ration asynchrone qui renvoie une promesse, vous devez toujours **renvoyer** la promesse (dans JavaScript). Cela garantit que l?op?ration **sync()** se termine avant que le script continue ? s?ex?cuter. Pour plus d?informations sur l?optimisation des performances avec **sync ()**, voir [Optimisation des performances de l?API JavaScript pour Excel](https://dev.office.com/reference/add-ins/excel/performance.md).
 
### <a name="load"></a>load()
 
Avant que vous puissiez lire les propri?t?s d?un objet proxy, vous devez charger explicitement les propri?t?s pour remplir l?objet proxy avec des donn?es ? partir du document Excel, puis appeler **context.sync()**. Par exemple, si vous cr?ez un objet proxy pour r?f?rencer une plage s?lectionn?e, puis que vous voulez lire la propri?t? **address** de la plage s?lectionn?e, vous devez charger la propri?t? **address** avant de pouvoir la lire. Pour demander le chargement de propri?t?s d?un objet, appelez la m?thode **load()** sur l?objet et sp?cifiez les propri?t?s ? charger. 

> [!NOTE]
> Si vous appelez uniquement des m?thodes ou d?finissez des propri?t?s sur un objet proxy, il est inutile d?appeler la m?thode **load()**. La m?thode **load()** n?est n?cessaire que lorsque vous souhaitez lire les propri?t?s sur un objet proxy.
 
? l?instar des demandes de d?finition de propri?t?s ou d?appel de m?thodes sur des objets proxy, des demandes de chargement de propri?t?s sur des objets proxy sont ajout?es ? la file d?attente des commandes sur le contexte de demande, qui s?ex?cutera la prochaine fois que vous appellerez la m?thode **sync()**. Vous pouvez mettre en file d?attente autant d?appels **load()** sur le contexte de la demande que n?cessaire.
 
Dans l?exemple suivant, seules les propri?t?s sp?cifiques de la plage sont charg?es.
 
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
 
Comme `format/font` n?est pas sp?cifi? dans l?appel ? **myRange.load()**, la propri?t? `format.font.color` ne peut pas ?tre lue dans l?exemple pr?c?dent.

Pour optimiser les performances, vous devez sp?cifier explicitement les propri?t?s et les relations de chargement lorsque vous utilisez la m?thode **load()** sur un objet, tel que d?crit dans [Optimisations des performances de l?API JavaScript pour Excel](performance.md). Pour plus d?informations sur la m?thode **load()**, reportez-vous ? la rubrique [Concepts avanc?s pour l?API JavaScript pour Excel](excel-add-ins-advanced-concepts.md).

## <a name="null-or-blank-property-values"></a>valeurs de propri?t? null ou vides
 
### <a name="null-input-in-2-d-array"></a>entr?e de valeurs null dans un tableau 2D
 
Dans Excel, une plage est repr?sent?e par un tableau 2D, o? les lignes repr?sentent la premi?re dimension et les colonnes la deuxi?me. Pour d?finir des valeurs, un format de nombre ou une formule uniquement pour des cellules sp?cifiques dans une plage, sp?cifiez des valeurs, un format de nombre ou une formule pour ces cellules dans le tableau 2D, et indiquez `null` pour toutes les autres cellules du tableau 2D.
 
Par exemple, pour mettre ? jour le format de nombre pour une seule cellule dans une plage et conserver le format de nombre existant pour toutes les autres cellules de la plage, sp?cifiez le nouveau format de nombre de la cellule ? mettre ? jour, puis sp?cifiez `null` pour toutes les autres cellules. L?extrait de code suivant d?finit un nouveau format de nombre pour la quatri?me cellule de la plage et ne modifie pas le format de nombre pour les trois premi?res cellules de la plage.
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a>Entr?e null pour une propri?t?
 
`null` n?est pas une entr?e valide pour une propri?t? unique. Par exemple, l?extrait de code suivant n?est pas valide, car la propri?t? **values** de la plage ne peut pas ?tre d?finie sur `null`.
 
```js
range.values = null;
```
 
De m?me, l?extrait de code suivant n?est pas valide, car `null` n?est pas une valeur valide pour la propri?t? **color**.
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a>valeurs de la propri?t? Null dans la r?ponse
 
Les propri?t?s de mise en forme comme `size` et `color` contiendront des valeurs `null` dans la r?ponse lorsque diff?rentes valeurs existent dans la plage sp?cifi?e. Par exemple, si vous r?cup?rez une plage et chargez sa propri?t? `format.font.color` :
 
* Si toutes les cellules de la plage ont la m?me couleur de police, `range.format.font.color` sp?cifie cette couleur.
* Si plusieurs couleurs de police sont pr?sentes dans la plage, `range.format.font.color` est `null`.
 
### <a name="blank-input-for-a-property"></a>Entr?e vide pour une propri?t?
 
Lorsque vous sp?cifiez une valeur vide pour une propri?t? (c?est-?-dire deux guillemets droits sans espace entre `''`), cela est interpr?t? comme une instruction d?effacement ou de r?initialisation de la propri?t?. Par exemple :
 
* Si vous sp?cifiez une valeur vide pour la propri?t? `values` d?une plage, le contenu de la plage est effac?.
 
* Si vous sp?cifiez une valeur vide pour la propri?t? `numberFormat`, le format de nombre est r?initialis? sur `General`.
 
* Si vous sp?cifiez une valeur vide pour les propri?t?s `formula` et `formulaLocale`, les valeurs de la formule sont effac?es.
 
### <a name="blank-property-values-in-the-response"></a>Valeurs de propri?t? vides dans la r?ponse
 
Pour les op?rations de lecture, une valeur de propri?t? vide dans la r?ponse (c'est-?-dire, deux guillemets droits sans espace entre `''`) indique que la cellule ne contient pas de donn?e ni de valeur. Dans le premier exemple ci-dessous, la premi?re et la derni?re cellules de la plage ne contiennent pas de donn?e. Dans le deuxi?me exemple, les deux premi?res cellules de la plage ne contiennent pas de formule.
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a>Lire ou ?crire dans une plage non li?e
 
### <a name="read-an-unbounded-range"></a>Lire une plage non li?e
 
Une adresse de plage non li?e est une adresse de plage qui sp?cifie des colonnes enti?res ou des lignes enti?res. Par exemple :
 
* Adresses de plage compos?es de colonnes enti?res :<ul><li>`C:C`</li><li>`A:F`</li></ul>
* Adresses de plage compos?es de lignes enti?res :<ul><li>`2:2`</li><li>`1:4`</li></ul>
 
Lorsque l?API effectue une demande de r?cup?ration d?une plage non li?e (par exemple, `getRange('C:C')`), la r?ponse contient des valeurs `null` pour les propri?t?s d?finies au niveau des cellules, telles que `values`, `text`, `numberFormat` et `formula`. Les autres propri?t?s de la plage, telles que `address` et `cellCount`, contiennent des valeurs valides pour la plage non li?e.
 
### <a name="write-to-an-unbounded-range"></a>?crire dans une plage non li?e
 
Vous ne pouvez pas d?finir des propri?t?s au niveau de la cellule telles que `values`, `numberFormat`, et `formula` sur plage non li?e, car la demande d?entr?e  est trop volumineuse. Par exemple, l?extrait de code suivant n?est pas valide, car il tente de sp?cifier `values` pour une plage non li?e. L?API renvoie une erreur si vous tentez de d?finir des propri?t?s au niveau de la cellule pour une plage non li?e.
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a>Lire ou ?crire dans une grande plage
 
Si une plage contient un grand nombre de cellules, de valeurs, de formats de nombre et/ou de formules, il n?est peut-?tre pas possible d?ex?cuter des op?rations d?API sur cette plage. L?API essaie toujours d?ex?cuter au mieux l?op?ration demand?e sur une plage (par exemple, pour extraire ou ?crire des donn?es sp?cifi?es), mais essayer d?effectuer des op?rations de lecture ou d??criture pour une grande plage peut provoquer une erreur d?API en raison de l?utilisation des ressources excessive. Pour ?viter ces erreurs, nous vous recommandons d?ex?cuter des op?rations de lecture ou d??criture distinctes pour des sous-ensembles plus petits d?une grande plage, au lieu d?essayer d?ex?cuter une seule op?ration de lecture ou d??criture sur une grande plage.
 
## <a name="update-all-cells-in-a-range"></a>Mettre ? jour toutes les cellules d?une plage
 
Pour appliquer la m?me mise ? jour ? toutes les cellules d?une plage, (par exemple, pour remplir toutes les cellules avec la m?me valeur, d?finir le m?me format de nombre ou renseigner toutes les cellules avec la m?me formule), d?finissez la propri?t? correspondante dans l?objet **range** sur la valeur (unique) de votre choix.
 
L?exemple suivant obtient une plage qui contient 20 cellules, puis d?finit le format de nombre et remplit toutes les cellules de la plage avec la valeur **3/11/2015**.
 
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
 
## <a name="error-messages"></a>Messages d?erreur
 
Lorsqu?une erreur d?API se produit, l?API renvoie un objet **error** qui contient un code et un message. Le tableau suivant d?finit une liste des erreurs que l?API peut renvoyer.
 
|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |L?argument est manquant ou non valide, ou a un format incorrect.|
|InvalidRequest  |Impossible de traiter la demande.|
|InvalidReference|Cette r?f?rence n?est pas valide pour l?op?ration en cours.|
|InvalidBinding  |Cette liaison d?objets n?est plus valide en raison de mises ? jour pr?c?dentes.|
|InvalidSelection|La s?lection en cours est incorrecte pour cette action.|
|Unauthenticated |Les informations d?authentification requises sont manquantes ou incorrectes.|
|AccessDenied |Vous ne pouvez pas effectuer l?op?ration demand?e.|
|ItemNotFound |La ressource demand?e n?existe pas.|
|ActivityLimitReached|La limite d?activit? a ?t? atteinte.|
|GeneralException|Une erreur interne s?est produite lors du traitement de la demande.|
|NotImplemented  |La fonctionnalit? demand?e n?est pas impl?ment?e|
|ServiceNotAvailable|Le service n?est pas disponible.|
|Conflict              |La demande n?a pas pu ?tre trait?e en raison d?un conflit.|
|ItemAlreadyExists|La ressource en cours de cr?ation existe d?j?.|
|UnsupportedOperation|L?op?ration tent?e n?est pas prise en charge.|
|RequestAborted|La demande a ?t? interrompue pendant l?ex?cution.|
|ApiNotAvailable|L?API demand?e n?est pas disponible.|
|InsertDeleteConflict|L?op?ration d?insertion ou de suppression tent?e a cr?? un conflit.|
|InvalidOperation|L?op?ration tent?e n?est pas valide sur l?objet.|
 
## <a name="see-also"></a>Voir aussi
 
* [Prise en main des compl?ments Excel](excel-add-ins-get-started-overview.md)
* [Exemples de code pour les compl?ments Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Optimisation des performances de l'API JavaScript pour Excel](https://dev.office.com/reference/add-ins/excel/performance.md)
* [R?f?rence de l?API JavaScript pour Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
