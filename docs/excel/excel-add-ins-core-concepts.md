---
title: Concepts fondamentaux de programmation avec l’API JavaScript pour Excel
description: Utilisez l’API JavaScript pour Excel afin de créer des compléments pour Excel.
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: dde7dc66e0746fc4d9cf91ed3df824fab05c109d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292592"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a>Concepts fondamentaux de programmation avec l’API JavaScript pour Excel

Cet article décrit comment utiliser l’[API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md) afin de créer des compléments pour Excel 2016 ou versions ultérieures. Il présente les concepts fondamentaux de l’utilisation des API et fournit des conseils pour effectuer des tâches spécifiques, comme la lecture ou l’écriture d’une grande plage, la mise à jour de toutes les cellules d’une plage, et bien plus encore.

> [!IMPORTANT]
> Pour en savoir plus sur la nature asynchrone des API Excel et la manière dont elles fonctionnent avec le classeur, voir [Utilisation du modèle d’API spécifique à l’application](../develop/application-specific-api-model.md).  

## <a name="officejs-apis-for-excel"></a>API Office.js pour Excel

Un complément Excel interagit avec des objets dans Excel en utilisant l’API Office JavaScript, qui inclut deux modèles d’objets JavaScript :

* **API JavaScript pour Excel** : inclut dans Office 2016, l’[API JavaScript Excel](../reference/overview/excel-add-ins-reference-overview.md) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore.

* **API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.

Vous utiliserez probablement l’API JavaScript Excel pour développer la majorité des fonctionnalités des compléments destinés à Excel 2016 ou version ultérieure, vous utiliserez également des objets dans l’API commune. Par exemple :

* [Context](/javascript/api/office/office.context) :le `Context` représente l’environnement d’exécution du complément et permet d’accéder à des objets clés de l’API. Il se compose de détails sur la configuration du classeur comme `contentLanguage` et `officeTheme`, et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`. En outre, il fournit la méthode `requirements.isSetSupported()` que vous pouvez utiliser pour vérifier si l’ensemble de conditions requises spécifié est pris en charge par l’application Excel dans laquelle le complément est exécuté.
* [Document](/javascript/api/office/office.document) : le `Document` fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Excel dans lequel le complément est exécuté.

L’image suivante illustre les situations dans lesquelles vous pouvez utiliser l’API JavaScript Excel ou les API communes.

![Image des différences entre l’API Excel et les API communes](../images/excel-js-api-common-api.png)

## <a name="object-model"></a>Modèle d’objet

Pour comprendre les API Excel, vous devez connaître la manière dont les composants d’un classeur sont liés les uns aux autres.

* Un **classeur** contient une ou plusieurs **feuilles de calcul**.
* Une **feuille de calcul** donne accès à des cellules via **plage** objets.
* Une **plage** représente un groupe de cellules contiguës.
* Les **plages** sont utilisées pour créer et placer des **tableaux**, des **graphiques**, des **formes** et d’autres objets d’organisation ou de visualisation de données.
* Une **feuille de calcul** contient des collections d’objets de données présents dans la feuille individuelle.
* Les **classeurs** contiennent des collections de certains de ces objets de données (par exemple : les **tableaux**) pour l'ensemble du **classeur**.

### <a name="ranges"></a>Plages

Une plage est un groupe de cellules contiguës dans le classeur. Les compléments utilisent généralement la notation de style A1 (par exemple : **B3** pour la cellule unique de la colonne **B** et de la ligne **3** ou **C2:F4** pour les cellules des colonnes **C** à **F** et des lignes **2** à **4**) pour définir les plages.

Les plages comportent trois propriétés principales : `values`, `formulas`et `format`. Ces propriétés obtiennent ou définissent les valeurs des cellules, les formules à évaluer et la mise en forme visuelle des cellules.

#### <a name="range-sample"></a>Exemple de plage

L’exemple de code suivant montre comment créer des registres des ventes. Cette fonction utilise les objets `Range` pour déterminer les valeurs, les formules et les formats.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

Cet exemple crée les données suivantes dans la feuille de calcul active :

![Un registre des ventes affiche des lignes de valeur, une colonne de formule et des en-têtes mis en forme.](../images/excel-overview-range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Graphiques, tableaux et autres objets de données

Les API JavaScript Excel peuvent créer et manipuler les structures de données et les visualisations dans Excel. Les tableaux et les graphiques sont deux des objets les plus fréquemment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et bien plus encore.

#### <a name="creating-a-table"></a>Création d’un tableau

Créez des tableaux à l’aide des plages de données remplies. Les contrôles de mise en forme et du tableau (par exemple, les filtres) sont automatiquement appliqués à la plage.

L’exemple suivant crée un tableau à l’aide des plages de l’exemple précédent.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

L’exécution de cet exemple de code sur la feuille de calcul avec les données précédentes crée le tableau suivant :

![Un tableau créée à partir du registre des ventes précédent.](../images/excel-overview-table-sample.png)

#### <a name="creating-a-chart"></a>Création d’un graphique

Vous pouvez créer un graphique pour visualiser les données d’une plage. Les API prennent en charge des dizaines de variétés de graphiques, chacun pouvant être personnalisé selon vos besoins.

L’exemple suivant crée un histogramme pour trois éléments et place celui-ci 100 pixels en dessous de la partie supérieure de la feuille de calcul.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

L’exécution de cet exemple sur la feuille de calcul avec le tableau précédent crée le graphique suivant :

![Histogramme montrant les quantités pour trois des articles présents dans le registre des ventes précédent.](../images/excel-overview-chart-sample.png)

## <a name="run-options"></a>Options d’exécution

`Excel.run` est associé à une surcharge liée à un objet [RunOptions](/javascript/api/excel/excel.runoptions). Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution. La propriété suivante est actuellement prise en charge :

* `delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **true**, la demande de lot est différée et s’exécute lorsque l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **false**, la demande de lot échoue automatiquement si l’utilisateur est en mode de modification de cellule (entraînant une erreur de contact de l’utilisateur). Le comportement par défaut sans propriété `delayForCellEdit` spécifiée est identique au comportement lorsque la valeur est **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="null-or-blank-property-values"></a>Valeurs de propriété null ou vides

`null` et les chaînes vides ont des implications particulières dans les API JavaScript Excel. Elles sont utilisées pour représenter les cellules vides, l’absence de mise en forme ou les valeurs par défaut. Cette section décrit l’utilisation de `null` et d’une chaîne vide lors de l’obtention et de la définition de propriétés.

### <a name="null-input-in-2-d-array"></a>entrée de valeurs null dans un tableau 2D

Dans Excel, une plage est représentée par un tableau 2D, où les lignes représentent la première dimension et les colonnes la deuxième. Pour définir des valeurs, un format de nombre ou une formule uniquement pour des cellules spécifiques dans une plage, spécifiez des valeurs, un format de nombre ou une formule pour ces cellules dans le tableau 2D, et indiquez `null` pour toutes les autres cellules du tableau 2D.

Par exemple, pour mettre à jour le format de nombre pour une seule cellule dans une plage et conserver le format de nombre existant pour toutes les autres cellules de la plage, spécifiez le nouveau format de nombre de la cellule à mettre à jour, puis spécifiez `null` pour toutes les autres cellules. L’extrait de code suivant définit un nouveau format de nombre pour la quatrième cellule de la plage et ne modifie pas le format de nombre pour les trois premières cellules de la plage.

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a>Entrée null pour une propriété

`null` n’est pas une entrée valide pour une propriété unique. Par exemple, l’extrait de code suivant n’est pas valide, car la propriété `values` de la plage ne peut pas être définie sur `null`.

```js
range.values = null;
```

De même, l’extrait de code suivant n’est pas valide, car `null` n’est pas une valeur valide pour la propriété `color`.

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a>valeurs de la propriété Null dans la réponse

Les propriétés de mise en forme comme `size` et `color` contiendront des valeurs `null` dans la réponse lorsque différentes valeurs existent dans la plage spécifiée. Par exemple, si vous récupérez une plage et chargez sa propriété `format.font.color`:

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

## <a name="requirement-sets"></a>Ensembles de conditions requises

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Le complément Office peut effectuer une vérification à l’exécution ou utiliser des ensembles de conditions requises spécifiés dans le manifeste pour déterminer si une application Office prend en charge les API requises par le complément. Pour identifier les ensembles de conditions requises spécifiques disponibles sur chaque plateforme prise en charge, reportez-vous à [Ensembles de conditions requises de l’API JavaScript pour Excel](../reference/requirement-sets/excel-api-requirement-sets.md).

### <a name="checking-for-requirement-set-support-at-runtime"></a>Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution

L’exemple de code suivant montre comment déterminer si l’application Office dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste

Vous pouvez utiliser l’[élément Requirements](../reference/manifest/requirements.md) dans le manifeste de complément pour spécifier les ensembles de conditions requises minimales et/ou les méthodes d’API que votre complément doit activer. Si la plateforme ou l’application Office ne prend pas en charge les ensembles de conditions requises ou les méthodes d’API spécifiées dans l’élément `Requirements` du manifeste, le complément ne s’exécute pas dans cette application ou plateforme et ne s’affiche pas dans la liste de compléments dans **Mes compléments**.

L’exemple de code suivant montre l’élément `Requirements` dans un manifeste indiquant que le complément doit être chargé dans toutes les applications clientes Office prenant en charge l’ensemble de conditions requises ExcelApi version 1.3 ou ultérieure.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Pour rendre votre complément disponible sur toutes les plateformes d’une application Office, comme Excel sur le web, Windows et iPad, nous vous recommandons de vérifier la prise en charge des conditions requises lors de l’exécution au lieu de définir la prise en charge d’ensemble de conditions requises dans le manifeste.

### <a name="requirement-sets-for-the-officejs-common-api"></a>Ensembles de conditions requises pour l’API commune Office.js

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](../reference/requirement-sets/office-add-in-requirement-sets.md).

## <a name="handle-errors"></a>Gestion des erreurs

Lorsqu’une erreur d’API se produit, l’API renvoie un objet `error` qui contient un code et un message. Pour plus d’informations sur la gestion des erreurs, notamment la liste des erreurs d’API, consultez la rubrique [Gestion des erreurs](excel-add-ins-error-handling.md).

## <a name="see-also"></a>Voir aussi

* [Création de votre premier complément Excel](../quickstarts/excel-quickstart-jquery.md)
* [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Optimisation des performances à l’aide de l’API JavaScript d’Excel](../excel/performance.md)
* [Référence de l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)
* [Problèmes courants liés au code et comportements de plateforme inattendus](../develop/common-coding-issues.md)
