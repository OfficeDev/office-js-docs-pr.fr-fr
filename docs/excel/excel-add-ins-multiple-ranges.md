---
title: Travailler simultanément avec plusieurs plages dans des compléments Excel
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: 37f9c8a9f3127d78e1cc794aea9e6d1502cdeaf9
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270977"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Travailler simultanément avec plusieurs plages dans des compléments Excel (Prévisualisation)

La bibliothèque JavaScript Excel permet à votre complément d’effectuer des opérations et définir des propriétés, de manière simultanée sur plusieurs plages. Les plages n’ont pas besoin d’être contigus. En plus de rendre votre code plus simple, cette manière de paramétrer une propriété s’exécute beaucoup plus rapidement que paramétrer la même propriété pour chaque les plages de manière individuelle.

> [!NOTE]
> Les APIs décrits dans cet article nécessitent**la version Office 2016 «Démarrer en un Clic» 1809 Build 10820.20000**ou ultérieure. (Vous devrez peut-être rejoindre le[programme Office Insider](https://products.office.com/office-insider) pour obtenir une build appropriée.) Par ailleurs, vous devez charger la version bêta de la bibliothèque JavaScript Office à partir de [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Enfin, nous n’avons pas encore les pages de référence pour ces APIs. Mais le fichier de type définition suivant comporte des descriptions à leur place: [office.d.ts bêta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).

## <a name="rangeareas"></a>RangeAreas

Un ensemble de plages (voire non contiguës) est représenté par un `Excel.RangeAreas` objet. Il possède des propriétés et des méthodes similaires au type`Range` (bon nombre des noms identiques ou similaires,), mais les ajustements ont été apportées à:

- Les types de données pour les propriétés et le comportement des méthodes et des getters.
- Les types de données de paramètres et des comportements de la méthode.
- Les types de données de méthodes renvoient des valeurs.

Quelques exemples :

- `RangeAreas` a une`address` propriété qui renvoie une chaîne séparée par des virgules de plage d’adresses, au lieu d’une adresse comme avec la `Range.address` propriété.
- `RangeAreas` a une`dataValidation` propriété qui renvoie un`DataValidation` objet qui représente la validation des données de toutes les plages dans la `RangeAreas`, s’il est cohérent. La propriété est`null` si les objets`DataValidation` identiques ne sont pas appliqués à toutes les plages dans la`RangeAreas`. Il s’agit d’un principe général, mais pas universel avec l’`RangeAreas` objet: *si une propriété ne comporte pas de valeurs cohérentes sur tous les plages dans la`RangeAreas`, cela signifie`null`.* Voir[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) pour plus d’informations et quelques exceptions.
- `RangeAreas.cellCount` Obtient le nombre total de cellules dans toutes les plages dans la`RangeAreas`.
- `RangeAreas.calculate` recalcule les cellules de toutes les plages dans la`RangeAreas`.
- `RangeAreas.getEntireColumn` et `RangeAreas.getEntireRow` retourner un autre`RangeAreas` objet qui représente toutes les colonnes (ou lignes) dans toutes les plages dans la `RangeAreas`. Par exemple, si le`RangeAreas` représente « A1 : C4 » et « F14:L15 », puis `RangeAreas.getEntireColumn` renvoie un`RangeAreas` objet qui représente « A:C » et « F:L ».
- `RangeAreas.copyFrom` peut prendre soit un`Range` ou d’un`RangeAreas` paramètre représentant la ou les plage(s) source de l’opération de copie.

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>La liste complète des membres plage sont également disponibles sur RangeAreas

##### <a name="properties"></a>Propriétés

Être familiarisé avec[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) avant d’écrire de code qui lit les propriétés répertoriées. Il existe des subtilités sur ce qui est renvoyé.

- address
- addressLocal
- cellCount
- conditionalFormats
- context
- dataValidation
- format
- isEntireColumn
- isEntireRow
- style
- worksheet

##### <a name="methods"></a>Méthodes

Les méthodes de plage dans l’aperçu sont marquées.

- calculate()
- clear()
- convertDataTypeToText() (preview)
- convertToLinkedDataType() (preview)
- copyFrom() (preview)
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange() (appelé getOffsetRangeAreas sur l’objet RangeAreas)
- getSpecialCells() (prévisualisation)
- getSpecialCellsOrNullObject() (prévisualisation)
- getTables() (prévisualisation)
- getUsedRange() (appelé getOffsetRangeAreas sur l’objet RangeAreas)
- getUsedRangeOrNullObject() (appelé getUsedRangeAreasOrNullObject sur l’objet RangeAreas)
- load()
- set()
- setDirty() (prévisualisation)
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>Méthodes et propriétés propres à une langue RangeArea

Le `RangeAreas` type possède des propriétés et des méthodes qui ne sont pas sur l’`Range`objet. Ce qui est une sélection de certains d’entre eux :

- `areas`: A`RangeCollection` objet qui contient toutes les plages représentées par l’ `RangeAreas`objet. L’`RangeCollection`objet est également nouveau et est semblable à d’autres objets de collection de sites Excel. Il possède une`items`propriété est une matrice d’`Range` objets représentant les plages.
- `areaCount`: Le nombre total de plages dans le`RangeAreas`.
- `getOffsetRangeAreas`: Fonctionne comme[Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), sauf qu’une `RangeAreas` est renvoyée et il contient des plages sont en décalage avec des plages du fichier d’origine`RangeAreas`.

## <a name="create-rangeareas-and-set-properties"></a>Créer RangeAreas et définir les propriétés

Vous pouvez créer`RangeAreas`l’objet selon deux méthodes de base:

- Appeler`Worksheet.getRanges()` et de transmettre une chaîne comportant des adresses de plage séparées par des virgules. Si une plage que vous souhaitez inclure a été modifiée en[NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), vous pouvez inclure le nom, au lieu de l’adresse, dans la chaîne.
- Appel `Workbook.getSelectedRanges()`. Cette méthode renvoie une`RangeAreas`représentation de toutes les plages qui sont sélectionnées sur la feuille de calcul active actuelle.

Une fois que vous avez un`RangeAreas`objet, vous pouvez en créer d’autres à l’aide des méthodes sur l’objet qui renvoie`RangeAreas`tel que`getOffsetRangeAreas`et`getIntersection`.

> [!NOTE]
> Vous ne pouvez pas ajouter directement des plages supplémentaires à un objet`RangeAreas`. Par exemple, la collection dans`RangeAreas.areas`n’a pas une méthode`add`.


> [!WARNING] 
> N’essayez pas d’ajouter ou de supprimer directement les membres du tableau`RangeAreas.areas.items`. Cela mènera à un comportement indésirable dans votre code. Par exemple, il est possible de pousser un objet`Range` supplémentaire sur le tableau, mais ceci entrainera des erreurs car les propriétés`RangeAreas`et les méthodes se comportent comme si le nouvel élément n’existait pas. Par exemple, la propriété`areaCount`n’inclut pas les plages poussées de cette manière, et le `RangeAreas.getItemAt(index)` lance une erreur si`index`est plus grand que`areasCount-1`. De la même façon, supprimer un objet`Range`dans la plage`RangeAreas.areas.items`en obtenant une référence liée à celui-ci et en appelant sa méthode`Range.delete` entraîne des bogues: bien que `Range`l’objet*soit*supprimé, les propriétés et les méthodes de l’objet`RangeAreas`parent se comporte, ou tente de le faire, comme s’il existait encore. Par exemple, si votre code appelle`RangeAreas.calculate`, Office essaiera de calculer la plage, mais comprendra une erreur car l’objet de la plage n’est plus là.

Paramétrer une propriété sur un`RangeAreas`établit une propriété correspondante sur toutes les plages dans la collection`RangeAreas.areas`.

Le suivant est un exemple de paramétrage d’une propriété sur des plages multiples. La fonction surligne les plages**F3:F5** and **H3:H5**.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Cet exemple s’applique aux scénarios dans lesquels vous pouvez coder en dur les adresses de plage que vous passez à`getRanges`ou facilement les calculer à l’exécution. Certains des scénarios dans lesquels ceci peut s’appliquer incluent: 

- Le code s’exécute dans le contexte d’un modèle connu.
- Le code s’exécute dans le contexte de données importées où le schéma des données est connu.

Lorsque vous ne pouvez pas connaitre au moment de coder quelles plages sont nécessaires pour opérer, vous devez les découvrir lors de l’exécution. La prochaine section traite de ces scénarios.

### <a name="discover-range-areas-programmatically"></a>Découvrez les zones de plage au niveau de la programmation

Les méthodes `Range.getSpecialCells()`et`Range.getSpecialCellsOrNullObject()`vous permettent de trouver lors de l’exécution les plages que vous souhaitez faire fonctionner sur la base des caractéristiques des cellules et du type des valeurs des cellules. Voici les signatures des méthodes à partir des types de fichiers de données TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

Voici un exemple d’utilisation de la première. Tenez compte du code suivant:

- Cela limite la partie de la feuille qui nécessite d’être recherchée en appelant d’abord`Worksheet.getUsedRange`et en appelant`getSpecialCells`uniquement pour cette plage.
- Il passe comme paramètre à la version chaîne`getSpecialCells`d’une valeur à partir du enum`Excel.SpecialCellType`. Certaines des autres valeurs qui peuvent être passées à la place sont «Vides»pour des cellules vides, «Constantes» pour des cellules avec des valeurs littérales au lieu des formules, et «SameConditionalFormat» pour les cellules qui disposent de la même mise en forme conditionnelle comme la première cellule dans le`usedRange`. La première cellule est la cellule en haut toute à gauche. Pour une liste complète des valeurs dans l’enum, voir[beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).
- La`getSpecialCells`méthode renvoie un`RangeAreas`objet, toutes les cellules alors dotées de formules seront colorées en rose même si elles ne sont pas adjacentes. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

Parfois la plage ne dispose pas *de*cellules avec la caractéristique ciblée. Si`getSpecialCells`n’en trouve pas, elle lance une erreur**ItemNotFound**. Cela dévie le flux de contrôle vers un(e)`catch`bloc/méthode, s’il en existe. S’il n’existe pas, l’erreur arrête la fonction. Il peut avoir des scénarios dans lesquels émettre l’erreur est exactement ce que vous souhaitez lorsqu’il n’y a pas de cellules avec de caractéristique ciblée. 

Mais dans les scénarios dans lesquels cela est normal, mais peut-être gênant, pour les cellules qui correspondent pas; votre code doit vérifier cette possibilité et le gérer gracieusement sans émettre d’erreur. Pour ces scénarios, utilisez la méthode`getSpecialCellsOrNullObject` et testez la propriété`RangeAreas.isNullObject`. Voici un exemple. Tenez compte du code suivant:

- La méthode`getSpecialCellsOrNullObject`renvoie toujours un objet proxy, donc il ne s’agit jamais du sens`null`JavaScript ordinaire. Mais si les cellules non correspondantes sont introuvables, la propriété`isNullObject` de l’objet est établi à`true`.
- Il appelle`context.sync`*avant*de tester la propriété`isNullObject`. Il s’agit d’une condition avec toutes les méthodes et propriétés`*OrNullObject`, car vous devez toujours télécharger et synchroniser une propriété afin de le lire.  Cependant, il n’est pas nécessaire de télécharger*de manière explicite*la propriété`isNullObject`. Il est automatiquement téléchargé par le`context.sync`même si`load`n’est pas appelé sur l’objet. Pour plus d'informations, consultez le[\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).
- Vous pouvez tester ce code en sélectionnant d’abord une plage qui n’a pas de cellules de formule et en l’exécutant. Puis sélectionnez une plage qui dispose au moins d’une cellule dotée d’une formule et en l’exécutant à nouveau.

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

Par simplicité, tous les autres exemples dans cet article, utilisez la méthode`getSpecialCells`au lieu de`getSpecialCellsOrNullObject`.

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>Réduisez les cellules cibles avec les types de valeur de cellule

Il existe un paramètre secondaire optionnel, de type enum`Excel.SpecialCellValueType`, qui réduise encore la cellule à la cible. Vous pouvez l’utiliser uniquement lorsque vous passez soit «Formules» ou «Constantes» à`getSpecialCells`ou`getSpecialCellsOrNullObject`. Le paramètre spécifie que vous souhaitez uniquement les cellules avec certains types de valeurs. Il existe quatre types de base: «Erreur», «Logique»(ce qui signifie booléen), «Nombres», et «Texte». (L’enum dispose d’autres valeurs hormis les quatre traités plus haut.) Ce qui suit en est un exemple. Tenez compte du code suivant:

- Il surlignera uniquement les cellules qui disposent une valeur de nombre littérale. Il surlignera les cellules qui disposent une formule (même si le résultat est un nombre) ou un booléen, un texte ou des cellules d’instruction d’erreur.
- Pour tester le code, assurez-vous que la feuille de calcul dispose de certaines cellules avec des valeurs de nombre littérales, certaines avec d’autres sortes de valeurs littérales, et certaines avec des formules.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

Parfois, vous avez besoin d’exécuter plus d’un type de valeur de cellule, tel que toutes les cellules à valeur de texte et à valeur booléen («Logique»). L’enum`Excel.SpecialCellValueType` dispose de valeurs qui vous laisse combiner les types. Par exemple, «LogicalText» ciblera toutes les cellules à valeur texte et booléen. Vous pouvez combiner deux ou trois des quatre types de base. Les noms de ces valeurs d’enum qui combinent les types de base sont toujours par ordre alphabétique. Donc pour combiner les cellules à valeur d’erreur, texte et booléen, utilisez «ErrorLogicalText»,et non «LogicalErrorText» ou «TextErrorLogical». Le paramètre par défaut de «Tous» combine les quatre types. L’exemple suivant surligne toutes les cellules dotées de formules qui produisent les valeurs de nombre ou booléennes:

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> Le paramètre `Excel.SpecialCellValueType` peut uniquement être utilisé si le paramètre `Excel.SpecialCellType` est défini «Formules» ou «Constantes».

### <a name="get-rangeareas-within-rangeareas"></a>Obtenez RangeAreas dans RangeAreas

Le type `RangeAreas` lui-même dispose également de méthodes `getSpecialCells`et`getSpecialCellsOrNullObject` qui prennent les deux paramètres identiques. Ces méthodes renvoient toutes les cellules ciblées à partir des plages dans la collection`RangeAreas.areas`. Il existe une petite différence dans le comportement des méthodes lors de l’appel d’un objet`RangeAreas` au lieu d’un objet`Range`: lorsque vous passez «SameConditionalFormat» comme premier paramètre, la méthode renvoie toutes les cellules qui disposent la même mise en forme conditionnelle que la cellule en haut à gauche* dans la première plage dans la `RangeAreas.areas`collection*. Le même point s’applique à «SameDataValidation»:lors du passage à`Range.getSpecialCells`, il renvoie toutes les cellules qui disposent la même règle de validation de données comme la cellule en haut à gauche*dans la plage*. Mais lors du passage à `RangeAreas.getSpecialCells`, il renvoie toutes les cellules qui disposent la même règle de validation de données comme la cellule en haut à gauche*dans la plage`RangeAreas.areas`dans la collection*.

## <a name="read-properties-of-rangeareas"></a>Lire les propriétés de RangeAreas

La lecture des valeurs de propriété de `RangeAreas` nécessite un soin, car une propriété donnée peut avoir des valeurs différentes pour des plages différentes au sein du`RangeAreas`. La règle générales est que si une valeur consistante*peut*être renvoyée, elle sera renvoyée. Par exemple, dans le code suivant, le code RGB pour rose (`#FFC0CB`) et`true`sera connecté à la console car les deux plages dans l’objet`RangeAreas` dispose d’un remplissage rose et les deux sont des colonnes entières.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

Les choses se compliquent lorsque la consistance est impossible. Le comportement de propriétés`RangeAreas` suit ces trois principes:

- Une propriété booléenne d’un objet`RangeAreas` renvoie`false`à moins que la propriété soit vraie pour toutes les plages membres.
- Les propriétés non-booléennes, avec l’exception de la propriété`address`renvoie`null`à moins que la propriété correspondante sur toutes les plages membres dispose de la même valeur.
- La propriété`address`renvoie une chaîne délimitée à virgule des adresses des plages membres.

Par exemple, le code suivante crée un`RangeAreas`dans lequel seule une plage est une colonne entière et seule une est remplie de rose. La console s’affichera`null`pour un remplissage de couleur,`false`pour la propriété`isEntireRow` et «Sheet1!F3:F5, Sheet1!H:H» (en présumant que la feuille de calcule soit «Sheet1») pour la propriété`address`. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Objet de plage (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [Objet RangeAreas (JavaScript API pout Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Ce lien peut ne peut pas fonctionner lorsque l’API est en prévisualisation. Comme alternative, consultez [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)