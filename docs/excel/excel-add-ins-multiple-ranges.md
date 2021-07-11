---
title: Travailler simultanément avec plusieurs plages dans des compléments Excel
description: Découvrez comment la Excel JavaScript permet à votre add-in d’effectuer des opérations et de définir des propriétés simultanément sur plusieurs plages.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 729b687b14beaeb74b329974bcca48dfd78bc11e
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349496"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins"></a>Travailler simultanément avec plusieurs plages dans des compléments Excel

La bibliothèque JavaScript Excel permet à votre complément d’effectuer des opérations et définir des propriétés, de manière simultanée sur plusieurs plages. Les plages n’ont pas besoin d’être contigus. En plus de rendre votre code plus simple, cette manière de paramétrer une propriété s’exécute beaucoup plus rapidement que paramétrer la même propriété pour chaque les plages de manière individuelle.

## <a name="rangeareas"></a>RangeAreas

Un ensemble de plages (éventuellement peuigues) est représenté par un [objet RangeAreas.](/javascript/api/excel/excel.rangeareas) Il possède des propriétés et des méthodes similaires au type`Range` (bon nombre des noms identiques ou similaires,), mais les ajustements ont été apportées à:

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

- `address`
- `addressLocal`
- `cellCount`
- `conditionalFormats`
- `context`
- `dataValidation`
- `format`
- `isEntireColumn`
- `isEntireRow`
- `style`
- `worksheet`

##### <a name="methods"></a>Méthodes

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- `getOffsetRange()` (nommé `getOffsetRangeAreas` sur `RangeAreas` l’objet)
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- `getUsedRange()` (nommé `getUsedRangeAreas` sur `RangeAreas` l’objet)
- `getUsedRangeOrNullObject()` (nommé `getUsedRangeAreasOrNullObject` sur `RangeAreas` l’objet)
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### <a name="rangearea-specific-properties-and-methods"></a>Méthodes et propriétés propres à une langue RangeArea

Le `RangeAreas` type possède des propriétés et des méthodes qui ne sont pas sur l’`Range`objet. Voici une sélection d’entre eux.

- `areas`: A`RangeCollection` objet qui contient toutes les plages représentées par l’ `RangeAreas`objet. L’`RangeCollection`objet est également nouveau et est semblable à d’autres objets de collection de sites Excel. Il possède une`items`propriété est une matrice d’`Range` objets représentant les plages.
- `areaCount`: Le nombre total de plages dans le`RangeAreas`.
- `getOffsetRangeAreas`: Fonctionne comme[Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), sauf qu’une `RangeAreas` est renvoyée et il contient des plages sont en décalage avec des plages du fichier d’origine`RangeAreas`.

## <a name="create-rangeareas"></a>Créer l’objet RangeAreas

Vous pouvez créer`RangeAreas`l’objet selon deux méthodes de base:

- Appeler`Worksheet.getRanges()` et de transmettre une chaîne comportant des adresses de plage séparées par des virgules. Si une plage que vous souhaitez inclure a été modifiée en[NamedItem](/javascript/api/excel/excel.nameditem), vous pouvez inclure le nom, au lieu de l’adresse, dans la chaîne.
- Appel `Workbook.getSelectedRanges()`. Cette méthode renvoie une`RangeAreas`représentation de toutes les plages qui sont sélectionnées sur la feuille de calcul active actuelle.

Une fois que vous avez un`RangeAreas`objet, vous pouvez en créer d’autres à l’aide des méthodes sur l’objet qui renvoie`RangeAreas`tel que`getOffsetRangeAreas`et`getIntersection`.

> [!NOTE]
> Vous ne pouvez pas ajouter directement des plages supplémentaires à un objet`RangeAreas`. Par exemple, la collection dans`RangeAreas.areas`n’a pas une méthode`add`.

> [!WARNING]
> N’essayez pas d’ajouter ou de supprimer directement les membres du tableau`RangeAreas.areas.items`. Cela mènera à un comportement indésirable dans votre code. Par exemple, il est possible de pousser un objet`Range` supplémentaire sur le tableau, mais ceci entrainera des erreurs car les propriétés`RangeAreas`et les méthodes se comportent comme si le nouvel élément n’existait pas. Par exemple, la propriété`areaCount`n’inclut pas les plages poussées de cette manière, et le `RangeAreas.getItemAt(index)` lance une erreur si`index`est plus grand que`areasCount-1`. De la même façon, supprimer un objet`Range`dans la plage`RangeAreas.areas.items`en obtenant une référence liée à celui-ci et en appelant sa méthode`Range.delete` entraîne des bogues: bien que `Range`l’objet *soit* supprimé, les propriétés et les méthodes de l’objet`RangeAreas`parent se comporte, ou tente de le faire, comme s’il existait encore. Par exemple, si votre code appelle`RangeAreas.calculate`, Office essaiera de calculer la plage, mais comprendra une erreur car l’objet de la plage n’est plus là.

## <a name="set-properties-on-multiple-ranges"></a>Définir les propriétés sur plusieurs plages

Paramétrer une propriété sur un objet `RangeAreas` établit une propriété correspondante sur toutes les plages dans la collection`RangeAreas.areas`.

Ce qui suit est un exemple de paramétrage d’une propriété sur des plages multiples. La fonction surligne les plages **F3:F5** and **H3:H5**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Cet exemple s’applique aux scénarios dans lesquels vous pouvez coder en dur les adresses de plage que vous passez à`getRanges`ou facilement les calculer à l’exécution. Certains des scénarios dans lesquels ceci peut s’appliquer incluent:

- Le code s’exécute dans le contexte d’un modèle connu.
- Le code s’exécute dans le contexte de données importées où le schéma des données est connu.

## <a name="get-special-cells-from-multiple-ranges"></a>Obtenir des cellules spéciales à partir de plusieurs plages

Les méthodes `getSpecialCells` et `getSpecialCellsOrNullObject` sur l’objet `RangeAreas` fonctionnent de manière analogue aux méthodes du même nom sur l’objet `Range`. Ces méthodes retournent les cellules disposant de la caractéristique spécifiée à partir de toutes les plages dans la collection `RangeAreas.areas`. Pour plus d’informations sur les cellules spéciales, voir [Rechercher des cellules spéciales dans une plage.](excel-add-ins-ranges-special-cells.md)

Lors de l’appel de la méthode `getSpecialCells` ou `getSpecialCellsOrNullObject` sur un objet `RangeAreas` :

- si vous passez `Excel.SpecialCellType.sameConditionalFormat` en tant que premier paramètre, la méthode renvoie toutes les cellules disposant de la même mise en forme conditionnelle que la cellule supérieure gauche de la première plage dans la collection `RangeAreas.areas`.
- Si vous passez `Excel.SpecialCellType.sameDataValidation` en tant que premier paramètre, la méthode renvoie toutes les cellules disposant de la même règle de validation des données que la cellule supérieure gauche de la première plage dans la collection `RangeAreas.areas`.

## <a name="read-properties-of-rangeareas"></a>Lire les propriétés de RangeAreas

La lecture des valeurs de propriété de `RangeAreas` nécessite un soin, car une propriété donnée peut avoir des valeurs différentes pour des plages différentes au sein du`RangeAreas`. La règle générales est que si une valeur consistante *peut* être renvoyée, elle sera renvoyée. Par exemple, dans le code suivant, le code RVB pour rose ( ) et sera enregistré dans la console car les deux plages de l’objet ont un remplissage rose et les deux sont des `#FFC0CB` `true` colonnes `RangeAreas` entières.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    var rangeAreas = sheet.getRanges("F:F, H:H");  
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
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H:H");

    var pinkColumnRange = sheet.getRange("H:H");
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

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Lire ou écrire dans une grande plage à l’aide de l Excel API JavaScript](excel-add-ins-ranges-large.md)
