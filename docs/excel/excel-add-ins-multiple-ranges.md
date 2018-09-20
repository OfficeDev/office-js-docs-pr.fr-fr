---
title: Travailler avec plusieurs plages simultanément dans les compléments Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: bcb14d1f4c015fe675c2d65cb5f1198d485dd4c5
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016457"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Travailler avec plusieurs plages simultanément dans les compléments Excel (Aperçu)

La bibliothèque JavaScript Excel permet à votre complément d'effectuer des opérations et de définir des propriétés sur plusieurs plages simultanément. Les plages n’ont pas à être contiguës. En plus de rendre votre code plus simple, cette méthode de définition de propriété s’exécute beaucoup plus rapidement que de définir la même propriété individuellement pour chacune des plages.

> [!NOTE]
> Les API décrites dans cet article nécessitent la **version Office 2016 Démarrer en un clic 1809 Build 10820.20000** ou une version ultérieure. (Vous devrez peut-être rejoindre le [programme Office Insider](https://products.office.com/office-insider) pour obtenir un build approprié). En outre, vous devez charger la version bêta de la bibliothèque JavaScript Office à partir du [CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Enfin, nous n’avons pas encore les pages de référence de ces API. Mais le fichier de type définition suivant comporte leurs descriptions : [office.d.ts bêta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).

## <a name="rangeareas"></a>RangeAreas

Un ensemble de plages (éventuellement discontinu) est représenté par un objet `Excel.RangeAreas`. Il possède des propriétés et méthodes similaires au type `Range` (beaucoup de noms identiques ou similaires), mais des ajustements ont été apportés :

- Aux types de données pour les propriétés et le comportement des méthodes setter et getter.
- Aux types de données des paramètres de méthode et des comportements de méthode.
- Les types de données des valeurs renvoyées par les méthodes.

Voici quelques exemples :

- `RangeAreas` a une propriété `address` qui retourne une chaîne délimitée par des virgules d’adresses de plage, au lieu d’une seule adresse comme avec la propriété `Range.address`.
- `RangeAreas` a une propriété `dataValidation` qui retourne un objet `DataValidation` qui représente la validation des données de toutes les plages dans le `RangeAreas`, si elle est cohérente. La propriété est `null` si des objets identiques `DataValidation` ne sont pas appliqués à toutes les plages dans le `RangeAreas`. Il s’agit d’un principe général, mais pas universel, pour l'objet `RangeAreas` : *Si une propriété ne dispose pas de valeurs cohérentes sur toutes les plages dans le `RangeAreas`, elle est `null`.* Pour plus d’informations et connaître certaines exceptions, voir [Propriétés de lecture de RangeAreas](#reading-properties-of-rangeareas).
- `RangeAreas.cellCount` obtient le nombre total de cellules dans toutes les plages dans le `RangeAreas`.
- `RangeAreas.calculate` recalcule les cellules de toutes les plages dans le `RangeAreas`.
- `RangeAreas.getEntireColumn` et `RangeAreas.getEntireRow` renvoient un autre objet `RangeAreas` qui représente toutes les colonnes (ou lignes) de toutes les plages dans le `RangeAreas`. Par exemple, si le `RangeAreas` représente « A1: C4 » et « F14:L15 », puis `RangeAreas.getEntireColumn` renvoie un objet `RangeAreas` qui représente « A:C » et « F:L ».
- `RangeAreas.copyFrom` peut adopter un paramètre `Range` ou un paramètre `RangeAreas` qui représente la ou les plages sources de l’opération de copie.

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>Liste complète des membres de Range qui sont également disponibles sur RangeAreas

##### <a name="properties"></a>Propriétés

Familiarisez-vous avec la [Lecture des propriétés de RangeAreas](#reading-properties-of-rangeareas) avant d’écrire le code qui lit les propriétés listées. Il existe quelques subtilités quant aux valeurs renvoyées.

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

Les méthodes de plage en préversion sont marquées comme telles.

- calculate()
- clear()
- convertDataTypeToText() (préversion)
- convertToLinkedDataType() (préversion)
- copyFrom() (préversion)
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange() (getOffsetRangeAreas nommé sur l’objet RangeAreas)
- getSpecialCells() (préversion)
- getSpecialCellsOrNullObject() (préversion)
- getTables() (préversion)
- getUsedRange() (getUsedRangeAreas nommé sur l’objet RangeAreas)
- getUsedRangeOrNullObject() (getUsedRangeAreasOrNullObject nommé sur l’objet RangeAreas)
- load()
- set()
- setDirty() (préversion)
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>Méthodes et propriétés spécifiques à RangeArea

Le type `RangeAreas` contient certaines propriétés et méthodes qui ne sont pas comprises dans l'objet `Range`. Vous trouverez ci-dessous une sélection ce celles-ci :

- `areas`: Un objet `RangeCollection` qui contient toutes les plages représentées par l'objet `RangeAreas`. L'objet `RangeCollection` est également nouveau et est similaire à d’autres objets de la collection Excel. Il possède une propriété `items` qui est un tableau des objets `Range` représentant les plages.
- `areaCount`: Nombre total de plages dans le `RangeAreas`.
- `getOffsetRangeAreas`: Fonctionne exactement comme [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), sauf qu’un `RangeAreas` est retourné et qu’il contient des plages qui sont décalées à partir d’une des plages dans le `RangeAreas` d’origine.

## <a name="create-rangeareas-and-set-properties"></a>Créer RangeAreas et définir les propriétés

Vous pouvez créer un objet `RangeAreas` de deux manières :

- Appelez `Worksheet.getRanges()` et lui passer une chaîne avec des adresses de plage délimitées par des virgules. Si une plage que vous souhaitez inclure a été créée dans un [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), vous pouvez inclure le nom, au lieu de l’adresse, dans la chaîne.
- Appelez `Workbook.getSelectedRanges()`. Cette méthode renvoie un `RangeAreas` représentant toutes les plages qui sont sélectionnées dans la feuille de calcul active.

Une fois que vous avez un objet `RangeAreas`, vous pouvez en créer d’autres à l’aide des méthodes de l’objet qui renvoient `RangeAreas` telles que `getOffsetRangeAreas` et `getIntersection`.

> [!NOTE]
> Vous ne pouvez pas ajouter directement des plages supplémentaires à un objet `RangeAreas`. Par exemple, la collection dans `RangeAreas.areas` ne possède pas de méthode `add`.


> [!WARNING] 
> N’essayez pas de directement ajouter ou supprimer des membres du tableau `RangeAreas.areas.items`. Cela entraîne un comportement indésirable dans votre code. Par exemple, il est possible d'acheminer un objet `Range` supplémentaire sur le tableau, mais cela peut provoquer des erreurs, car les propriétés et méthodes `RangeAreas` se comportent comme si le nouvel élément n’est pas là. Par exemple, la propriété `areaCount` n’inclut pas les plages poussées de cette manière et le `RangeAreas.getItemAt(index)` génère une erreur si `index` est supérieur(e) à `areasCount-1`. De même, la suppression d'un objet `Range` dans le tableau `RangeAreas.areas.items` en en en obtenant une référence et en appelant sa méthode `Range.delete` provoque des bogues : bien que l'`Range`objet*est* supprimé, les propriétés et méthodes de l'objet parent `RangeAreas` se comportent ou essayent de se comporter, comme s’il était toujours présent. Par exemple, si votre code appelle `RangeAreas.calculate`, Office tente de calculer la plage, mais une erreur se produit car l’objet range n’apparaît plus.

La définition d’une propriété sur une `RangeAreas` définit la propriété correspondante sur toutes les plages dans la collection `RangeAreas.areas`.

Voici un exemple de définition d’une propriété sur plusieurs plages. La fonction met en évidence les plages **F3:F5** et **H3:H5**.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Cet exemple s'applique aux scénarios dans lesquels vous pouvez coder en dur les adresses de la plage que vous passez à `getRanges` ou facilement les calculer lors de l’exécution. Voici quelques scénarios où cela est possible : 

- Le code s’exécute dans le contexte d’un modèle connu.
- Le code s’exécute dans le contexte de données importées pour lesquelles le schéma des données est connu.

Lorsque vous ne connaissez les plages sur lesquelles vous devez travailler au moment du codage, vous devez les découvrir lors de l’exécution. La section suivante décrit ces scénarios.

### <a name="discover-range-areas-programmatically"></a>Découvrir les zones de plages par programme

Les méthodes `Range.getSpecialCells()` et `Range.getSpecialCellsOrNullObject()` permettent de rechercher lors de l’exécution les plages que vous souhaitez utiliser en fonction des caractéristiques des cellules et du type des valeurs de cellules. Voici les signatures des méthodes obtenues à partir du fichier de types de données TypeScript :

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

Voici un exemple d’utilisation de la première. Tenez compte des informations suivantes :

- Il limite la partie de la feuille qui doit être recherchée en appelant d’abord `Worksheet.getUsedRange`, puis en appelant `getSpecialCells` pour cette plage seulement.
- Il passe en tant que paramètre à `getSpecialCells` la version de chaîne d’une valeur à partir de l'enum `Excel.SpecialCellType`. Certaines valeurs qui peuvent être passées sont : « Blanks » pour les cellules vides, « Constants » pour les cellules contenant des valeurs littérales au lieu de formules et « SameConditionalFormat » pour les cellules qui ont la même mise en forme conditionnelle que la première cellule de la `usedRange`. La première cellule est la cellule supérieure gauche. Pour une liste complète des valeurs dans l'enum, voir [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).
- La méthode `getSpecialCells` renvoie un objet `RangeAreas`, de sorte que toutes les cellules contenant des formules seront colorés en rose, même si elles ne sont pas toutes contiguës. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

Parfois la plage ne contient *aucune* cellule avec la caractéristique ciblée. Si `getSpecialCells` ne trouve aucune cellule, elle génère une erreur **ItemNotFound**. Cela redirige le flux de contrôle vers un bloc / une méthode `catch`, le cas échéant. S'il n’en existe pas, l’erreur arrête la fonction. Dans certains scénario, il est possible que vous vouliez justement qu'une erreur soit levée si aucune cellule avec la caractéristique ciblée n'existe. 

Mais dans les scénarios dans lesquels cela est normal, même rare, qu’aucune cellule ne corresponde ; votre code doit prendre en compte cette possibilité et la traiter correctement sans lever une erreur. Pour ces scénarios, utilisez la méthode `getSpecialCellsOrNullObject` et testez la propriété `RangeAreas.isNullObject`. Voici un exemple. Tenez compte des informations suivantes :

- La méthode `getSpecialCellsOrNullObject` retourne toujours un objet proxy, elle n’est donc jamais `null` dans le sens JavaScript ordinaire. Mais si aucune cellule n’est détectée, la propriété `isNullObject` de l’objet est définie sur `true`.
- Elle appelle `context.sync` *avant* qu’il teste la propriété `isNullObject`. Il s’agit d’une exigence de toutes les méthodes et propriétés `*OrNullObject`, car vous devez toujours charger et synchroniser une propriété afin de le lire. Toutefois, il n’est pas nécessaire de charger *explicitement* la propriété `isNullObject`. Elle est automatiquement chargée par le `context.sync` même si `load` n’est pas appelée sur l’objet. Pour plus d’informations, voir [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).
- Vous pouvez tester ce code en sélection d'abord plage qui n’a aucune cellule de formule et en l’exécutant. Sélectionnez ensuite une plage qui a au moins une cellule contenant une formule et exécutez-la à nouveau.

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

Par souci de simplicité, tous les autres exemples dans cet article utilisent la méthode `getSpecialCells` au lieu de `getSpecialCellsOrNullObject`.

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>Limiter les cellules cibles avec des types de valeur de cellule

Il existe un deuxième paramètre facultatif, de type enum `Excel.SpecialCellValueType`, qui permet de préciser encore le ciblage de cellules. Vous pouvez l’utiliser uniquement lorsque vous passez « Formulas » ou « Constants » à `getSpecialCells` ou `getSpecialCellsOrNullObject`. Le paramètre spécifie que vous voulez uniquement les cellules avec certains types de valeurs. Il existe quatre types de base : « Erreur », « Logique » (ce qui signifie booléenne), « Chiffres » et « Texte ». (L’enum possède d'autres valeurs en plus de ces quatre qui sont présentées ci-dessous). Voici un exemple. Tenez compte des informations suivantes :

- Elle mettra uniquement en surbrillance les cellules qui ont une valeur numérique littérale. Elle ne mettra pas en surbrillance les cellules qui contiennent une formule (même si le résultat est un nombre) ou une valeur booléenne, du texte ou les cellules d’état d’erreur.
- Pour tester le code, assurez-vous que la feuille de calcul contienne des cellules avec des valeurs littérales numériques, d'autres avec d'autres types de valeurs littérales et certaines des formules.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

Vous devez parfois travailler avec plus d’un type de valeur de cellule, telles qu'avec les cellules de texte et booléennes (« logical »). L'enum `Excel.SpecialCellValueType` possède des valeurs qui vous permettent de combiner des types. Par exemple, « LogicalText » ciblera toutes les cellules de type booléen et toutes les cellules de texte. Vous pouvez combiner deux ou trois des quatre types de base. Les noms de ces valeurs enum qui associent des types de base sont toujours dans l’ordre alphabétique. Pour combiner des cellules d’erreur, de texte et des cellules booléennes, utilisez donc « ErrorLogicalText », pas « LogicalErrorText » ou « TextErrorLogical ». Le paramètre par défaut de « All » regroupe les quatre types. L’exemple suivant met en évidence toutes les cellules contenant des formules qui génèrent numéro ou valeurs booléennes :

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
> Le paramètre `Excel.SpecialCellValueType` peut être utilisé uniquement si le paramètre `Excel.SpecialCellType` est « Formulas » ou « Constants ».

### <a name="get-rangeareas-within-rangeareas"></a>Obtenir RangeAreas dans RangeAreas

Le type `RangeAreas` lui-même contient également les méthodes `getSpecialCells` et `getSpecialCellsOrNullObject` qui comprennent les deux mêmes paramètres. Ces méthodes retournent toutes les cellules ciblées à partir de toutes les plages dans la collection `RangeAreas.areas`. Il existe une petite différence dans le comportement des méthodes lorsqu’elles sont appelées sur un objet `RangeAreas` au lieu d’un objet `Range` : lorsque vous passez « SameConditionalFormat » en tant que le premier paramètre, la méthode renvoie toutes les cellules qui ont la même mise en forme conditionnelle que la cellule de gauche supérieure *dans la première plage de la `RangeAreas.areas` collection*. Le même point s’applique à « SameDataValidation » : lorsqu’il est passé à `Range.getSpecialCells`, elle renvoie toutes les cellules qui ont la même règle de validation de données que la cellule supérieure gauche *de la plage*de cellules. Mais lorsqu’il est passé à `RangeAreas.getSpecialCells`, elle renvoie toutes les cellules qui ont la même règle de validation de données que la cellule supérieure gauche *dans la première plage dans la `RangeAreas.areas` collection*.

## <a name="read-properties-of-rangeareas"></a>Lire les propriétés de RangeAreas

Lire les valeurs de propriété de `RangeAreas` nécessite une attention particulière, car une propriété donnée peut avoir des valeurs différentes pour des plages différentes dans la `RangeAreas`. La règle générale est que si une valeur cohérente *peut* être retournée, elle le sera. Par exemple, dans le code suivant, le code RVB pour le rose (`#FFC0CB`) et `true` seront enregistrés dans la console, car à la fois les plages de l'objet `RangeAreas` ont un remplissage rose et les deux sont des colonnes entières.

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

Lorsque la cohérence n’est pas possible les choses deviennent plus complexes. Le comportement des propriétés `RangeAreas` suit ces trois principes :

- Une propriété de type booléen d'un objet`RangeAreas` renvoie `false`, sauf si la propriété a la valeur true pour toutes les plages de membre.
- Les propriétés non booléen, à l’exception de la propriété `address`, renvoient `null`, sauf si la propriété correspondante possède la même valeur sur toutes les plages de membre.
- La propriété `address` renvoie une chaîne délimitée par des virgules des adresses des plages de membre.

Par exemple, le code suivant crée un `RangeAreas` dans lequel une seule plage est une colonne entière et une seule est rempli en rose. La console affiche `null` pour la couleur de remplissage, `false` pour la propriété `isEntireRow` et « Sheet1!F3:F5, Sheet1!H:H» (en supposant que le nom de la feuille est « Sheet1 ») pour la propriété `address`. 

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

- [Concepts de base de l’API JavaScript pour Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [Objet Range (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [Objet RangeAreas (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Il est possible que ce lien ne fonctionne pas si l’API est en préversion. Comme alternative, voir [bêta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)