---
title: Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)
description: Les fonctions et scénarios d’objet de plage avancés, tels que les cellules spéciales, suppriment les doublons et utilisent des dates.
ms.date: 10/13/2020
localization_priority: Normal
ms.openlocfilehash: 144012177e0e070149f6cef825c63392a468773d
ms.sourcegitcommit: 6fa29989dfaec4dfa0f8df3fe5fb038d7afbae30
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/16/2020
ms.locfileid: "48487886"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a>Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)

Cet article génère des informations dans[ Utiliser des plages à l’aide de l’API JavaScript Excel (fondamental)](excel-add-ins-ranges.md) en fournissant les exemples de code qui affichent la manière d’exécuter plus de tâches avancées avec des plages à l’aide de l’API JavaScript Excel. Pour obtenir la liste complète des propriétés et des méthodes `Range` prises en charge par l’objet, reportez-vous à la rubrique [objet Range (interface API JavaScript pour Excel)](/javascript/api/excel/excel.range).

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a>Utiliser des dates à l’aide de plug-in Moment-MSDate

La[bibliothèque Moment JavaScript](https://momentjs.com/)fournit une manière pratique d’utiliser les dates et les horodateurs. Le[plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate)convertit le format des moments dans un préférable pour Excel. Il s’agit du même format que la[fonction NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)renvoie.

Le code suivant affiche la manière d’établir la plage à**B4**vers un horodateur du moment:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

Il s’agit d’une technique similaire pour récupérer la date de la cellule et la convertir en un moment ou autre format, comme démontré dans le code suivant:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

Votre complément devra mettre en forme les plages pour afficher les dates dans une forme plus lisible. L’exemple de`"[$-409]m/d/yy h:mm AM/PM;@"`affiche une heure comme «12/3/18 3:57 PM». Pour plus d’informations concernant les formats de date et d’heure , veuillez consulter les «Instructions relatifs aux formats de date et heure» dans l’article[ Instructions revoir afin de personnaliser le format numérique](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).

## <a name="work-with-multiple-ranges-simultaneously"></a>Travailler simultanément avec plusieurs plages

L’objet [RangeAreas](/javascript/api/excel/excel.rangeareas) permet à votre complément d’effectuer des opérations sur plusieurs plages à la fois. Ces plages peuvent être adjacentes, mais cela n’est pas obligatoire. `RangeAreas`sont abordés plus loin dans l’article[Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md).

## <a name="find-special-cells-within-a-range"></a>Rechercher des cellules spéciales dans une plage

Les méthodes [Range. getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) et [Range. getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) recherchent des plages basées sur les caractéristiques de leurs cellules et les types de valeurs de leurs cellules. Ces deux méthodes renvoient à des`RangeAreas`objets. Voici les signatures des méthodes à partir des types de fichiers de données TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

L’exemple suivant utilise la méthode`getSpecialCells`pour rechercher toutes les cellules contenant les formules. Tenez compte du code suivant:

- Cela limite la partie de la feuille qui nécessite d’être recherchée en appelant d’abord`Worksheet.getUsedRange`et en appelant`getSpecialCells`uniquement pour cette plage.
- La`getSpecialCells`méthode renvoie un`RangeAreas`objet, toutes les cellules alors dotées de formules seront colorées en rose même si elles ne sont pas adjacentes.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

Si aucune cellule avec la caractéristique ciblée n’existe dans la plage `getSpecialCells` lève une erreur**ItemNotFound**. Cela dévie le flux de contrôle vers un(e)`catch`bloc/méthode, s’il en existe. S’il n’existe pas de `catch` bloc, l’erreur interrompt la méthode.

Si vous attendez que des cellules avec la caractéristique ciblée existent toujours, vous souhaiterez probablement que votre code  lève une erreur si ces cellules ne sont pas là. Mais dans les scénarios où les cellules ne correspondent pas; votre code doit vérifier cette possibilité et le gérer gracieusement sans émettre d’erreur. Vous pouvez obtenir ce comportement avec la `getSpecialCellsOrNullObject`méthode et sa propriété renvoyée`isNullObject`. Cet exemple utilise les valeurs suivantes. Tenez compte du code suivant:

- La méthode`getSpecialCellsOrNullObject`renvoie toujours un objet proxy, donc il ne s’agit jamais du sens`null`JavaScript ordinaire. Mais si les cellules non correspondantes sont introuvables, la propriété`isNullObject` de l’objet est établi à`true`.
- Il appelle`context.sync`*avant*de tester la propriété`isNullObject`. Il s’agit d’une condition avec toutes les méthodes et propriétés`*OrNullObject`, car vous devez toujours télécharger et synchroniser une propriété afin de le lire.  Cependant, il n’est pas nécessaire de télécharger*de manière explicite*la propriété`isNullObject`. Il est automatiquement téléchargé par le`context.sync`même si`load`n’est pas appelé sur l’objet. Pour plus d’informations, consultez la rubrique [ \* OrNullObject Methods and Properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).
- Vous pouvez tester ce code en sélectionnant d’abord une plage qui n’a pas de cellules de formule et en l’exécutant. Puis sélectionnez une plage qui dispose au moins d’une cellule dotée d’une formule et en l’exécutant à nouveau.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
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

### <a name="narrow-the-target-cells-with-cell-value-types"></a>Réduisez les cellules cibles avec les types de valeur de cellule

Les méthodes`Range.getSpecialCells()` et `Range.getSpecialCellsOrNullObject()`acceptent un deuxième paramètre facultatif utilisé pour affiner davantage les cellules ciblées. Ce deuxième paramètre est un`Excel.SpecialCellValueType` que vous utilisez afin de spécifier que vous souhaitez uniquement les cellules qui contiennent certains types de valeurs.

> [!NOTE]
> Le paramètre `Excel.SpecialCellValueType` peut uniquement être utilisé si le paramètre `Excel.SpecialCellType` est défini sur `Excel.SpecialCellType.formulas`ou `Excel.SpecialCellType.constants`.

#### <a name="test-for-a-single-cell-value-type"></a>Test d’un type de valeur de cellule unique

Le `Excel.SpecialCellValueType` enum dispose de ces quatre types de base (outre les autres valeurs combinées décrites plus loin dans cette section):

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical` (ce qui signifie booléen)
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

L’exemple suivant recherche les cellules spéciaux qui sont des constantes numériques et colore les cellules en rose. Tenez compte du code suivant:

- Il surlignera uniquement les cellules qui disposent une valeur de nombre littérale. Il surlignera les cellules qui disposent une formule (même si le résultat est un nombre) ou un booléen, un texte ou des cellules d’instruction d’erreur.
- Pour tester le code, assurez-vous que la feuille de calcul dispose de certaines cellules avec des valeurs de nombre littérales, certaines avec d’autres sortes de valeurs littérales, et certaines avec des formules.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

#### <a name="test-for-multiple-cell-value-types"></a>Test d’un type de valeur de cellule multiple

Parfois, vous avez besoin d’exécuter plus d’un type de valeur de cellule, tel que toutes les cellules à valeur de texte et à valeur booléen (`Excel.SpecialCellValueType.logical`). Le `Excel.SpecialCellValueType` enum comporte des valeurs avec les types combinés. Par exemple,`Excel.SpecialCellValueType.logicalText`cible toutes les cellules à valeur texte et booléen. `Excel.SpecialCellValueType.all` est la valeur par défaut, ce qui ne limite pas les types de valeur de cellule renvoyés. L’exemple suivant surligne toutes les cellules dotées de formules qui produisent les valeurs de nombre ou booléennes.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="cut-copy-and-paste"></a>Couper, copier et coller

### <a name="copy-and-paste"></a>Copy and paste

La méthode [Range. CopyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) réplique les actions **copier** et **coller** de l’interface utilisateur Excel. L’objet plage sur lequel`copyFrom`est appelé est la destination. La source à copier est transmise en tant que plage ou qu’adresse de chaîne représentant une plage.

L’exemple de code suivant copie les données de la plage **A1:E1** dans la plage commençant en **G1** (ce qui aboutit à un collage dans la plage **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

`Range.copyFrom`dispose de trois paramètres facultatifs.

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` spécifie les données copiées de la source vers la destination.

- `Excel.RangeCopyType.formulas` transfère les formules dans les cellules sources et conserve le positionnement relatif des plages de ces formules. Les entrées autres que des formules sont copiées telles quelles.
- `Excel.RangeCopyType.values` copie les valeurs des données et, s’il s’agit d’une formule, le résultat de celle-ci.
- `Excel.RangeCopyType.formats` copie la mise en forme de la plage, y compris la police, la couleur et d’autres paramètres de mise en forme, mais aucune valeur.
- `Excel.RangeCopyType.all` (option par défaut) copie les données et la mise en forme, en conservant les formules, le cas échéant.

`skipBlanks` définit si les cellules vides sont copiées dans la destination. Quand la valeur est true, `copyFrom` ignore les cellules vides de la plage source.
Les cellules ignorées ne remplacent pas les données existantes dans les cellules correspondantes de la plage de destination. La valeur par défaut est false.

`transpose` détermine si les données sont ou non transposées, ce qui signifie que ses lignes et colonnes sont permutées dans l’emplacement source.
Une plage transposée est renversée le long de la diagonale principale, de sorte que les lignes **1**, **2** et **3** deviennent les colonnes **A**, **B** et **C**.

L’exemple de code et les images suivants illustrent ce comportement dans un scénario simple.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*Avant l’exécution de la fonction précédente.*

![Données dans Excel avant l’exécution de la méthode de copie de la plage](../images/excel-range-copyfrom-skipblanks-before.png)

*Après l’exécution de la fonction précédente.*

![Données dans Excel après exécution de la méthode de copie de la plage](../images/excel-range-copyfrom-skipblanks-after.png)

### <a name="cut-and-paste-move-cells"></a>Couper et coller (déplacer) des cellules

La méthode [Range. MoveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) déplace les cellules vers un nouvel emplacement dans le classeur. Ce comportement de déplacement de cellule fonctionne de la même manière que lorsque les cellules sont déplacées en [faisant glisser la bordure de la plage](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) ou lorsque vous effectuez des opérations **couper** - **coller** . La mise en forme et les valeurs de la plage sont déplacées vers l’emplacement spécifié en tant que `destinationRange` paramètre.

L’exemple de code suivant montre une plage déplacée avec la `Range.moveTo` méthode. Notez que si la plage de destination est plus petite que la source, elle sera étendue de façon à inclure le contenu source.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="remove-duplicates"></a>Supprimer les doublons

La méthode [Range. removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) supprime les lignes contenant des entrées en double dans les colonnes spécifiées. La méthode passe par chaque ligne de la plage comprise entre l’index à la valeur la plus faible et l’index de la valeur la plus haute dans la plage (du haut vers le bas). Une rangée est supprimée si une valeur dans sa/ses colonne(s) spécifiée(s) apparue(s) plus tôt dans la plage. Les rangées de la plage en-dessous de la rangée supprimée sont déplacées. `removeDuplicates` n’affecte pas la position des cellules en dehors de la rangée.

`removeDuplicates`prend un `number[]` représentant les indices de la colonne qui sont vérifiés pour les doublons. Ce tableau est à base zéro et lié à la rangée, et non à la feuille de calcul. La méthode prend également un paramètre Boolean qui spécifie si la première ligne est un en-tête. Lorsque**true**, la rangée du dessus est ignorée lorsque les doublons sont pris en considération. La `removeDuplicates` méthode renvoie un `RemoveDuplicatesResult` objet qui spécifie le nombre de lignes supprimées et le nombre de lignes uniques restantes.

Lorsque vous utilisez la méthode d’une plage `removeDuplicates` , gardez les points suivants à l’esprit :

- `removeDuplicates`considère les valeurs de cellule, et non les résultats de la fonction. Si deux fonctions différentes évaluent le même résultat, les valeurs de la cellule ne sont pas considérées comme doublons.
- Les cellules vides ne sont pas ignorées par`removeDuplicates`. La valeur d’une cellule vide est traitée comme toute autre valeur. Cela signifie que les rangées vides contenues au sein de la plage seront incluses dans le `RemoveDuplicatesResult`.

L’exemple suivant affiche la suppression des entrées avec des valeurs de doublons dans la première colonne.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*Avant l’exécution de la fonction précédente.*

![Données dans Excel avant l’exécution de la méthode de suppression des doublons de la plage](../images/excel-ranges-remove-duplicates-before.png)

*Après l’exécution de la fonction précédente.*

![Données dans Excel après exécution de la méthode Remove Duplicates de la plage](../images/excel-ranges-remove-duplicates-after.png)

## <a name="group-data-for-an-outline"></a>Données de groupe pour un plan

Les lignes ou les colonnes d’une plage peuvent être regroupées pour créer un [plan](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF). Ces groupes peuvent être réduits et développés pour masquer et afficher les cellules correspondantes. Cela facilite l’analyse rapide des données de haut niveau. Utilisez [Range. Group](/javascript/api/excel/excel.range#group-groupoption-) pour créer ces groupes de plan.

Un plan peut avoir une hiérarchie, où les groupes de plus petite taille sont imbriqués sous des groupes plus grands. Cela permet d’afficher le plan à différents niveaux. Vous pouvez modifier le niveau de plan visible par programmation à l’aide de la méthode [Worksheet. showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) . Notez qu’Excel ne prend en charge que huit niveaux de groupes de plan.

L’exemple de code suivant montre comment créer un plan avec deux niveaux de groupes pour les lignes et les colonnes. L’image suivante montre les regroupements de ce plan. Notez que dans l’exemple de code, les plages qui sont groupées n’incluent pas la ligne ou la colonne du contrôle de plan (« totaux » pour cet exemple). Un groupe définit ce qui sera réduit, pas la ligne ou la colonne avec le contrôle.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);

```

![Une plage avec un contour à deux niveaux et deux dimensions](../images/excel-outline.png)

Pour dissocier un groupe de lignes ou de colonnes, utilisez la méthode [Range. Ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) . Cette opération supprime le niveau le plus à l’extérieur du plan. Si plusieurs groupes du même type de ligne ou de colonne se trouvent au même niveau au sein de la plage spécifiée, tous ces groupes sont dissociés.

## <a name="handle-dynamic-arrays-and-spilling"></a>Gérer les tableaux dynamiques et les débordements

Certaines formules Excel renvoient des [tableaux dynamiques](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531). Ces valeurs remplissent les valeurs de plusieurs cellules en dehors de la cellule d’origine de la formule. Ce débordement de valeur est appelé « déversement ». Votre complément peut trouver la plage utilisée pour un déversement avec la méthode [Range. getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) . Il existe également une [version * OrNullObject](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject` .

L’exemple suivant montre une formule de base qui copie le contenu d’une plage dans une cellule, qui se transforme en cellules voisines. Le complément enregistre ensuite la plage qui contient le déversement.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

Vous pouvez également trouver la cellule responsable du débordement dans une cellule donnée à l’aide de la méthode [Range. getSpillParent](/javascript/api/excel/excel.range#getspillparent--) . Notez que `getSpillParent` ne fonctionne que si l’objet Range est une seule cellule. `getSpillParent`L’appel sur une plage comportant plusieurs cellules entraîne la levée d’une erreur (ou une plage null renvoyée pour `Range.getSpillParentOrNullObject` ).

## <a name="get-formula-precedents"></a>Obtenir les antécédents de la formule

Une formule Excel fait souvent référence à d’autres cellules. Lorsqu’une cellule fournit des données à une formule, il s’agit d’une formule « antécédent ». Pour en savoir plus sur les fonctionnalités Excel liées aux relations entre les cellules, consultez l’article [afficher les relations entre les formules et les cellules](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507) . 

Avec [Range. getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--), votre complément peut localiser les cellules précédentes d’une formule. `Range.getDirectPrecedents` renvoie un `WorkbookRangeAreas` objet. Cet objet contient les adresses de tous les antécédents du classeur. Il possède un `RangeAreas` objet distinct pour chaque feuille de calcul contenant au moins un précédent de formule. Pour plus d’informations sur l’utilisation de l’objet, voir [travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md) `RangeAreas` .

Dans l’interface utilisateur Excel, le bouton **repérer les antécédents** dessine une flèche de cellules précédentes à la formule sélectionnée. Contrairement au bouton de l’interface utilisateur Excel, la `getDirectPrecedents` méthode ne dessine pas de flèches. 

> [!IMPORTANT]
> La `getDirectPrecedents` méthode ne peut pas récupérer les cellules précédentes entre les classeurs. 

L’exemple suivant montre comment obtenir les antécédents directs de la plage active, puis comment modifier la couleur d’arrière-plan des cellules précédentes en jaune. 

> [!NOTE]
> La plage active doit contenir une formule qui fait référence à d’autres cellules du même classeur pour que la mise en surbrillance fonctionne correctement. 

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Voir aussi

- [Utiliser les plages à l’aide de l’API JavaScript Excel](excel-add-ins-ranges.md)
- [Modèle objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
