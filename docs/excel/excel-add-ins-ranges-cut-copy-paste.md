---
title: Couper, copier et coller des plages à l’aide Excel API JavaScript
description: Découvrez comment couper, copier et coller des plages à l’aide Excel API JavaScript.
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 3d55e4d868a15c35ab9c68c799865560547e8188
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745103"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a>Couper, copier et coller des plages à l’aide Excel API JavaScript

Cet article fournit des exemples de code qui coupent, copient et collent des plages à l’aide Excel’API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que `Range` l’objet prend en charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a>Copy and paste

La [méthode Range.copyFrom](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1)) réplique **les actions** **Copier** et coller de l’interface Excel’utilisateur. La destination est l’objet `Range` qui `copyFrom` est appelé. La source à copier est transmise en tant que plage ou qu’adresse de chaîne représentant une plage.

L’exemple de code suivant copie les données de la plage **A1:E1** dans la plage commençant en **G1** (ce qui aboutit à un collage dans la plage **G1:K1**).

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    // Copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1").
    sheet.getRange("G1").copyFrom("A1:E1");
    await context.sync();
});
```

`Range.copyFrom`dispose de trois paramètres facultatifs.

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` spécifie les données copiées de la source vers la destination.

- `Excel.RangeCopyType.formulas` transfère les formules dans les cellules sources et conserve le positionnement relatif des plages de ces formules. Les entrées autres que des formules sont copiées telles quelles.
- `Excel.RangeCopyType.values` copie les valeurs des données et, s’il s’agit d’une formule, le résultat de celle-ci.
- `Excel.RangeCopyType.formats` copie la mise en forme de la plage, y compris la police, la couleur et d’autres paramètres de mise en forme, mais aucune valeur.
- `Excel.RangeCopyType.all` (option par défaut) copie les données et la mise en forme, en conservant les formules des cellules si elles sont trouvées.

`skipBlanks` définit si les cellules vides sont copiées dans la destination. Quand la valeur est true, `copyFrom` ignore les cellules vides de la plage source.
Les cellules ignorées ne remplacent pas les données existantes dans les cellules correspondantes de la plage de destination. La valeur par défaut est false.

`transpose` détermine si les données sont ou non transposées, ce qui signifie que ses lignes et colonnes sont permutées dans l’emplacement source.
Une plage transposée est renversée le long de la diagonale principale, de sorte que les lignes **1**, **2** et **3** deviennent les colonnes **A**, **B** et **C**.

L’exemple de code et les images suivants illustrent ce comportement dans un scénario simple.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    // Copy a range, omitting the blank cells so existing data is not overwritten in those cells.
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // Copy a range, including the blank cells which will overwrite existing data in the target cells.
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    await context.sync();
});
```

### <a name="data-before-range-is-copied-and-pasted"></a>Données avant que la plage ne soit copiée et copiée

![Données dans Excel avant l’exécutement de la méthode de copie de la plage.](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a>Données une fois la plage copiée et copiée

![Données dans Excel une fois que la méthode de copie de plage a été exécuté.](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a>Couper et coller (déplacer) des cellules

La [méthode Range.moveTo](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1)) déplace les cellules vers un nouvel emplacement dans le workbook. Ce comportement de déplacement de cellule fonctionne de la même manière que lorsque [](https://support.microsoft.com/office/803d65eb-6a3e-4534-8c6f-ff12d1c4139e) les cellules sont déplacées en faisant glisser la  bordure de la plage ou lors de l’action Couper **et** coller. La mise en forme et les valeurs de la plage sont déplacées vers l’emplacement spécifié en tant que `destinationRange` paramètre.

L’exemple de code suivant déplace une plage avec la `Range.moveTo` méthode. Notez que si la plage de destination est plus petite que la source, elle sera étendue pour englober le contenu source.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    await context.sync();
});
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Supprimer les doublons à l’aide Excel API JavaScript](excel-add-ins-ranges-remove-duplicates.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
