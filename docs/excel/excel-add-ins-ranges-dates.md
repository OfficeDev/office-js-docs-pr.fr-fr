---
title: Utiliser des dates à l’aide Excel API JavaScript
description: Utilisez le plug-in Moment-MSDate avec l’API JavaScript Excel pour travailler avec les dates.
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: becbbc9deb6f07e244ed0aac1f04b3dad1a800eb
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340567"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>Utiliser des dates à l’aide Excel API JavaScript et du plug-in Moment-MSDate'interface utilisateur

Cet article fournit des exemples de code qui montrent comment utiliser des dates à l’aide de l’API JavaScript Excel et du [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate). Pour obtenir la liste complète des propriétés et méthodes que `Range` l’objet prend en charge, voir [la Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a>Utiliser le plug-in Moment-MSDate pour travailler avec des dates

La[bibliothèque Moment JavaScript](https://momentjs.com/)fournit une manière pratique d’utiliser les dates et les horodateurs. Le[plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate)convertit le format des moments dans un préférable pour Excel. Il s’agit du même format que la[fonction NOW](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46)renvoie.

Le code suivant montre comment définir la plage **À B4** sur l’timestamp d’un moment.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let now = Date.now();
    let nowMoment = moment(now);
    let nowMS = nowMoment.toOADate();

    let dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    await context.sync();
});
```

L’exemple de code suivant illustre une technique similaire pour récupérer la date de la cellule et la convertir dans un `Moment` format ou un autre format.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let dateRange = sheet.getRange("B4");
    dateRange.load("values");

    await context.sync();

    let nowMS = dateRange.values[0][0];

    // Log the date as a moment.
    let nowMoment = moment.fromOADate(nowMS);
    console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

    // Log the date as a UNIX-style timestamp.
    let now = nowMoment.unix();
    console.log(`get (timestamp): ${now}`);
});
```

Votre add-in doit mettre en forme les plages pour afficher les dates sous une forme plus lisible. Par exemple, `"[$-409]m/d/yy h:mm AM/PM;@"` affiche « 03/12/2018 15:57 ». Pour plus d’informations sur les formats de nombre de date et d’heure, voir « Recommandations en matière de formats de date et d’heure » dans l’article Recommandations en matière de révision pour la personnalisation d’un [format de](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) nombre.

## <a name="see-also"></a>Voir aussi

- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
