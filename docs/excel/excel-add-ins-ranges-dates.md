---
title: Utiliser des dates à l’aide de Excel API JavaScript
description: Utilisez le plug-in Moment-MSDate avec l’API JavaScript Excel pour travailler avec les dates.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdfc39f12b3374d9903156b1ba71a9bbd4f296735f0ed41dac56d62243058c1d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084730"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>Utiliser des dates à l’aide Excel API JavaScript et Moment-MSDate plug-in

Cet article fournit des exemples de code qui montrent comment utiliser des dates à l’aide de l’API JavaScript Excel et du [plug-in Moment-MSDate.](https://www.npmjs.com/package/moment-msdate) Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en charge, voir `Range` [la Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a>Utiliser le plug-in Moment-MSDate pour travailler avec des dates

La[bibliothèque Moment JavaScript](https://momentjs.com/)fournit une manière pratique d’utiliser les dates et les horodateurs. Le[plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate)convertit le format des moments dans un préférable pour Excel. Il s’agit du même format que la[fonction NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)renvoie.

Le code suivant montre comment définir la plage **À B4** sur l’timestamp d’un moment.

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

L’exemple de code suivant illustre une technique similaire pour récupérer la date de la cellule et la convertir dans un format ou `Moment` un autre format.

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

Votre add-in doit mettre en forme les plages pour afficher les dates sous une forme plus lisible. Par exemple, `"[$-409]m/d/yy h:mm AM/PM;@"` affiche « 03/12/2018 15:57 ». Pour plus d’informations sur les formats de nombre de date et d’heure, voir « Recommandations en matière de formats de date et d’heure » dans l’article Révision pour la personnalisation d’un [format de](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) nombre.


## <a name="see-also"></a>Voir aussi

- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
