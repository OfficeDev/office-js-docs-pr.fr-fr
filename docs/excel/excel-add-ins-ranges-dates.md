---
title: Utiliser des dates à l’aide de l’API JavaScript pour Excel
description: Utilisez le Moment-MSDate plug-in avec l’API JavaScript pour Excel pour utiliser les dates.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d3f59e5daad042541bd933fb4e644d40f27a6e5e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652852"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a><span data-ttu-id="2c430-103">Utiliser des dates à l’aide de l’API JavaScript pour Excel et Moment-MSDate plug-in</span><span class="sxs-lookup"><span data-stu-id="2c430-103">Work with dates using the Excel JavaScript API and the Moment-MSDate plug-in</span></span>

<span data-ttu-id="2c430-104">Cet article fournit des exemples de code qui montrent comment utiliser des dates à l’aide de l’API JavaScript pour Excel et du [plug-in Moment-MSDate.](https://www.npmjs.com/package/moment-msdate)</span><span class="sxs-lookup"><span data-stu-id="2c430-104">This article provides code samples that show how to work with dates using the Excel JavaScript API and the [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate).</span></span> <span data-ttu-id="2c430-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la [classe Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="2c430-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a><span data-ttu-id="2c430-106">Utiliser le plug-in Moment-MSDate pour travailler avec des dates</span><span class="sxs-lookup"><span data-stu-id="2c430-106">Use the Moment-MSDate plug-in to work with dates</span></span>

<span data-ttu-id="2c430-107">La[bibliothèque Moment JavaScript](https://momentjs.com/)fournit une manière pratique d’utiliser les dates et les horodateurs.</span><span class="sxs-lookup"><span data-stu-id="2c430-107">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="2c430-108">Le[plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate)convertit le format des moments dans un préférable pour Excel.</span><span class="sxs-lookup"><span data-stu-id="2c430-108">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="2c430-109">Il s’agit du même format que la[fonction NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)renvoie.</span><span class="sxs-lookup"><span data-stu-id="2c430-109">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="2c430-110">Le code suivant montre comment définir la plage **À B4** sur l’timestamp d’un moment.</span><span class="sxs-lookup"><span data-stu-id="2c430-110">The following code shows how to set the range at **B4** to a moment's timestamp.</span></span>

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

<span data-ttu-id="2c430-111">L’exemple de code suivant illustre une technique similaire pour récupérer la date de la cellule et la convertir dans un format ou `Moment` un autre format.</span><span class="sxs-lookup"><span data-stu-id="2c430-111">The following code sample demonstrates a similar technique to get the date back out of the cell and convert it to a `Moment` or other format.</span></span>

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

<span data-ttu-id="2c430-112">Votre add-in doit mettre en forme les plages pour afficher les dates sous une forme plus lisible.</span><span class="sxs-lookup"><span data-stu-id="2c430-112">Your add-in has to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="2c430-113">Par exemple, `"[$-409]m/d/yy h:mm AM/PM;@"` affiche « 03/12/2018 15:57 ».</span><span class="sxs-lookup"><span data-stu-id="2c430-113">For example, `"[$-409]m/d/yy h:mm AM/PM;@"` displays "12/3/18 3:57 PM".</span></span> <span data-ttu-id="2c430-114">Pour plus d’informations sur les formats de nombre de date et d’heure, voir « Recommandations en matière de formats de date et d’heure » dans l’article Révision pour la personnalisation d’un [format de](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) nombre.</span><span class="sxs-lookup"><span data-stu-id="2c430-114">For more information about date and time number formats, see "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>


## <a name="see-also"></a><span data-ttu-id="2c430-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2c430-115">See also</span></span>

- [<span data-ttu-id="2c430-116">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="2c430-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="2c430-117">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="2c430-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="2c430-118">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="2c430-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
