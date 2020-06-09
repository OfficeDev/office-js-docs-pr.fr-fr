---
title: Obtenir et définir la récurrence dans un complément Outlook
description: Cette rubrique vous explique comment utiliser l’API JavaScript Office pour obtenir et définir différentes propriétés de récurrence d’un élément dans un complément Outlook.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 6a50ba5eab39145d8e50a5a888a6ed0900200bc4
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44606454"
---
# <a name="get-and-set-recurrence"></a>Obtenir et définir la récurrence

Vous devez parfois créer et mettre à jour un rendez-vous périodique, par exemple une réunion hebdomadaire pour un projet d’équipe ou un rappel anniversaire annuel. Vous pouvez utiliser l’API JavaScript pour Office pour gérer les périodicités d’une série de rendez-vous dans votre complément.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1,7. Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="available-recurrence-patterns"></a>Modèles de récurrence disponibles

Pour configurer la récurrence, vous devez combiner les [types de récurrence](/javascript/api/outlook/office.mailboxenums.recurrencetype) et ses [propriétés de récurrence](/javascript/api/outlook/office.recurrenceproperties) applicables (le cas échéant).

**Tableau 1. Types de récurrence et leurs propriétés applicables**

|Type de récurrence|Propriétés de récurrence valide|Utilisation|
|---|---|---|
|`daily`|- [`interval`][interval link]|Un rendez-vous se produit tous les *intervalle* jours. Exemple : Un rendez-vous se produit tous les **_2_** jours.|
|`weekday`|Aucun.|Un rendez-vous se produit tous les jours de la semaine.|
|`monthly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]|- Un rendez-vous a lieu le *dayOfMonth* de chaque *intervalle* mois. Exemple : Un rendez-vous se produit tous les **_5_** du mois**_4_**.<br/><br/>- Un rendez-vous a lieu le *dayOfWeek* de la semaine *weekNumber* de chaque mois*intervalle*. Exemple : Un rendez-vous se produit tous les **_jeudis_** **_3_** tous les **_2_** mois.|
|`weekly`|- [`interval`][interval link]<br/>- [`days`][days link]|Un rendez-vous se produit chaque *jours*toutes les *intervalle*semaines. Exemple : Un rendez-vous se produit chaque **_mardi_ and _jeudi_** toutes les **_2_** semaines.|
|`yearly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]<br/>- [`month`][month link]|- Un rendez-vous a lieu le *dayOfMonth* de chaque *intervalle* mois tous les *intervalle* ans. Exemple : Un rendez-vous se produit tous les **_7_** du mois**_septembre_** tous les **_4_** ans.<br/><br/>- Un rendez-vous a lieu le *dayOfWeek* de la semaine *weekNumber* de chaque*mois* tous les *intervalle* ans. Exemple : Un rendez-vous se produit tous les **_1er_** **_jeudi_** du mois**_Septembre_** tous les **_2_** ans.|

> [!NOTE]
> Vous pouvez également utiliser la [ `firstDayOfWeek` ][firstDayOfWeek link] `weekly` propriété avec le  type de récurrence. Le jour spécifié commencera la liste des jours affichés dans la boîte de dialogue Récurrence.

## <a name="access-recurrence"></a>Accéder à la récurrence

Comment vous accédez à la récurrence et ce que vous pouvez en faire dépend de si vous êtes l’organisateur de rendez-vous ou un participant.

**Tableau 2. États de rendez-vous applicables**

|État de rendez-vous|La récurrence est-elle modifiable ?|La récurrence est-elle visible ?|
|---|---|---|
|Organisateur de rendez-vous - séries composer|Oui ([`setAsync`][setAsync link])|Oui ([`getAsync`][getAsync link])|
|Organisateur de rendez-vous - instance composer|Non (`setAsync` renvoie une erreur)|Oui ([`getAsync`][getAsync link])|
|Participant rendez-vous - séries lire|Non (`setAsync` non disponible)|Oui ([`item.recurrence`][item.recurrence link])|
|Participant rendez-vous - instance lire|Non (`setAsync` non disponible)|Oui ([`item.recurrence`][item.recurrence link])|
|Demande de réunion - série lire|Non (`setAsync` non disponible)|Oui ([`item.recurrence`][item.recurrence link])|
|Demande de réunion - instance lire|Non (`setAsync` non disponible)|Oui ([`item.recurrence`][item.recurrence link])|

## <a name="set-recurrence-as-the-organizer"></a>Configurer la récurrence en tant qu’organisateur

Tout comme le modèle de récurrence, vous devez également déterminer les dates de début et de fin et heures de vos séries de rendez-vous. L' [`SeriesTime`][SeriesTime link] objet est utilisé pour gérer ces informations.

L’organisateur de rendez-vous peut spécifier la récurrence pour une série de rendez-vous dans le mode Composer uniquement. Dans l’exemple suivant, la série de rendez-vous est définie comme se produisant de 10 h 30 à 11 h 00 PST chaque mardi et jeudi dans la période du 2 novembre 2019 au 2 décembre 2019.

```js
var seriesTimeObject = new Office.SeriesTime();
seriesTimeObject.setStartDate(2019,10,2);
seriesTimeObject.setEndDate(2019,11,2);
seriesTimeObject.setStartTime(10,30);
seriesTimeObject.setDuration(30);

var pattern = {
    "seriesTime": seriesTimeObject,
    "recurrenceType": "weekly",
    "recurrenceProperties": {"interval": 1, "days": ["tue", "thu"]},
    "recurrenceTimeZone": {"name": "Pacific Standard Time"}};

Office.context.mailbox.item.recurrence.setAsync(pattern, callback);

function callback(asyncResult)
{
    console.log(JSON.stringify(asyncResult));
}
```

## <a name="get-recurrence"></a>Obtenir la récurrence

### <a name="get-recurrence-as-the-organizer"></a>Obtenir la récurrence en tant qu’organisateur

Dans l’exemple suivant, dans le mode composer, l’organisateur de rendez-vous obtient l’objet de récurrence d’une série de rendez-vous ou une instance de ces séries.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult){
    var context = asyncResult.context;
    var recurrence = asyncResult.value;

    if (recurrence == null) {
        console.log("Non-recurring meeting");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

L’exemple suivant montre les résultats de l’appel `getAsync` qui récupère la récurrence d’une série.

> [!NOTE]
> Dans cet exemple, `seriesTimeObject` est un espace réservé pour JSON représentant la `recurrence.seriesTime` propriété. Vous devez utiliser les [`SeriesTime`][SeriesTime link] méthodes pour obtenir les propriétés de date et d’heure de périodicité.

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-recurrence-as-an-attendee"></a>Obtenir la récurrence en tant que participant

Dans l’exemple suivant, dans le mode composer, le participant au rendez-vous peut obtenir l’objet de récurrence d’une série de rendez-vous, une instance de ces séries, ou une demande de réunion.

```js
outputRecurrence(Office.context.mailbox.item);

function outputRecurrence(item) {
    var recurrence = item.recurrence;
    var seriesId = item.seriesId;

    if (recurrence == null) {
        console.log("Non-recurring item");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

L’exemple suivant montre la valeur de la `item.recurrence` propriété pour une série de rendez-vous.

> [!NOTE]
> Dans cet exemple, `seriesTimeObject` est un espace réservé pour JSON représentant la `recurrence.seriesTime` propriété. Vous devez utiliser les [`SeriesTime`][SeriesTime link] méthodes pour obtenir les propriétés de date et d’heure de périodicité.

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-the-recurrence-details"></a>Obtenir les détails de récurrence

Une fois que vous avez récupéré l’objet récurrence (soit à partir du `getAsync` rappel ou à partir de `item.recurrence`), vous pouvez obtenir les propriétés spécifiques de la récurrence. Par exemple, vous pouvez accéder aux dates de début et de fin et heures de la série via [méthodes][SeriesTime link] `recurrence.seriesTime` sur la  propriété.

```js
// Get series date and time info
var seriesTime = recurrence.seriesTime;
var startTime = recurrence.seriesTime.getStartTime();
var endTime = recurrence.seriesTime.getEndTime();
var startDate = recurrence.seriesTime.getStartDate();
var endDate = recurrence.seriesTime.getEndDate();
var duration = recurrence.seriesTime.getDuration();

// Get series time zone
var timeZone = recurrence.recurrenceTimeZone;

// Get recurrence properties
var recurrenceProperties = recurrence.recurrenceProperties;

// Get recurrence type
var recurrenceType = recurrence.recurrenceType;
```

## <a name="see-also"></a>Voir aussi

[Événement RecurrenceChanged](/javascript/api/office/office.eventtype)

[getAsync link]: /javascript/api/outlook/office.recurrence#getasync-options--callback-
[item.recurrence link]: ../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setasync-recurrencepattern--options--callback-

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayofmonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayofweek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstdayofweek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weeknumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
