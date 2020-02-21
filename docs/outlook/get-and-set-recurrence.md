---
title: Obtenir et définir la récurrence dans un complément Outlook
description: Cette rubrique vous explique comment utiliser l’API JavaScript Office pour obtenir et définir différentes propriétés de récurrence d’un élément dans un complément Outlook.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: cc7160ed8fe82a02ced9c03bab181df57e66bb54
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166228"
---
# <a name="get-and-set-recurrence"></a><span data-ttu-id="5065e-103">Obtenir et définir la récurrence</span><span class="sxs-lookup"><span data-stu-id="5065e-103">Get and set recurrence</span></span>

<span data-ttu-id="5065e-104">Vous devez parfois créer et mettre à jour un rendez-vous périodique, par exemple une réunion hebdomadaire pour un projet d’équipe ou un rappel anniversaire annuel.</span><span class="sxs-lookup"><span data-stu-id="5065e-104">Sometimes you need to create and update a recurring appointment, such as a weekly status meeting for a team project or a yearly birthday reminder.</span></span> <span data-ttu-id="5065e-105">Vous pouvez utiliser l’API JavaScript pour Office pour gérer les modèles de périodicité d’une série de rendez-vous dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="5065e-105">You can use the JavaScript API for Office to manage the recurrence patterns of an appointment series in your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="5065e-106">La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1,7.</span><span class="sxs-lookup"><span data-stu-id="5065e-106">Support for this feature was introduced in requirement set 1.7.</span></span> <span data-ttu-id="5065e-107">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="5065e-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-recurrence-patterns"></a><span data-ttu-id="5065e-108">Modèles de récurrence disponibles</span><span class="sxs-lookup"><span data-stu-id="5065e-108">Available recurrence patterns</span></span>

<span data-ttu-id="5065e-109">Pour configurer la récurrence, vous devez combiner les [types de récurrence](/javascript/api/outlook/office.mailboxenums.recurrencetype) et ses [propriétés de récurrence](/javascript/api/outlook/office.recurrenceproperties) applicables (le cas échéant).</span><span class="sxs-lookup"><span data-stu-id="5065e-109">To configure the recurrence pattern, you need to combine the [recurrence type](/javascript/api/outlook/office.mailboxenums.recurrencetype) and its applicable [recurrence properties](/javascript/api/outlook/office.recurrenceproperties) (if any).</span></span>

<span data-ttu-id="5065e-110">**Tableau 1. Types de récurrence et leurs propriétés applicables**</span><span class="sxs-lookup"><span data-stu-id="5065e-110">**Table 1. Recurrence types and their applicable properties**</span></span>

|<span data-ttu-id="5065e-111">Type de récurrence</span><span class="sxs-lookup"><span data-stu-id="5065e-111">Recurrence type</span></span>|<span data-ttu-id="5065e-112">Propriétés de récurrence valide</span><span class="sxs-lookup"><span data-stu-id="5065e-112">Valid recurrence properties</span></span>|<span data-ttu-id="5065e-113">Utilisation</span><span class="sxs-lookup"><span data-stu-id="5065e-113">Usage</span></span>|
|---|---|---|
|`daily`|- [`interval`][interval link]|<span data-ttu-id="5065e-114">Un rendez-vous se produit tous les *intervalle* jours.</span><span class="sxs-lookup"><span data-stu-id="5065e-114">An appointment occurs every *interval* days.</span></span> <span data-ttu-id="5065e-115">Exemple : Un rendez-vous se produit tous les **_2_** jours.</span><span class="sxs-lookup"><span data-stu-id="5065e-115">Example: An appointment occurs every **_2_** days.</span></span>|
|`weekday`|<span data-ttu-id="5065e-116">Aucun.</span><span class="sxs-lookup"><span data-stu-id="5065e-116">None.</span></span>|<span data-ttu-id="5065e-117">Un rendez-vous se produit tous les jours de la semaine.</span><span class="sxs-lookup"><span data-stu-id="5065e-117">An appointment occurs every weekday.</span></span>|
|`monthly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]|<span data-ttu-id="5065e-118">- Un rendez-vous a lieu le *dayOfMonth* de chaque *intervalle* mois.</span><span class="sxs-lookup"><span data-stu-id="5065e-118">- An appointment occurs on day *dayOfMonth* every *interval* months.</span></span> <span data-ttu-id="5065e-119">Exemple : Un rendez-vous se produit tous les **_5_** du mois**_4_**.</span><span class="sxs-lookup"><span data-stu-id="5065e-119">Example: An appointment occurs on day **_5_** every **_4_** months.</span></span><br/><br/><span data-ttu-id="5065e-120">- Un rendez-vous a lieu le *dayOfWeek* de la semaine *weekNumber* de chaque mois*intervalle*.</span><span class="sxs-lookup"><span data-stu-id="5065e-120">- An appointment occurs on the *weekNumber* *dayOfWeek* every *interval* months.</span></span> <span data-ttu-id="5065e-121">Exemple : Un rendez-vous se produit tous les **_jeudis_** **_3_** tous les **_2_** mois.</span><span class="sxs-lookup"><span data-stu-id="5065e-121">Example: An appointment occurs on the **_third_** **_Thursday_** every **_2_** months.</span></span>|
|`weekly`|- [`interval`][interval link]<br/>- [`days`][days link]|<span data-ttu-id="5065e-122">Un rendez-vous se produit chaque *jours*toutes les *intervalle*semaines.</span><span class="sxs-lookup"><span data-stu-id="5065e-122">An appointment occurs on *days* every *interval* weeks.</span></span> <span data-ttu-id="5065e-123">Exemple : Un rendez-vous se produit chaque **_mardi_ and _jeudi_** toutes les **_2_** semaines.</span><span class="sxs-lookup"><span data-stu-id="5065e-123">Example: An appointment occurs on **_Tuesday_ and _Thursday_** every **_2_** weeks.</span></span>|
|`yearly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]<br/>- [`month`][month link]|<span data-ttu-id="5065e-124">- Un rendez-vous a lieu le *dayOfMonth* de chaque *intervalle* mois tous les *intervalle* ans.</span><span class="sxs-lookup"><span data-stu-id="5065e-124">- An appointment occurs on day *dayOfMonth* of *month* every *interval* years.</span></span> <span data-ttu-id="5065e-125">Exemple : Un rendez-vous se produit tous les **_7_** du mois**_septembre_** tous les **_4_** ans.</span><span class="sxs-lookup"><span data-stu-id="5065e-125">Example: An appointment occurs on day **_7_** of **_September_** every **_4_** years.</span></span><br/><br/><span data-ttu-id="5065e-126">- Un rendez-vous a lieu le *dayOfWeek* de la semaine *weekNumber* de chaque*mois* tous les *intervalle* ans.</span><span class="sxs-lookup"><span data-stu-id="5065e-126">- An appointment occurs on the *weekNumber* *dayOfWeek* of *month* every *interval* years.</span></span> <span data-ttu-id="5065e-127">Exemple : Un rendez-vous se produit tous les **_1er_** **_jeudi_** du mois**_Septembre_** tous les **_2_** ans.</span><span class="sxs-lookup"><span data-stu-id="5065e-127">Example: An appointment occurs on the **_first_** **_Thursday_** of **_September_** every **_2_** years.</span></span>|

> [!NOTE]
> <span data-ttu-id="5065e-128">Vous pouvez également utiliser la [ `firstDayOfWeek` ][firstDayOfWeek link] `weekly` propriété avec le  type de récurrence.</span><span class="sxs-lookup"><span data-stu-id="5065e-128">You can also use the [`firstDayOfWeek`][firstDayOfWeek link] property with the `weekly` recurrence type.</span></span> <span data-ttu-id="5065e-129">Le jour spécifié commencera la liste des jours affichés dans la boîte de dialogue Récurrence.</span><span class="sxs-lookup"><span data-stu-id="5065e-129">The specified day will start the list of days displayed in the recurrence dialog.</span></span>

## <a name="access-recurrence"></a><span data-ttu-id="5065e-130">Accéder à la récurrence</span><span class="sxs-lookup"><span data-stu-id="5065e-130">Access recurrence</span></span>

<span data-ttu-id="5065e-131">Comment vous accédez à la récurrence et ce que vous pouvez en faire dépend de si vous êtes l’organisateur de rendez-vous ou un participant.</span><span class="sxs-lookup"><span data-stu-id="5065e-131">How you access the recurrence pattern and what you can do with it depends on if you're the appointment organizer or an attendee.</span></span>

<span data-ttu-id="5065e-132">**Tableau 2. États de rendez-vous applicables**</span><span class="sxs-lookup"><span data-stu-id="5065e-132">**Table 2. Applicable appointment states**</span></span>

|<span data-ttu-id="5065e-133">État de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="5065e-133">Appointment state</span></span>|<span data-ttu-id="5065e-134">La récurrence est-elle modifiable ?</span><span class="sxs-lookup"><span data-stu-id="5065e-134">Is recurrence editable?</span></span>|<span data-ttu-id="5065e-135">La récurrence est-elle visible ?</span><span class="sxs-lookup"><span data-stu-id="5065e-135">Is recurrence viewable?</span></span>|
|---|---|---|
|<span data-ttu-id="5065e-136">Organisateur de rendez-vous - séries composer</span><span class="sxs-lookup"><span data-stu-id="5065e-136">Appointment organizer - compose series</span></span>|<span data-ttu-id="5065e-137">Oui ([`setAsync`][setAsync link])</span><span class="sxs-lookup"><span data-stu-id="5065e-137">Yes ([`setAsync`][setAsync link])</span></span>|<span data-ttu-id="5065e-138">Oui ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="5065e-138">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="5065e-139">Organisateur de rendez-vous - instance composer</span><span class="sxs-lookup"><span data-stu-id="5065e-139">Appointment organizer - compose instance</span></span>|<span data-ttu-id="5065e-140">Non (`setAsync` renvoie une erreur)</span><span class="sxs-lookup"><span data-stu-id="5065e-140">No (`setAsync` returns an error)</span></span>|<span data-ttu-id="5065e-141">Oui ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="5065e-141">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="5065e-142">Participant rendez-vous - séries lire</span><span class="sxs-lookup"><span data-stu-id="5065e-142">Appointment attendee - read series</span></span>|<span data-ttu-id="5065e-143">Non (`setAsync` non disponible)</span><span class="sxs-lookup"><span data-stu-id="5065e-143">No (`setAsync` not available)</span></span>|<span data-ttu-id="5065e-144">Oui ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="5065e-144">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="5065e-145">Participant rendez-vous - instance lire</span><span class="sxs-lookup"><span data-stu-id="5065e-145">Appointment attendee - read instance</span></span>|<span data-ttu-id="5065e-146">Non (`setAsync` non disponible)</span><span class="sxs-lookup"><span data-stu-id="5065e-146">No (`setAsync` not available)</span></span>|<span data-ttu-id="5065e-147">Oui ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="5065e-147">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="5065e-148">Demande de réunion - série lire</span><span class="sxs-lookup"><span data-stu-id="5065e-148">Meeting request - read series</span></span>|<span data-ttu-id="5065e-149">Non (`setAsync` non disponible)</span><span class="sxs-lookup"><span data-stu-id="5065e-149">No (`setAsync` not available)</span></span>|<span data-ttu-id="5065e-150">Oui ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="5065e-150">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="5065e-151">Demande de réunion - instance lire</span><span class="sxs-lookup"><span data-stu-id="5065e-151">Meeting request - read instance</span></span>|<span data-ttu-id="5065e-152">Non (`setAsync` non disponible)</span><span class="sxs-lookup"><span data-stu-id="5065e-152">No (`setAsync` not available)</span></span>|<span data-ttu-id="5065e-153">Oui ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="5065e-153">Yes ([`item.recurrence`][item.recurrence link])</span></span>|

## <a name="set-recurrence-as-the-organizer"></a><span data-ttu-id="5065e-154">Configurer la récurrence en tant qu’organisateur</span><span class="sxs-lookup"><span data-stu-id="5065e-154">Set recurrence as the organizer</span></span>

<span data-ttu-id="5065e-155">Tout comme le modèle de récurrence, vous devez également déterminer les dates de début et de fin et heures de vos séries de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="5065e-155">Along with the recurrence pattern, you also need to determine the start and end dates and times of your appointment series.</span></span> <span data-ttu-id="5065e-156">L' [`SeriesTime`][SeriesTime link] objet est utilisé pour gérer ces informations.</span><span class="sxs-lookup"><span data-stu-id="5065e-156">The [`SeriesTime`][SeriesTime link] object is used to manage that information.</span></span>

<span data-ttu-id="5065e-157">L’organisateur de rendez-vous peut spécifier la récurrence pour une série de rendez-vous dans le mode Composer uniquement.</span><span class="sxs-lookup"><span data-stu-id="5065e-157">The appointment organizer can specify the recurrence pattern for an appointment series in compose mode only.</span></span> <span data-ttu-id="5065e-158">Dans l’exemple suivant, la série de rendez-vous est définie comme se produisant de 10 h 30 à 11 h 00 PST chaque mardi et jeudi dans la période du 2 novembre 2019 au 2 décembre 2019.</span><span class="sxs-lookup"><span data-stu-id="5065e-158">In the following example, the appointment series is set to occur from 10:30 AM to 11:00 AM PST every Tuesday and Thursday during the period November 2, 2019 to December 2, 2019.</span></span>

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

## <a name="get-recurrence"></a><span data-ttu-id="5065e-159">Obtenir la récurrence</span><span class="sxs-lookup"><span data-stu-id="5065e-159">Get recurrence</span></span>

### <a name="get-recurrence-as-the-organizer"></a><span data-ttu-id="5065e-160">Obtenir la récurrence en tant qu’organisateur</span><span class="sxs-lookup"><span data-stu-id="5065e-160">Get recurrence as the organizer</span></span>

<span data-ttu-id="5065e-161">Dans l’exemple suivant, dans le mode composer, l’organisateur de rendez-vous obtient l’objet de récurrence d’une série de rendez-vous ou une instance de ces séries.</span><span class="sxs-lookup"><span data-stu-id="5065e-161">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series.</span></span>

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

<span data-ttu-id="5065e-162">L’exemple suivant montre les résultats de l’appel `getAsync` qui récupère la récurrence d’une série.</span><span class="sxs-lookup"><span data-stu-id="5065e-162">The following example shows the results of the `getAsync` call that retrieves the recurrence for a series.</span></span>

> [!NOTE]
> <span data-ttu-id="5065e-163">Dans cet exemple, `seriesTimeObject` est un espace réservé pour JSON représentant la `recurrence.seriesTime` propriété.</span><span class="sxs-lookup"><span data-stu-id="5065e-163">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="5065e-164">Vous devez utiliser les [`SeriesTime`][SeriesTime link] méthodes pour obtenir les propriétés de date et d’heure de périodicité.</span><span class="sxs-lookup"><span data-stu-id="5065e-164">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-recurrence-as-an-attendee"></a><span data-ttu-id="5065e-165">Obtenir la récurrence en tant que participant</span><span class="sxs-lookup"><span data-stu-id="5065e-165">Get recurrence as an attendee</span></span>

<span data-ttu-id="5065e-166">Dans l’exemple suivant, dans le mode composer, le participant au rendez-vous peut obtenir l’objet de récurrence d’une série de rendez-vous, une instance de ces séries, ou une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="5065e-166">In the following example, an appointment attendee can get the recurrence object of an appointment series given the series, an instance of that series, or a meeting request.</span></span>

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

<span data-ttu-id="5065e-167">L’exemple suivant montre la valeur de la `item.recurrence` propriété pour une série de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="5065e-167">The following example shows the value of the `item.recurrence` property for an appointment series.</span></span>

> [!NOTE]
> <span data-ttu-id="5065e-168">Dans cet exemple, `seriesTimeObject` est un espace réservé pour JSON représentant la `recurrence.seriesTime` propriété.</span><span class="sxs-lookup"><span data-stu-id="5065e-168">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="5065e-169">Vous devez utiliser les [`SeriesTime`][SeriesTime link] méthodes pour obtenir les propriétés de date et d’heure de périodicité.</span><span class="sxs-lookup"><span data-stu-id="5065e-169">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-the-recurrence-details"></a><span data-ttu-id="5065e-170">Obtenir les détails de récurrence</span><span class="sxs-lookup"><span data-stu-id="5065e-170">Get the recurrence details</span></span>

<span data-ttu-id="5065e-171">Une fois que vous avez récupéré l’objet récurrence (soit à partir du `getAsync` rappel ou à partir de `item.recurrence`), vous pouvez obtenir les propriétés spécifiques de la récurrence.</span><span class="sxs-lookup"><span data-stu-id="5065e-171">After you've retrieved the recurrence object (either from the `getAsync` callback or from `item.recurrence`), you can get specific properties of the recurrence.</span></span> <span data-ttu-id="5065e-172">Par exemple, vous pouvez accéder aux dates de début et de fin et heures de la série via [méthodes][SeriesTime link] `recurrence.seriesTime` sur la  propriété.</span><span class="sxs-lookup"><span data-stu-id="5065e-172">For example, you can get the start and end dates and times of the series by using [methods][SeriesTime link] on the `recurrence.seriesTime` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="5065e-173">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5065e-173">See also</span></span>

[<span data-ttu-id="5065e-174">Événement RecurrenceChanged</span><span class="sxs-lookup"><span data-stu-id="5065e-174">RecurrenceChanged event</span></span>](/javascript/api/office/office.eventtype)

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
