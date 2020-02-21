---
title: Gestion des valeurs de date dans les compléments Outlook
description: L’interface API JavaScript pour Office utilise l’objet JavaScript Date pour stocker et récupérer la plupart des dates et des heures.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 5718839ebda433df6fb14886da34d734f81eb5f2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166073"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a><span data-ttu-id="3d8f3-103">Conseils pour la gestion des valeurs de date dans les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="3d8f3-103">Tips for handling date values in Outlook add-ins</span></span>

<span data-ttu-id="3d8f3-104">L’interface API JavaScript pour Office utilise l’objet JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) pour stocker et récupérer la plupart des dates et des heures.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-104">The JavaScript API for Office uses the JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) object for most of the storage and retrieval of dates and times.</span></span> 

<span data-ttu-id="3d8f3-105">Cet objet **Date** fournit des méthodes telles que [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) et [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), qui renvoient la date ou l’heure UTC demandée.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-105">That **Date** object provides methods such as [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp), and [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), which return the requested date or time value according to Universal Coordinated Time (UTC) time.</span></span>

<span data-ttu-id="3d8f3-106">L’objet **Date** fournit également d’autres méthodes telles que [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) et [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), qui renvoient la date ou l’heure locale demandée.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-106">The **Date** object also provides other methods such as [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp), and [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), which return the requested date or time according to "local time".</span></span>

<span data-ttu-id="3d8f3-p101">Le concept d’« heure locale » est principalement déterminé par le navigateur et le système d’exploitation de l’ordinateur client. Par exemple, dans la plupart des navigateurs s’exécutant sur un ordinateur client Windows, un appel JavaScript à **getDate** renvoie une date en fonction du fuseau horaire défini dans Windows sur l’ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-p101">The concept of "local time" is largely determined by the browser and operating system on the client computer. For instance, on most browsers running on a Windows-based client computer, a JavaScript call to **getDate**, returns a date based on the time zone set in Windows on the client computer.</span></span>

<span data-ttu-id="3d8f3-109">L’exemple suivant crée un objet **Date** `myLocalDate` au format de l’heure locale, et appelle **toUTCString** pour convertir cette date en chaîne de date au format UTC.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-109">The following example creates a **Date** object `myLocalDate` in local time, and calls **toUTCString** to convert that date to a date string in UTC.</span></span>

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

<span data-ttu-id="3d8f3-110">Si vous pouvez utiliser le code JavaScript **Date** pour obtenir une valeur de date ou l’heure en fonction de UTC ou le fuseau horaire d’ordinateur client, l’objet **Date** est limité à un égard : il ne fournit pas de méthodes pour renvoyer une date ou valeur de temps pour n’importe quel autre fuseau horaire.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-110">While you can use the JavaScript **Date** object to get a date or time value based on UTC or the client computer time zone, the **Date** object is limited in one respect - it does not provide methods to return a date or time value for any other specific time zone.</span></span> <span data-ttu-id="3d8f3-111">Par exemple, si votre ordinateur client est défini pour être en horaire Standard est (EST), il n’existe aucune méthode**Date** qui vous permet d’obtenir la valeur d’heure autre que dans h EST ou UTC, comme par exemple l’heure du Pacifique (PST).</span><span class="sxs-lookup"><span data-stu-id="3d8f3-111">For example, if your client computer is set to be on Eastern Standard Time (EST), there is no **Date** method that allows you to get the hour value other than in EST or UTC, such as Pacific Standard Time (PST).</span></span>


## <a name="date-related-features-for-outlook-add-ins"></a><span data-ttu-id="3d8f3-112">Fonctionnalités liées à la date pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="3d8f3-112">Date-related features for Outlook add-ins</span></span>

<span data-ttu-id="3d8f3-113">La limitation JavaScript mentionnée ci-dessus a une implication, lorsque vous utilisez l’API JavaScript pour Office pour gérer les valeurs de date ou d’heure dans les compléments Outlook qui s’exécutent dans un client riche Outlook, ainsi que dans Outlook sur le Web ou les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-113">The aforementioned JavaScript limitation has an implication for you, when you use the JavaScript API for Office to handle date or time values in Outlook add-ins that run in an Outlook rich client, and in Outlook on the web or mobile devices.</span></span>


### <a name="time-zones-for-outlook-clients"></a><span data-ttu-id="3d8f3-114">Fuseaux horaires pour les clients Outlook</span><span class="sxs-lookup"><span data-stu-id="3d8f3-114">Time zones for Outlook clients</span></span>

<span data-ttu-id="3d8f3-115">Pour clarifier les choses, définissons les fuseaux horaires en question.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-115">For clarity, let's define the time zones in question.</span></span>

|<span data-ttu-id="3d8f3-116">**Fuseau horaire**</span><span class="sxs-lookup"><span data-stu-id="3d8f3-116">**Time zone**</span></span>|<span data-ttu-id="3d8f3-117">**Description**</span><span class="sxs-lookup"><span data-stu-id="3d8f3-117">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="3d8f3-118">Fuseau horaire de l’ordinateur client</span><span class="sxs-lookup"><span data-stu-id="3d8f3-118">Client computer time zone</span></span>|<span data-ttu-id="3d8f3-119">Ce champ est défini sur le système d’exploitation de l’ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-119">This is set on the operating system of the client computer.</span></span> <span data-ttu-id="3d8f3-120">La plupart des navigateurs utilisent le fuseau horaire de l’ordinateur client pour afficher les valeurs de date ou d’heure de l’objet JavaScript **Date**.  </span><span class="sxs-lookup"><span data-stu-id="3d8f3-120">Most browsers use the client computer time zone to display date or time values of the JavaScript **Date** object.</span></span><br/><br/><span data-ttu-id="3d8f3-121">Le client Outlook utilise ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-121">An Outlook rich client uses this time zone to display date or time values in the user interface.</span></span> <br/><br/><span data-ttu-id="3d8f3-122">Par exemple, sur un ordinateur client exécutant Windows, Outlook utilise le fuseau horaire défini sur Windows comme fuseau horaire local.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-122">For example, on a client computer running Windows, Outlook uses the time zone set on Windows as the local time zone.</span></span> <span data-ttu-id="3d8f3-123">Sur Mac, si l’utilisateur modifie le fuseau horaire sur l’ordinateur client, Outlook sur Mac invite également l’utilisateur à mettre à jour le fuseau horaire dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-123">On the Mac, if the user changes the time zone on the client computer, Outlook on Mac would prompt the user to update the time zone in Outlook as well.</span></span>|
|<span data-ttu-id="3d8f3-124">Fuseau horaire EAC (Exchange Admin Center)</span><span class="sxs-lookup"><span data-stu-id="3d8f3-124">Exchange Admin Center (EAC) time zone</span></span>|<span data-ttu-id="3d8f3-125">L’utilisateur définit cette valeur de fuseau horaire (et la langue préférée) lorsqu’il se connecte à Outlook sur le Web ou les appareils mobiles la première fois.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-125">The user sets this time zone value (and the preferred language) when he or she logs on to Outlook on the web or mobile devices the first time.</span></span> <br/><br/><span data-ttu-id="3d8f3-126">Outlook sur le Web et les appareils mobiles utilisez ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-126">Outlook on the web and mobile devices use this time zone to display date or time values in the user interface.</span></span>|

<span data-ttu-id="3d8f3-127">Étant donné qu’un client riche Outlook utilise le fuseau horaire de l’ordinateur client et que l’interface utilisateur d’Outlook sur le Web et les appareils mobiles utilise le fuseau horaire du centre d’administration Exchange, l’heure locale pour le même complément installé pour la même boîte aux lettres peut être différente lors de l’exécution dans une Clie riche Outlook NT et dans Outlook sur le Web ou les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-127">Because an Outlook rich client uses the client computer time zone, and the user interface of Outlook on the web and mobile devices uses the EAC time zone, the local time for the same add-in installed for the same mailbox can be different when running in an Outlook rich client and in Outlook on the web or mobile devices.</span></span> <span data-ttu-id="3d8f3-128">En tant que développeur de complément Outlook, vous devez entrer et sortir de façon appropriée les valeurs de date afin qu’elles soient toujours en accord avec le fuseau horaire attendu par l’utilisateur sur le client correspondant.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-128">As an Outlook add-in developer, you should appropriately input and output date values so that those values are always consistent with the time zone that the user expects on the corresponding client.</span></span>


### <a name="date-related-api"></a><span data-ttu-id="3d8f3-129">API liée à la date</span><span class="sxs-lookup"><span data-stu-id="3d8f3-129">Date-related API</span></span>

<span data-ttu-id="3d8f3-130">Les propriétés et méthodes suivantes de l’API JavaScript pour Office prennent en charge des fonctionnalités associées à la date.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-130">The following are the properties and methods in the JavaScript API for Office that support date-related features.</span></span>

<span data-ttu-id="3d8f3-131">**Membre de l'API**</span><span class="sxs-lookup"><span data-stu-id="3d8f3-131">**API member**</span></span>|<span data-ttu-id="3d8f3-132">**Représentation du fuseau horaire**</span><span class="sxs-lookup"><span data-stu-id="3d8f3-132">**Time zone representation**</span></span>|<span data-ttu-id="3d8f3-133">**Exemple dans un client riche Outlook**</span><span class="sxs-lookup"><span data-stu-id="3d8f3-133">**Example in an Outlook rich client**</span></span>|<span data-ttu-id="3d8f3-134">**Exemple dans Outlook sur le Web ou les appareils mobiles**</span><span class="sxs-lookup"><span data-stu-id="3d8f3-134">**Example in Outlook on the web or mobile devices**</span></span>
--------------|----------------------------|-------------------------------------|-------------------
[<span data-ttu-id="3d8f3-135">Office.context.mailbox.userProfile.timeZone</span><span class="sxs-lookup"><span data-stu-id="3d8f3-135">Office.context.mailbox.userProfile.timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone)|<span data-ttu-id="3d8f3-136">Dans un client riche Outlook, cette propriété renvoie le fuseau horaire de l’ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-136">In an Outlook rich client, this property returns the client computer time zone.</span></span> <span data-ttu-id="3d8f3-137">Dans Outlook sur le Web et les appareils mobiles, cette propriété renvoie le fuseau horaire du centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-137">In Outlook on the web and mobile devices, this property returns the EAC time zone.</span></span> |<span data-ttu-id="3d8f3-138">EST</span><span class="sxs-lookup"><span data-stu-id="3d8f3-138">EST</span></span>|<span data-ttu-id="3d8f3-139">PST</span><span class="sxs-lookup"><span data-stu-id="3d8f3-139">PST</span></span>
<span data-ttu-id="3d8f3-140">[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)</span><span class="sxs-lookup"><span data-stu-id="3d8f3-140">[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)</span></span>|<span data-ttu-id="3d8f3-141">Chacune de ces propriétés renvoie un objet JavaScript **Date**.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-141">Each of these properties returns a JavaScript **Date** object.</span></span> <span data-ttu-id="3d8f3-142">Cette valeur de **Date** est au format UTC, comme indiqué dans l’exemple suivant `myUTCDate` : a la même valeur dans un client riche Outlook, Outlook sur le Web et les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-142">This **Date** value is UTC-correct, as shown in the following example - `myUTCDate` has the same value in an Outlook rich client, Outlook on the web and mobile devices.</span></span><br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/><span data-ttu-id="3d8f3-143">Toutefois, l' `myDate.getDate` appel renvoie une valeur de date dans le fuseau horaire de l’ordinateur client, qui est cohérente avec le fuseau horaire utilisé pour afficher les valeurs de date et d’heure dans l’interface client riche Outlook, mais peut être différent du fuseau horaire du centre d’administration Exchange sur le Web et les appareils mobiles utilisés dans son interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-143">However, calling  `myDate.getDate` returns a date value in the client computer time zone, which is consistent with the time zone used to display date times values in the Outlook rich client interface, but may be different from the EAC time zone that Outlook on the web and mobile devices use in its user interface.</span></span>|<span data-ttu-id="3d8f3-144">Si l’élément est créé à 9 h 00 UTC :</span><span class="sxs-lookup"><span data-stu-id="3d8f3-144">If the item is created at 9am UTC:</span></span><br/><br/>`Office.mailbox.item.`<br/><span data-ttu-id="3d8f3-145">`dateTimeCreated.getHours` renvoie 4 h 00 EST.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-145">`dateTimeCreated.getHours` returns 4am EST.</span></span><br/><br/><span data-ttu-id="3d8f3-146">Si l’élément est modifié à 11 h 00 UTC :</span><span class="sxs-lookup"><span data-stu-id="3d8f3-146">If the item is modified at 11am UTC:</span></span><br/><br/>`Office.mailbox.item.`<br/><span data-ttu-id="3d8f3-147">`dateTimeModified.getHours` renvoie 6h00 EST.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-147">`dateTimeModified.getHours` returns 6am EST.</span></span>|<span data-ttu-id="3d8f3-148">Si l’élément est créé à 9 h 00 UTC :</span><span class="sxs-lookup"><span data-stu-id="3d8f3-148">If the item creation time is 9am UTC:</span></span><br/><br/>`Office.mailbox.item.`</br><span data-ttu-id="3d8f3-149">`dateTimeCreated.getHours` renvoie 4 h 00 EST.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-149">`dateTimeCreated.getHours` returns 4am EST.</span></span><br/><br/><span data-ttu-id="3d8f3-150">Si l’élément est modifié à 11 h 00 UTC :</span><span class="sxs-lookup"><span data-stu-id="3d8f3-150">If the item is modified at 11am UTC:</span></span><br/><br/>`Office.mailbox.item.`</br><span data-ttu-id="3d8f3-151">`dateTimeModified.getHours` renvoie 6 h 00 EST.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-151">`dateTimeModified.getHours` returns 6am EST.</span></span><br/><br/><span data-ttu-id="3d8f3-152">Notez que si vous souhaitez afficher l’heure de création ou de modification dans l’interface utilisateur, vous pouvez d’abord convertir l’heure au format PST pour rester cohérent avec le reste de l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-152">Notice that if you want to display the creation or modification time in the user interface, you would want to first convert the time to PST to be consistent with the rest of the user interface.</span></span>
[<span data-ttu-id="3d8f3-153">Office.context.mailbox.displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3d8f3-153">Office.context.mailbox.displayNewAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|<span data-ttu-id="3d8f3-154">Chacun des paramètres  _Start_ et _End_ nécessite un objet JavaScript **Date**.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-154">Each of the  _Start_ and _End_ parameters requires a JavaScript **Date** object.</span></span> <span data-ttu-id="3d8f3-155">Les arguments doivent être au format UTC, quel que soit le fuseau horaire utilisé dans l’interface utilisateur d’un client riche Outlook, ou Outlook sur le Web ou les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-155">The arguments should be UTC-correct regardless of the time zone used in the user interface of an Outlook rich client, or Outlook on the web or mobile devices.</span></span>|<span data-ttu-id="3d8f3-156">Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :</span><span class="sxs-lookup"><span data-stu-id="3d8f3-156">If the start and end times for the appointment form are 9am UTC and 11am UTC, then you should assure that the `start` and `end` arguments are UTC-correct, which means:</span></span><br/><br/><ul><li><span data-ttu-id="3d8f3-157">`start.getUTCHours` renvoie 9 h 00 UTC</span><span class="sxs-lookup"><span data-stu-id="3d8f3-157">`start.getUTCHours` returns 9am UTC</span></span></li><li><span data-ttu-id="3d8f3-158">`end.getUTCHours` renvoie 11 h 00 UTC</span><span class="sxs-lookup"><span data-stu-id="3d8f3-158">`end.getUTCHours` returns 11am UTC</span></span></li></ul>|<span data-ttu-id="3d8f3-159">Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :</span><span class="sxs-lookup"><span data-stu-id="3d8f3-159">If the start and end times for the appointment form are 9am UTC and 11am UTC, then you should assure that the `start` and `end` arguments are UTC-correct, which means:</span></span><br/><br/><ul><li><span data-ttu-id="3d8f3-160">`start.getUTCHours` renvoie 9 h 00 UTC</span><span class="sxs-lookup"><span data-stu-id="3d8f3-160">`start.getUTCHours` returns 9am UTC</span></span></li><li><span data-ttu-id="3d8f3-161">`end.getUTCHours` renvoie 11 h 00 UTC</span><span class="sxs-lookup"><span data-stu-id="3d8f3-161">`end.getUTCHours` returns 11am UTC</span></span></li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a><span data-ttu-id="3d8f3-162">Méthodes d’assistance pour les scénarios liés à la date</span><span class="sxs-lookup"><span data-stu-id="3d8f3-162">Helper methods for date-related scenarios</span></span>


<span data-ttu-id="3d8f3-163">Comme décrit dans les sections précédentes, étant donné que la « durée locale » pour un utilisateur dans Outlook sur le Web ou les appareils mobiles peut être différente sur un client riche Outlook, mais que l’objet JavaScript **Date** prend en charge la conversion uniquement du fuseau horaire de l’ordinateur client ou de l’heure UTC, l’interface API JavaScript pour Office fournit deux [](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)méthodes d’assistance : [Office](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)</span><span class="sxs-lookup"><span data-stu-id="3d8f3-163">As described in the preceding sections, because the "local time" for a user in Outlook on the web or mobile devices can be different on an Outlook rich client, but the JavaScript **Date** object supports converting to only the client computer time zone or UTC, the JavaScript API for Office provides two helper methods: [Office.context.mailbox.convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) and [Office.context.mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span></span>

<span data-ttu-id="3d8f3-164">Ces méthodes d’assistance ont besoin de gérer la date ou l’heure différemment pour les deux scénarios de date suivants, dans un client riche Outlook, Outlook sur le Web et les appareils mobiles, renforçant ainsi « l’écriture unique » pour les différents clients de votre complément.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-164">These helper methods take care of any need to handle date or time differently for the following two date-related scenarios, in an Outlook rich client, Outlook on the web and mobile devices, thus reinforcing "write-once" for different clients of your add-in.</span></span>


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a><span data-ttu-id="3d8f3-165">Scénario A : affichage de l’heure de création ou de modification d’un élément</span><span class="sxs-lookup"><span data-stu-id="3d8f3-165">Scenario A: Displaying item creation or modified time</span></span>

<span data-ttu-id="3d8f3-166">Si vous affichez l’heure de création (**Item.dateTimeCreated**) ou de modification (**Item.dateTimeModified**) d’un élément dans l’interface utilisateur, utilisez d’abord **convertToLocalClientTime** pour convertir l’objet **Date** fourni par ces propriétés pour obtenir une représentation de dictionnaire dans l’heure locale appropriée. </span><span class="sxs-lookup"><span data-stu-id="3d8f3-166">If you are displaying the item creation time (**Item.dateTimeCreated**) or modification time (**Item.dateTimeModified**) in the user interface, first use **convertToLocalClientTime** to convert the **Date** object provided by these properties to obtain a dictionary representation in the appropriate local time.</span></span> <span data-ttu-id="3d8f3-167">Affichez ensuite les parties de la date de dictionnaire. </span><span class="sxs-lookup"><span data-stu-id="3d8f3-167">Then display the parts of the dictionary date.</span></span> <span data-ttu-id="3d8f3-168">L’exemple suivant illustre ce scénario :</span><span class="sxs-lookup"><span data-stu-id="3d8f3-168">The following is an example of this scenario:</span></span>


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

<span data-ttu-id="3d8f3-169">Notez que **convertToLocalClientTime** prend en charge la différence entre un client riche Outlook et Outlook sur le Web ou les appareils mobiles :</span><span class="sxs-lookup"><span data-stu-id="3d8f3-169">Note that **convertToLocalClientTime** takes care of the difference between an Outlook rich client, and Outlook on the web or mobile devices:</span></span>


- <span data-ttu-id="3d8f3-170">Si **convertToLocalClientTime** détecte que l’hôte actuel est un client riche, la méthode convertit la représentation **Date** en une représentation de dictionnaire dans le fuseau horaire de l’ordinateur client, en accord avec le reste de l’interface utilisateur du client riche.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-170">If **convertToLocalClientTime** detects the current host is a rich client, the method converts the **Date** representation to a dictionary representation in the same client computer time zone, consistent with the rest of the rich client user interface.</span></span>
    
- <span data-ttu-id="3d8f3-171">Si **convertToLocalClientTime** détecte que l’hôte actuel est Outlook sur le Web ou les appareils mobiles, la méthode convertit la représentation de **Date** UTC correcte en un format de dictionnaire dans le fuseau horaire d’un centre d’administration Exchange, cohérent avec le reste de l’interface utilisateur d’Outlook sur le Web ou sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-171">If **convertToLocalClientTime** detects the current host is Outlook on the web or mobile devices, the method converts the UTC-correct **Date** representation to a dictionary format in the EAC time zone, consistent with the rest of the Outlook on the web or mobile devices user interface.</span></span>
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a><span data-ttu-id="3d8f3-172">Scénario B : affichage des dates de début et de fin dans un formulaire de nouveau rendez-vous</span><span class="sxs-lookup"><span data-stu-id="3d8f3-172">Scenario B: Displaying start and end dates in a new appointment form</span></span>

<span data-ttu-id="3d8f3-173">Si vous obtenez différentes parties d’une valeur d’entrée date-heure à l’heure locale et que vous souhaitez fournir la valeur d’entrée du dictionnaire sous la forme d’une heure de début ou de fin dans un formulaire de rendez-vous, utilisez d’abord la méthode d’assistance **convertToUtcClientTime** pour convertir la valeur de dictionnaire en objet **Date** au format UTC.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-173">If you are obtaining as input different parts of a date-time value represented in the local time, and would like to provide this dictionary input value as a start or end time in an appointment form, first use the **convertToUtcClientTime** helper method to convert the dictionary value to a UTC-correct **Date** object.</span></span>

<span data-ttu-id="3d8f3-p110">Dans l’exemple suivant, supposons que  `myLocalDictionaryStartDate` et `myLocalDictionaryEndDate` sont des valeurs de date et d’heure au format de dictionnaire que vous avez obtenues auprès de l’utilisateur. Ces valeurs sont basées sur l’heure locale, qui dépend elle-même de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-p110">In the following example, assume  `myLocalDictionaryStartDate` and `myLocalDictionaryEndDate` are date-time values in dictionary format that you have obtained from the user. These values are based on the local time, dependent on the host application.</span></span>

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

<span data-ttu-id="3d8f3-176">Les valeurs qui en résultent, `myUTCCorrectStartDate` et `myUTCCorrectEndDate`, sont au format UTC.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-176">The resultant values,  `myUTCCorrectStartDate` and `myUTCCorrectEndDate`, are UTC-correct.</span></span> <span data-ttu-id="3d8f3-177">Transférez ensuite ces objets **Date** en tant qu’arguments pour les paramètres_Start_ et _End_ de la méthode **Mailbox.displayNewAppointmentForm** pour afficher le nouveau formulaire de rendez-vous. </span><span class="sxs-lookup"><span data-stu-id="3d8f3-177">Then pass these **Date** objects as arguments for the _Start_ and _End_ parameters of the **Mailbox.displayNewAppointmentForm** method to display the new appointment form.</span></span>

<span data-ttu-id="3d8f3-178">Notez que **convertToUtcClientTime** prend en charge la différence entre un client riche Outlook et Outlook sur le Web ou les appareils mobiles :</span><span class="sxs-lookup"><span data-stu-id="3d8f3-178">Note that **convertToUtcClientTime** takes care of the difference between an Outlook rich client, and Outlook on the web or mobile devices:</span></span>


- <span data-ttu-id="3d8f3-179">Si **convertToUtcClientTime** détecte que l’hôte actuel est un client riche Outlook, la méthode convertit simplement la représentation de dictionnaire en objet **Date**. </span><span class="sxs-lookup"><span data-stu-id="3d8f3-179">If **convertToUtcClientTime** detects the current host is an Outlook rich client, the method simply converts the dictionary representation to a **Date** object.</span></span> <span data-ttu-id="3d8f3-180">Cet objet **Date** est conforme au format UTC, comme attendu par **displayNewAppointmentForm**.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-180">This **Date** object is UTC-correct, as expected by **displayNewAppointmentForm**.</span></span>
    
- <span data-ttu-id="3d8f3-181">Si **convertToUtcClientTime** détecte que l’hôte actuel est Outlook sur le Web ou les appareils mobiles, la méthode convertit le format de dictionnaire des valeurs de date et d’heure exprimées dans le fuseau horaire du centre d’administration Exchange en un objet **Date** .</span><span class="sxs-lookup"><span data-stu-id="3d8f3-181">If **convertToUtcClientTime** detects the current host is Outlook on the web or mobile devices, the method converts the dictionary format of the date and time values expressed in the EAC time zone to a **Date** object.</span></span> <span data-ttu-id="3d8f3-182">Cet objet **Date** est conforme au format UTC, comme attendu par **displayNewAppointmentForm**.</span><span class="sxs-lookup"><span data-stu-id="3d8f3-182">This **Date** object is UTC-correct, as expected by **displayNewAppointmentForm**.</span></span>
    

## <a name="see-also"></a><span data-ttu-id="3d8f3-183">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3d8f3-183">See also</span></span>

- [<span data-ttu-id="3d8f3-184">Déployer et installer des compléments Outlook à des fins de test</span><span class="sxs-lookup"><span data-stu-id="3d8f3-184">Deploy and install Outlook add-ins for testing</span></span>](testing-and-tips.md)
    

