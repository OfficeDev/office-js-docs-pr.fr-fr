---
title: Gestion des valeurs de date dans les compléments Outlook
description: L’API JavaScript Office utilise l’objet Date JavaScript pour la plupart du stockage et la récupération des dates et heures.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 49de8db712400e006dc919e9ad62ae6cbaaa11cf
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713076"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Conseils pour la gestion des valeurs de date dans les compléments Outlook

L’API JavaScript Office utilise l’objet [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) JavaScript pour la plupart du stockage et la récupération des dates et heures.

Cet `Date` objet fournit des méthodes telles que [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) et [toUTCString, qui retournent](https://www.w3schools.com/jsref/jsref_toutcstring.asp) la valeur de date ou d’heure demandée en fonction de l’heure UTC (Universal Coordinated Time).

L’objet `Date` fournit également d’autres méthodes telles que [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) et [toString, qui retournent](https://www.w3schools.com/jsref/jsref_tostring_date.asp) la date ou l’heure demandée en fonction de l'« heure locale ».

Le concept d’« heure locale » est principalement déterminé par le navigateur et le système d’exploitation de l’ordinateur client. Par exemple, sur la plupart des navigateurs s’exécutant sur un ordinateur client Windows, un appel JavaScript retourne `getDate`une date basée sur le fuseau horaire défini dans Windows sur l’ordinateur client.

L’exemple suivant crée un `Date` objet `myLocalDate` en heure locale et appelle `toUTCString` pour convertir cette date en chaîne de date en UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
const myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Bien que vous puissiez utiliser l’objet JavaScript `Date` pour obtenir une valeur de date ou d’heure en fonction de l’heure UTC ou du fuseau horaire de l’ordinateur client, l’objet **Date** est limité d’un point de vue : il ne fournit pas de méthodes pour retourner une valeur de date ou d’heure pour un autre fuseau horaire spécifique. Par exemple, si votre ordinateur client est défini sur eastern standard time (EST), il n’existe aucune `Date` méthode qui vous permet d’obtenir la valeur d’heure autre que dans EST ou UTC, comme pacific standard time (PST).

## <a name="date-related-features-for-outlook-add-ins"></a>Fonctionnalités liées à la date pour les compléments Outlook

La limitation JavaScript mentionnée ci-dessus a une incidence pour vous, lorsque vous utilisez l’API JavaScript Office pour gérer les valeurs de date ou d’heure dans les compléments Outlook qui s’exécutent dans un client riche Outlook et dans Outlook sur le web ou appareils mobiles.

### <a name="time-zones-for-outlook-clients"></a>Fuseaux horaires pour les clients Outlook

Pour clarifier les choses, définissons les fuseaux horaires en question.

|**Fuseau horaire**|**Description**|
|:-----|:-----|
|Fuseau horaire de l’ordinateur client|Cette valeur est définie sur le système d’exploitation de l’ordinateur client. La plupart des navigateurs utilisent le fuseau horaire de l’ordinateur client pour afficher les valeurs de date ou d’heure de l’objet JavaScript `Date` .<br/><br/>Le client riche Outlook utilise ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur. <br/><br/>Par exemple, sur un ordinateur client exécutant Windows, Outlook utilise le fuseau horaire défini sur Windows comme fuseau horaire local. Sur mac, si l’utilisateur modifie le fuseau horaire sur l’ordinateur client, Outlook sur Mac invite également l’utilisateur à mettre à jour le fuseau horaire dans Outlook.|
|Fuseau horaire EAC (Exchange Admin Center)|L’utilisateur définit cette valeur de fuseau horaire (et la langue par défaut) lorsqu’il se connecte à Outlook sur le web ou à des appareils mobiles la première fois. <br/><br/>Outlook sur le web et les appareils mobiles utilisent ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur.|

Étant donné qu’un client riche Outlook utilise le fuseau horaire de l’ordinateur client et que l’interface utilisateur de Outlook sur le web et des appareils mobiles utilise le fuseau horaire EAC, l’heure locale du même complément installé pour la même boîte aux lettres peut être différente lors de l’exécution dans un client riche Outlook et dans Outlook sur le web ou des appareils mobiles. En tant que développeur de complément Outlook, vous devez entrer et sortir de façon appropriée les valeurs de date afin qu’elles soient toujours en accord avec le fuseau horaire attendu par l’utilisateur sur le client correspondant.

### <a name="date-related-api"></a>API liée à la date

Voici les propriétés et méthodes de l’API JavaScript Office qui prennent en charge les fonctionnalités liées aux dates.

|Membre de l'API|Représentation du fuseau horaire|Exemple dans un client riche Outlook|Exemple dans Outlook sur le web ou appareils mobiles|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#outlook-office-userprofile-timezone-member)|Dans un client riche Outlook, cette propriété renvoie le fuseau horaire de l’ordinateur client. Dans Outlook sur le web et les appareils mobiles, cette propriété renvoie le fuseau horaire EAC. |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) et [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Chacune de ces propriétés retourne un objet JavaScript `Date` . Cette `Date` valeur est correcte au format UTC, comme indiqué dans l’exemple suivant : `myUTCDate` elle a la même valeur dans un client riche Outlook, Outlook sur le web et des appareils mobiles.<br/><br/>`const myDate = Office.mailbox.item.dateTimeCreated;`<br/>`const myUTCDate = myDate.getUTCDate;`<br/><br/>Toutefois, l’appel `myDate.getDate` retourne une valeur de date dans le fuseau horaire de l’ordinateur client, qui est cohérente avec le fuseau horaire utilisé pour afficher les valeurs d’heures de date dans l’interface cliente riche Outlook, mais peut être différent du fuseau horaire EAC que Outlook sur le web et les appareils mobiles utilisent dans son interface utilisateur.|Si l’élément est créé à 9 h UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` renvoie 6 h 00 EST.|Si l’heure de création de l’élément est 9 h UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` renvoie 6 h 00 EST.<br/><br/>Notez que si vous souhaitez afficher l’heure de création ou de modification dans l’interface utilisateur, vous pouvez d’abord convertir l’heure au format PST pour rester cohérent avec le reste de l’interface utilisateur.|
|[Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)|Chacun des paramètres _Start_ et _End_ nécessite un objet JavaScript `Date` . Les arguments doivent être corrects au format UTC, quel que soit le fuseau horaire utilisé dans l’interface utilisateur d’un client riche Outlook, ou Outlook sur le web ou des appareils mobiles.|Si les heures de début et de fin du formulaire de rendez-vous sont 9 h UTC et 11 h UTC, vous devez vous assurer que les arguments et `end` les `start` arguments sont UTC corrects, ce qui signifie :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>|Si les heures de début et de fin du formulaire de rendez-vous sont 9 h UTC et 11 h UTC, vous devez vous assurer que les arguments et `end` les `start` arguments sont UTC corrects, ce qui signifie :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>Méthodes d’assistance pour les scénarios liés à la date

Comme décrit dans les sections précédentes, étant donné que l'« heure locale » d’un utilisateur dans Outlook sur le web ou des appareils mobiles peut être différente sur un client riche Outlook, mais que l’objet **Date** JavaScript prend en charge la conversion uniquement en fuseau horaire de l’ordinateur client ou en UTC, l’API JavaScript Office fournit deux méthodes d’assistance : [Office.context.mailbox.convertToLocalClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) et [Office.context.mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).

Ces méthodes d’assistance prennent en charge tout besoin de gérer la date ou l’heure différemment pour les deux scénarios liés aux dates suivants, dans un client riche Outlook, Outlook sur le web et des appareils mobiles, renforçant ainsi l’écriture en une seule fois pour les différents clients de votre complément.

### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Scénario A : affichage de l’heure de création ou de modification d’un élément

Si vous affichez l’heure de création (`Item.dateTimeCreated`) ou de modification de l’élément (`Item.dateTimeModified`dans l’interface utilisateur), commencez `convertToLocalClientTime` par convertir l’objet `Date` fourni par ces propriétés pour obtenir une représentation de dictionnaire à l’heure locale appropriée. Affichez ensuite les parties de la date de dictionnaire. Voici un exemple de ce scénario.

```js
// This date is UTC-correct.
const myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
const myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Notez que `convertToLocalClientTime` vous prenez en charge la différence entre un client riche Outlook et Outlook sur le web ou des appareils mobiles :

- Si `convertToLocalClientTime` elle détecte que l’application actuelle est un client riche, la méthode convertit la `Date` représentation en représentation de dictionnaire dans le même fuseau horaire de l’ordinateur client, conformément au reste de l’interface utilisateur cliente enrichie.

- Si `convertToLocalClientTime` elle détecte que l’application actuelle est Outlook sur le web ou des appareils mobiles, la méthode convertit la représentation UTC correcte `Date` en un format de dictionnaire dans le fuseau horaire EAC, conformément au reste de l’interface utilisateur de l’Outlook sur le web ou des appareils mobiles.

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Scénario B : affichage des dates de début et de fin dans un formulaire de nouveau rendez-vous

Si vous obtenez en entrée différentes parties d’une valeur date-heure représentée à l’heure locale et que vous souhaitez fournir cette valeur d’entrée de dictionnaire sous forme d’heure de début ou de fin dans un formulaire de rendez-vous, commencez par utiliser la `convertToUtcClientTime` méthode d’assistance pour convertir la valeur du dictionnaire en objet correct `Date` UTC.

Dans l’exemple suivant, supposons que  `myLocalDictionaryStartDate` et `myLocalDictionaryEndDate` sont des valeurs de date et d’heure au format de dictionnaire que vous avez obtenues auprès de l’utilisateur. Ces valeurs sont basées sur l’heure locale, en fonction de la plateforme cliente.

```js
const myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
const myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Les valeurs qui en résultent, `myUTCCorrectStartDate` et `myUTCCorrectEndDate`, sont au format UTC. Transmettez ensuite ces `Date` objets en tant qu’arguments pour les paramètres _Start_ et _End_ de la `Mailbox.displayNewAppointmentForm` méthode afin d’afficher le nouveau formulaire de rendez-vous.

Notez que `convertToUtcClientTime` vous prenez en charge la différence entre un client riche Outlook et Outlook sur le web ou des appareils mobiles :

- Si `convertToUtcClientTime` elle détecte que l’application actuelle est un client riche Outlook, la méthode convertit simplement la représentation de dictionnaire en `Date` objet. Cet `Date` objet est correct UTC, comme prévu par `displayNewAppointmentForm`.

- Si `convertToUtcClientTime` elle détecte que l’application actuelle est Outlook sur le web ou des appareils mobiles, la méthode convertit le format de dictionnaire des valeurs de date et d’heure exprimées dans le fuseau horaire EAC en objet`Date`. Cet `Date` objet est correct UTC, comme prévu par `displayNewAppointmentForm`.

## <a name="see-also"></a>Voir aussi

- [Déployer et installer des compléments Outlook à des fins de test](testing-and-tips.md)
