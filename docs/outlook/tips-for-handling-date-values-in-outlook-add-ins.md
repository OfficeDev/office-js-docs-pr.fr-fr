---
title: Gestion des valeurs de date dans les compléments Outlook
description: L’API JavaScript pour Office utilise l’objet Date JavaScript pour la plupart du stockage et de la récupération des dates et heures.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4f0e93d147760f91c55fd5666f02b2c6cc5d5470
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/24/2021
ms.locfileid: "51177991"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Conseils pour la gestion des valeurs de date dans les compléments Outlook

L’API JavaScript pour Office utilise l’objet [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) JavaScript pour la plupart du stockage et de la récupération des dates et heures. 

Cet objet fournit des méthodes telles que `Date` [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)et [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), qui retournent la valeur de date ou d’heure demandée en fonction de l’heure UTC (Universal Coordinated Time).

L’objet fournit également d’autres méthodes telles que `Date` [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)et [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), qui retournent la date ou l’heure demandée en fonction de « l’heure locale ».

Le concept d’« heure locale » est principalement déterminé par le navigateur et le système d’exploitation de l’ordinateur client. Par exemple, sur la plupart des navigateurs s’exécutant sur un ordinateur client Windows, un appel JavaScript à , renvoie une date basée sur le fuseau horaire définie dans Windows sur `getDate` l’ordinateur client.

L’exemple suivant crée un objet à l’heure locale et appelle pour convertir cette date en `Date` une chaîne de date au `myLocalDate` `toUTCString` UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Bien que vous pouvez utiliser l’objet JavaScript pour obtenir une valeur de date ou d’heure basée sur UTC ou le fuseau horaire de l’ordinateur client, l’objet Date est limité à un point : il ne fournit pas de méthodes pour renvoyer une valeur de date ou d’heure pour tout autre fuseau horaire `Date` spécifique.  Par exemple, si votre ordinateur client est définie pour être à l’heure standard est (EST), il n’existe aucune méthode qui vous permet d’obtenir la valeur d’heure autre qu’est ou UTC, telle que l’heure `Date` PST (Pacific Standard Time).


## <a name="date-related-features-for-outlook-add-ins"></a>Fonctionnalités liées à la date pour les compléments Outlook

La limitation JavaScript susmentionnée a une incidence sur vous, lorsque vous utilisez l’API JavaScript Office pour gérer les valeurs de date ou d’heure dans les applications Outlook qui s’exécutent dans un client riche Outlook et dans Outlook sur le web ou les appareils mobiles.


### <a name="time-zones-for-outlook-clients"></a>Fuseaux horaires pour les clients Outlook

Pour clarifier les choses, définissons les fuseaux horaires en question.

|**Fuseau horaire**|**Description**|
|:-----|:-----|
|Fuseau horaire de l’ordinateur client|Cette valeur est définie sur le système d’exploitation de l’ordinateur client. La plupart des navigateurs utilisent le fuseau horaire de l’ordinateur client pour afficher les valeurs de date ou d’heure de l’objet `Date` JavaScript.<br/><br/>Le client riche Outlook utilise ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur. <br/><br/>Par exemple, sur un ordinateur client exécutant Windows, Outlook utilise le fuseau horaire défini sur Windows comme fuseau horaire local. Sur mac, si l’utilisateur modifie le fuseau horaire sur l’ordinateur client, Outlook sur Mac invite l’utilisateur à mettre à jour le fuseau horaire dans Outlook également.|
|Fuseau horaire EAC (Exchange Admin Center)|L’utilisateur définit cette valeur de fuseau horaire (et la langue préférée) lorsqu’il se connecte à Outlook sur le web ou aux appareils mobiles la première fois. <br/><br/>Outlook sur le web et les appareils mobiles utilisent ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur.|

Étant donné qu’un client riche Outlook utilise le fuseau horaire de l’ordinateur client et que l’interface utilisateur d’Outlook sur le web et des appareils mobiles utilisent le fuseau horaire EAC, l’heure locale pour le même module installé pour la même boîte aux lettres peut être différente lors de l’exécution dans un client riche Outlook et dans Outlook sur le web ou les appareils mobiles. En tant que développeur de complément Outlook, vous devez entrer et sortir de façon appropriée les valeurs de date afin qu’elles soient toujours en accord avec le fuseau horaire attendu par l’utilisateur sur le client correspondant.


### <a name="date-related-api"></a>API liée à la date

Les propriétés et méthodes de l’API JavaScript Office qui prisent en charge les fonctionnalités liées à la date sont les suivantes.

|Membre de l'API|Représentation du fuseau horaire|Exemple dans un client riche Outlook|Exemple dans Outlook sur le web ou les appareils mobiles|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#timezone)|Dans un client riche Outlook, cette propriété renvoie le fuseau horaire de l’ordinateur client. Dans Outlook sur le web et les appareils mobiles, cette propriété renvoie le fuseau horaire EAC. |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Chacune de ces propriétés renvoie un objet `Date` JavaScript. Cette valeur est correcte au niveau UTC, comme illustré dans l’exemple suivant : elle a la même valeur dans un client riche Outlook, Outlook sur le web et `Date` `myUTCDate` les appareils mobiles.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Toutefois, l’appel renvoie une valeur de date dans le fuseau horaire de l’ordinateur client, qui est cohérente avec le fuseau horaire utilisé pour afficher les valeurs des heures de date dans l’interface client riche Outlook, mais peut être différent du fuseau horaire EAC utilisé par Outlook sur le web et les appareils mobiles dans son  `myDate.getDate` interface utilisateur.|Si l’élément est créé à 9 h 00 UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h 00 UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` renvoie 6 h 00 EST.|Si l’élément est créé à 9 h 00 UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h 00 UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` renvoie 6 h 00 EST.<br/><br/>Notez que si vous souhaitez afficher l’heure de création ou de modification dans l’interface utilisateur, vous pouvez d’abord convertir l’heure au format PST pour rester cohérent avec le reste de l’interface utilisateur.|
|[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Chacun des paramètres  _Start_ et _End_ nécessite un objet `Date` JavaScript. Les arguments doivent être corrects en UTC, quel que soit le fuseau horaire utilisé dans l’interface utilisateur d’un client riche Outlook, ou Outlook sur le web ou les appareils mobiles.|Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>|Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>Méthodes d’assistance pour les scénarios liés à la date


Comme décrit dans les sections précédentes, étant donné que l'« heure locale » pour un utilisateur dans Outlook sur le web ou les appareils mobiles peut être différente sur un client riche Outlook, mais que l’objet **Date** JavaScript prend en charge la conversion vers le fuseau horaire ou UTC de l’ordinateur client, l’API JavaScript Office fournit deux méthodes d’aide : [Office.context.mailbox.convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) et [Office.context.mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).

Ces méthodes d’aide s’occupent de n’importe quel besoin de gérer la date ou l’heure différemment pour les deux scénarios suivants liés à la date, dans un client riche Outlook, Outlook sur le web et les appareils mobiles, ce qui renforce l'« écriture une seule fois » pour les différents clients de votre application.


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Scénario A : affichage de l’heure de création ou de modification d’un élément

Si vous affichez l’heure de création de l’élément () ou l’heure de modification ( dans l’interface utilisateur, utilisez d’abord pour convertir l’objet fourni par ces propriétés afin d’obtenir une représentation de dictionnaire à l’heure `Item.dateTimeCreated` `Item.dateTimeModified` locale `convertToLocalClientTime` `Date` appropriée. Affichez ensuite les parties de la date de dictionnaire. L’exemple suivant illustre ce scénario :


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

Notez que la différence entre un client riche Outlook et Outlook sur le web ou les appareils mobiles est `convertToLocalClientTime` prise en charge :


- Si elle détecte que l’application actuelle est un client riche, la méthode convertit la représentation en représentation de dictionnaire dans le même fuseau horaire de l’ordinateur client, en accord avec le reste de l’interface utilisateur `convertToLocalClientTime` `Date` du client riche.
    
- Si l’application actuelle est Outlook sur le web ou les appareils mobiles, la méthode convertit la représentation UTC correcte au format de dictionnaire dans le fuseau horaire EAC, en accord avec le reste de l’interface utilisateur d’Outlook sur le web ou des appareils `convertToLocalClientTime` `Date` mobiles.
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Scénario B : affichage des dates de début et de fin dans un formulaire de nouveau rendez-vous

Si vous obtenez en tant qu’entrée différentes parties d’une valeur date-heure représentée dans l’heure locale et que vous souhaitez fournir cette valeur d’entrée de dictionnaire en tant qu’heure de début ou de fin dans un formulaire de rendez-vous, utilisez d’abord la méthode d’aide pour convertir la valeur du dictionnaire en un objet `convertToUtcClientTime` UTC `Date` correct.

Dans l’exemple suivant, supposons que  `myLocalDictionaryStartDate` et `myLocalDictionaryEndDate` sont des valeurs de date et d’heure au format de dictionnaire que vous avez obtenues auprès de l’utilisateur. Ces valeurs sont basées sur l’heure locale, en fonction de la plateforme cliente.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Les valeurs qui en résultent, `myUTCCorrectStartDate` et `myUTCCorrectEndDate`, sont au format UTC. Passez ensuite ces objets en tant qu’arguments pour les paramètres Début et Fin de la méthode pour `Date` afficher le nouveau formulaire de   `Mailbox.displayNewAppointmentForm` rendez-vous.

Notez que la différence entre un client riche Outlook et Outlook sur le web ou les appareils mobiles est `convertToUtcClientTime` prise en charge :


- Si l’application actuelle est un client riche Outlook, la méthode convertit simplement la représentation `convertToUtcClientTime` de dictionnaire en `Date` objet. Cet `Date` objet est au UTC correct, comme prévu par `displayNewAppointmentForm` .
    
- Si l’application actuelle est Outlook sur le web ou les appareils mobiles, la méthode convertit en objet le format de dictionnaire des valeurs de date et d’heure exprimées dans le fuseau horaire `convertToUtcClientTime` `Date` EAC. Cet `Date` objet est au UTC correct, comme prévu par `displayNewAppointmentForm` .
    
## <a name="see-also"></a>Voir aussi

- [Déployer et installer des compléments Outlook à des fins de test](testing-and-tips.md)
