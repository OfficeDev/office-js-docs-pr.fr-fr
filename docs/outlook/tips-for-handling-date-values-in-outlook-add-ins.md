---
title: Gestion des valeurs de date dans les compléments Outlook
description: L’API JavaScript pour Office utilise l’objet JavaScript date pour la plupart du stockage et de l’extraction des dates et des heures.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 48cbc407e21e377ed64dc873574d938b136bfd22
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292565"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Conseils pour la gestion des valeurs de date dans les compléments Outlook

L’API JavaScript pour Office utilise l’objet JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) pour la plupart du stockage et de l’extraction des dates et des heures. 

Cet `Date` objet fournit des méthodes telles que [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [GetUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)et [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), qui renvoient la date ou l’heure demandée en fonction de l’heure UTC (Universal Coordinated Time).

L' `Date` objet fournit également d’autres méthodes telles que [GETDATE](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [GetHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)et [ToString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), qui renvoient la date ou l’heure demandée en heure locale.

Le concept d’« heure locale » est principalement déterminé par le navigateur et le système d’exploitation de l’ordinateur client. Par exemple, dans la plupart des navigateurs qui s’exécutent sur un ordinateur client Windows, un appel JavaScript à `getDate` , renvoie une date basée sur le fuseau horaire défini dans Windows sur l’ordinateur client.

L’exemple suivant crée un `Date` objet `myLocalDate` en heure locale, et appelle `toUTCString` pour convertir cette date en une chaîne de date au format UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Bien que vous puissiez utiliser l' `Date` objet JavaScript pour obtenir une valeur de date ou d’heure basée sur UTC ou sur le fuseau horaire de l’ordinateur client, l’objet **Date** est limité à un égard : il ne fournit pas de méthode pour renvoyer une valeur de date ou d’heure pour un autre fuseau horaire spécifique. Par exemple, si votre ordinateur client est défini sur l’heure de l’est (est), il n’existe pas de `Date` méthode qui vous permet d’obtenir la valeur horaire autre que est ou UTC, telle que Pacifique (PST).


## <a name="date-related-features-for-outlook-add-ins"></a>Fonctionnalités liées à la date pour les compléments Outlook

La limitation JavaScript susmentionnée a une implication pour vous, lorsque vous utilisez l’API JavaScript pour Office pour gérer les valeurs de date ou d’heure dans les compléments Outlook qui s’exécutent dans un client riche Outlook, et dans Outlook sur le Web ou les appareils mobiles.


### <a name="time-zones-for-outlook-clients"></a>Fuseaux horaires pour les clients Outlook

Pour clarifier les choses, définissons les fuseaux horaires en question.

|**Fuseau horaire**|**Description**|
|:-----|:-----|
|Fuseau horaire de l’ordinateur client|Cette valeur est définie sur le système d’exploitation de l’ordinateur client. La plupart des navigateurs utilisent le fuseau horaire de l’ordinateur client pour afficher les valeurs de date ou d’heure de l' `Date` objet JavaScript.<br/><br/>Le client riche Outlook utilise ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur. <br/><br/>Par exemple, sur un ordinateur client exécutant Windows, Outlook utilise le fuseau horaire défini sur Windows comme fuseau horaire local. Sur Mac, si l’utilisateur modifie le fuseau horaire sur l’ordinateur client, Outlook sur Mac invite également l’utilisateur à mettre à jour le fuseau horaire dans Outlook.|
|Fuseau horaire EAC (Exchange Admin Center)|L’utilisateur définit cette valeur de fuseau horaire (et la langue préférée) lorsqu’il se connecte à Outlook sur le Web ou les appareils mobiles la première fois. <br/><br/>Outlook sur le Web et les appareils mobiles utilisez ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur.|

Étant donné qu’un client riche Outlook utilise le fuseau horaire de l’ordinateur client et que l’interface utilisateur d’Outlook sur le Web et les appareils mobiles utilise le fuseau horaire du centre d’administration Exchange, l’heure locale pour le même complément installé pour la même boîte aux lettres peut être différente lors de l’exécution dans un client riche Outlook et dans Outlook sur le Web ou les appareils mobiles. En tant que développeur de complément Outlook, vous devez entrer et sortir de façon appropriée les valeurs de date afin qu’elles soient toujours en accord avec le fuseau horaire attendu par l’utilisateur sur le client correspondant.


### <a name="date-related-api"></a>API liée à la date

Les propriétés et méthodes suivantes de l’API JavaScript pour Office prennent en charge les fonctionnalités liées à la date.

**Membre de l'API**|**Représentation du fuseau horaire**|**Exemple dans un client riche Outlook**|**Exemple dans Outlook sur le Web ou les appareils mobiles**
--------------|----------------------------|-------------------------------------|-------------------
[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone)|Dans un client riche Outlook, cette propriété renvoie le fuseau horaire de l’ordinateur client. Dans Outlook sur le Web et les appareils mobiles, cette propriété renvoie le fuseau horaire du centre d’administration Exchange. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Chacune de ces propriétés renvoie un `Date` objet JavaScript. Cette `Date` valeur est correcte (UTC), comme indiqué dans l’exemple suivant- `myUTCDate` a la même valeur dans un client riche Outlook, Outlook sur le Web et les appareils mobiles.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Toutefois, l’appel  `myDate.getDate` renvoie une valeur de date dans le fuseau horaire de l’ordinateur client, qui est cohérente avec le fuseau horaire utilisé pour afficher les valeurs de date et d’heure dans l’interface client riche Outlook, mais peut être différent du fuseau horaire du centre d’administration Exchange sur le Web et les appareils mobiles utilisés dans son interface utilisateur.|Si l’élément est créé à 9 h 00 UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h 00 UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` renvoie 6 h 00 EST.|Si l’élément est créé à 9 h 00 UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h 00 UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` renvoie 6 h 00 EST.<br/><br/>Notez que si vous souhaitez afficher l’heure de création ou de modification dans l’interface utilisateur, vous pouvez d’abord convertir l’heure au format PST pour rester cohérent avec le reste de l’interface utilisateur.
[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Chacun des paramètres de  _début_ et de _fin_ requiert un `Date` objet JavaScript. Les arguments doivent être au format UTC, quel que soit le fuseau horaire utilisé dans l’interface utilisateur d’un client riche Outlook, ou Outlook sur le Web ou les appareils mobiles.|Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>|Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a>Méthodes d’assistance pour les scénarios liés à la date


Comme décrit dans les sections précédentes, étant donné que la « durée locale » pour un utilisateur dans Outlook sur le Web ou les appareils mobiles peut être différente sur un client riche Outlook, mais que l’objet JavaScript **Date** prend uniquement en charge la conversion vers le fuseau horaire de l’ordinateur client ou l’UTC, l’API JavaScript Office fournit deux méthodes d’assistance : [Office. Context. Mailbox. convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) et [Office](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)

Ces méthodes d’assistance ont besoin de gérer la date ou l’heure différemment pour les deux scénarios de date suivants, dans un client riche Outlook, Outlook sur le Web et les appareils mobiles, renforçant ainsi « l’écriture unique » pour les différents clients de votre complément.


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Scénario A : affichage de l’heure de création ou de modification d’un élément

Si vous affichez l’heure de création de l’élément ( `Item.dateTimeCreated` ) ou l’heure de modification ( `Item.dateTimeModified` dans l’interface utilisateur, utilisez `convertToLocalClientTime` d’abord pour convertir l' `Date` objet fourni par ces propriétés afin d’obtenir une représentation de dictionnaire dans l’heure locale appropriée. Affichez ensuite les parties de la date de dictionnaire. L’exemple suivant illustre ce scénario :


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

Notez que `convertToLocalClientTime` prend en charge la différence entre un client riche Outlook et Outlook sur le Web ou les appareils mobiles :


- Si `convertToLocalClientTime` détecte que l’application actuelle est un client riche, la méthode convertit la `Date` représentation en une représentation de dictionnaire dans le même fuseau horaire d’ordinateur client, conformément au reste de l’interface utilisateur du client riche.
    
- Si `convertToLocalClientTime` détecte que l’application active est Outlook sur le Web ou sur des appareils mobiles, la méthode convertit la représentation UTC (UTC) en `Date` un format de dictionnaire dans le fuseau horaire d’un centre d’administration Exchange, cohérent avec le reste de l’interface utilisateur d’Outlook sur le Web ou sur les appareils mobiles.
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Scénario B : affichage des dates de début et de fin dans un formulaire de nouveau rendez-vous

Si vous obtenez en entrée des parties différentes d’une valeur de date et d’heure représentée dans l’heure locale et que vous souhaitez fournir cette valeur d’entrée de dictionnaire comme heure de début ou de fin dans un formulaire de rendez-vous, utilisez d’abord la `convertToUtcClientTime` méthode d’assistance pour convertir la valeur de dictionnaire en objet UTC correct `Date` .

Dans l’exemple suivant, supposons que  `myLocalDictionaryStartDate` et `myLocalDictionaryEndDate` sont des valeurs de date et d’heure au format de dictionnaire que vous avez obtenues auprès de l’utilisateur. Ces valeurs sont basées sur l’heure locale, dépendante de la plateforme cliente.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Les valeurs qui en résultent, `myUTCCorrectStartDate` et `myUTCCorrectEndDate`, sont au format UTC. Transmettez ensuite ces `Date` objets comme arguments pour les paramètres de _début_ et de _fin_ de la `Mailbox.displayNewAppointmentForm` méthode pour afficher le nouveau formulaire de rendez-vous.

Notez que `convertToUtcClientTime` prend en charge la différence entre un client riche Outlook et Outlook sur le Web ou les appareils mobiles :


- Si `convertToUtcClientTime` détecte que l’application active est un client riche Outlook, la méthode convertit simplement la représentation du dictionnaire en `Date` objet. Cet `Date` objet est conforme au format UTC, comme attendu par `displayNewAppointmentForm` .
    
- Si `convertToUtcClientTime` détecte que l’application active est Outlook sur le Web ou sur des appareils mobiles, la méthode convertit le format de dictionnaire des valeurs de date et d’heure exprimées dans le fuseau horaire du centre d’administration Exchange en un `Date` objet. Cet `Date` objet est conforme au format UTC, comme attendu par `displayNewAppointmentForm` .
    
## <a name="see-also"></a>Voir aussi

- [Déployer et installer des compléments Outlook à des fins de test](testing-and-tips.md)
