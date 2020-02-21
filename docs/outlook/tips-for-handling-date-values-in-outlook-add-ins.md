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
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Conseils pour la gestion des valeurs de date dans les compléments Outlook

L’interface API JavaScript pour Office utilise l’objet JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) pour stocker et récupérer la plupart des dates et des heures. 

Cet objet **Date** fournit des méthodes telles que [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) et [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), qui renvoient la date ou l’heure UTC demandée.

L’objet **Date** fournit également d’autres méthodes telles que [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) et [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), qui renvoient la date ou l’heure locale demandée.

Le concept d’« heure locale » est principalement déterminé par le navigateur et le système d’exploitation de l’ordinateur client. Par exemple, dans la plupart des navigateurs s’exécutant sur un ordinateur client Windows, un appel JavaScript à **getDate** renvoie une date en fonction du fuseau horaire défini dans Windows sur l’ordinateur client.

L’exemple suivant crée un objet **Date** `myLocalDate` au format de l’heure locale, et appelle **toUTCString** pour convertir cette date en chaîne de date au format UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Si vous pouvez utiliser le code JavaScript **Date** pour obtenir une valeur de date ou l’heure en fonction de UTC ou le fuseau horaire d’ordinateur client, l’objet **Date** est limité à un égard : il ne fournit pas de méthodes pour renvoyer une date ou valeur de temps pour n’importe quel autre fuseau horaire. Par exemple, si votre ordinateur client est défini pour être en horaire Standard est (EST), il n’existe aucune méthode**Date** qui vous permet d’obtenir la valeur d’heure autre que dans h EST ou UTC, comme par exemple l’heure du Pacifique (PST).


## <a name="date-related-features-for-outlook-add-ins"></a>Fonctionnalités liées à la date pour les compléments Outlook

La limitation JavaScript mentionnée ci-dessus a une implication, lorsque vous utilisez l’API JavaScript pour Office pour gérer les valeurs de date ou d’heure dans les compléments Outlook qui s’exécutent dans un client riche Outlook, ainsi que dans Outlook sur le Web ou les appareils mobiles.


### <a name="time-zones-for-outlook-clients"></a>Fuseaux horaires pour les clients Outlook

Pour clarifier les choses, définissons les fuseaux horaires en question.

|**Fuseau horaire**|**Description**|
|:-----|:-----|
|Fuseau horaire de l’ordinateur client|Ce champ est défini sur le système d’exploitation de l’ordinateur client. La plupart des navigateurs utilisent le fuseau horaire de l’ordinateur client pour afficher les valeurs de date ou d’heure de l’objet JavaScript **Date**.  <br/><br/>Le client Outlook utilise ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur. <br/><br/>Par exemple, sur un ordinateur client exécutant Windows, Outlook utilise le fuseau horaire défini sur Windows comme fuseau horaire local. Sur Mac, si l’utilisateur modifie le fuseau horaire sur l’ordinateur client, Outlook sur Mac invite également l’utilisateur à mettre à jour le fuseau horaire dans Outlook.|
|Fuseau horaire EAC (Exchange Admin Center)|L’utilisateur définit cette valeur de fuseau horaire (et la langue préférée) lorsqu’il se connecte à Outlook sur le Web ou les appareils mobiles la première fois. <br/><br/>Outlook sur le Web et les appareils mobiles utilisez ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur.|

Étant donné qu’un client riche Outlook utilise le fuseau horaire de l’ordinateur client et que l’interface utilisateur d’Outlook sur le Web et les appareils mobiles utilise le fuseau horaire du centre d’administration Exchange, l’heure locale pour le même complément installé pour la même boîte aux lettres peut être différente lors de l’exécution dans une Clie riche Outlook NT et dans Outlook sur le Web ou les appareils mobiles. En tant que développeur de complément Outlook, vous devez entrer et sortir de façon appropriée les valeurs de date afin qu’elles soient toujours en accord avec le fuseau horaire attendu par l’utilisateur sur le client correspondant.


### <a name="date-related-api"></a>API liée à la date

Les propriétés et méthodes suivantes de l’API JavaScript pour Office prennent en charge des fonctionnalités associées à la date.

**Membre de l'API**|**Représentation du fuseau horaire**|**Exemple dans un client riche Outlook**|**Exemple dans Outlook sur le Web ou les appareils mobiles**
--------------|----------------------------|-------------------------------------|-------------------
[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone)|Dans un client riche Outlook, cette propriété renvoie le fuseau horaire de l’ordinateur client. Dans Outlook sur le Web et les appareils mobiles, cette propriété renvoie le fuseau horaire du centre d’administration Exchange. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Chacune de ces propriétés renvoie un objet JavaScript **Date**. Cette valeur de **Date** est au format UTC, comme indiqué dans l’exemple suivant `myUTCDate` : a la même valeur dans un client riche Outlook, Outlook sur le Web et les appareils mobiles.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Toutefois, l' `myDate.getDate` appel renvoie une valeur de date dans le fuseau horaire de l’ordinateur client, qui est cohérente avec le fuseau horaire utilisé pour afficher les valeurs de date et d’heure dans l’interface client riche Outlook, mais peut être différent du fuseau horaire du centre d’administration Exchange sur le Web et les appareils mobiles utilisés dans son interface utilisateur.|Si l’élément est créé à 9 h 00 UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h 00 UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` renvoie 6h00 EST.|Si l’élément est créé à 9 h 00 UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h 00 UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` renvoie 6 h 00 EST.<br/><br/>Notez que si vous souhaitez afficher l’heure de création ou de modification dans l’interface utilisateur, vous pouvez d’abord convertir l’heure au format PST pour rester cohérent avec le reste de l’interface utilisateur.
[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Chacun des paramètres  _Start_ et _End_ nécessite un objet JavaScript **Date**. Les arguments doivent être au format UTC, quel que soit le fuseau horaire utilisé dans l’interface utilisateur d’un client riche Outlook, ou Outlook sur le Web ou les appareils mobiles.|Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>|Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a>Méthodes d’assistance pour les scénarios liés à la date


Comme décrit dans les sections précédentes, étant donné que la « durée locale » pour un utilisateur dans Outlook sur le Web ou les appareils mobiles peut être différente sur un client riche Outlook, mais que l’objet JavaScript **Date** prend en charge la conversion uniquement du fuseau horaire de l’ordinateur client ou de l’heure UTC, l’interface API JavaScript pour Office fournit deux [](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)méthodes d’assistance : [Office](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)

Ces méthodes d’assistance ont besoin de gérer la date ou l’heure différemment pour les deux scénarios de date suivants, dans un client riche Outlook, Outlook sur le Web et les appareils mobiles, renforçant ainsi « l’écriture unique » pour les différents clients de votre complément.


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Scénario A : affichage de l’heure de création ou de modification d’un élément

Si vous affichez l’heure de création (**Item.dateTimeCreated**) ou de modification (**Item.dateTimeModified**) d’un élément dans l’interface utilisateur, utilisez d’abord **convertToLocalClientTime** pour convertir l’objet **Date** fourni par ces propriétés pour obtenir une représentation de dictionnaire dans l’heure locale appropriée.  Affichez ensuite les parties de la date de dictionnaire.  L’exemple suivant illustre ce scénario :


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

Notez que **convertToLocalClientTime** prend en charge la différence entre un client riche Outlook et Outlook sur le Web ou les appareils mobiles :


- Si **convertToLocalClientTime** détecte que l’hôte actuel est un client riche, la méthode convertit la représentation **Date** en une représentation de dictionnaire dans le fuseau horaire de l’ordinateur client, en accord avec le reste de l’interface utilisateur du client riche.
    
- Si **convertToLocalClientTime** détecte que l’hôte actuel est Outlook sur le Web ou les appareils mobiles, la méthode convertit la représentation de **Date** UTC correcte en un format de dictionnaire dans le fuseau horaire d’un centre d’administration Exchange, cohérent avec le reste de l’interface utilisateur d’Outlook sur le Web ou sur les appareils mobiles.
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Scénario B : affichage des dates de début et de fin dans un formulaire de nouveau rendez-vous

Si vous obtenez différentes parties d’une valeur d’entrée date-heure à l’heure locale et que vous souhaitez fournir la valeur d’entrée du dictionnaire sous la forme d’une heure de début ou de fin dans un formulaire de rendez-vous, utilisez d’abord la méthode d’assistance **convertToUtcClientTime** pour convertir la valeur de dictionnaire en objet **Date** au format UTC.

Dans l’exemple suivant, supposons que  `myLocalDictionaryStartDate` et `myLocalDictionaryEndDate` sont des valeurs de date et d’heure au format de dictionnaire que vous avez obtenues auprès de l’utilisateur. Ces valeurs sont basées sur l’heure locale, qui dépend elle-même de l’application hôte.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Les valeurs qui en résultent, `myUTCCorrectStartDate` et `myUTCCorrectEndDate`, sont au format UTC. Transférez ensuite ces objets **Date** en tant qu’arguments pour les paramètres_Start_ et _End_ de la méthode **Mailbox.displayNewAppointmentForm** pour afficher le nouveau formulaire de rendez-vous. 

Notez que **convertToUtcClientTime** prend en charge la différence entre un client riche Outlook et Outlook sur le Web ou les appareils mobiles :


- Si **convertToUtcClientTime** détecte que l’hôte actuel est un client riche Outlook, la méthode convertit simplement la représentation de dictionnaire en objet **Date**.  Cet objet **Date** est conforme au format UTC, comme attendu par **displayNewAppointmentForm**.
    
- Si **convertToUtcClientTime** détecte que l’hôte actuel est Outlook sur le Web ou les appareils mobiles, la méthode convertit le format de dictionnaire des valeurs de date et d’heure exprimées dans le fuseau horaire du centre d’administration Exchange en un objet **Date** . Cet objet **Date** est conforme au format UTC, comme attendu par **displayNewAppointmentForm**.
    

## <a name="see-also"></a>Voir aussi

- [Déployer et installer des compléments Outlook à des fins de test](testing-and-tips.md)
    


