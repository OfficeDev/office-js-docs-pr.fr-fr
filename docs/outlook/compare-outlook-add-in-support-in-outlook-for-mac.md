---
title: Comparer la prise en charge des compléments Outlook dans Outlook sur Mac
description: Découvrez comment la prise en charge des compléments dans Outlook sur Mac se compare à d’autres clients Outlook.
ms.date: 10/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: c38e546575446254d54ad13e5d75d997ca6cd6d8
ms.sourcegitcommit: 787fbe4d4a5462ff6679ad7fd00748bf07391610
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68546444"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Comparer la prise en charge des compléments Outlook dans Outlook sur Mac avec d’autres clients Outlook

Vous pouvez créer et exécuter un complément Outlook de la même façon dans Outlook sur Mac que dans les autres clients, notamment Outlook sur le web, Windows, iOS et Android, sans personnaliser javaScript pour chaque client. Les mêmes appels du complément à l’API JavaScript Office fonctionnent généralement de la même façon, à l’exception des zones décrites dans le tableau suivant.

Pour plus d'informations, voir [Déployer et installer des compléments Outlook à des fins de test](testing-and-tips.md).

Pour plus d’informations sur la nouvelle prise en charge de l’interface utilisateur, consultez la [prise en charge des compléments dans Outlook sur la nouvelle interface utilisateur Mac](#add-in-support-in-outlook-on-new-mac-ui).

| Zone | Outlook sur le web, Windows et les appareils mobiles | Outlook sur Mac |
|:-----|:-----|:-----|
| Versions prises en charge de office.js| Toutes les API dans Office.js. | Toutes les API dans Office.js.<br><br>**REMARQUE** : Dans Outlook sur Mac, seule la build 16.35.308 ou ultérieure prend en charge l’enregistrement d’une réunion. Sinon, la `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition. Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745). |
| Instances d’une série de rendez-vous périodiques | <ul><li>Peut obtenir l’ID d’élément et d’autres propriétés d’un rendez-vous principal ou d’une instance de rendez-vous d’une série périodique.</li><li>peut utiliser [mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) pour afficher une instance ou le masque d’une série périodique.</li></ul> | <ul><li>Peut obtenir l’ID d’élément et d’autres propriétés du rendez-vous principal, mais pas ceux d’une instance d’une série périodique.</li><li>Can display the master appointment of a recurring series. Without the item ID, cannot display an instance of a recurring series.</li></ul> |
| Type de destinataire d’un participant de rendez-vous | Peut utiliser [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-recipienttype-member) pour identifier le type de destinataire d’un participant. | `EmailAddressDetails.recipientType` Renvoie `undefined` pour les participants à un rendez-vous. |
| Chaîne de version de l’application cliente | Le format de la chaîne de version retournée par [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) dépend du type réel du client. Par exemple :<ul><li>Outlook sur Windows : `15.0.4454.1002`</li><li>Outlook sur le web :`15.0.918.2`</li></ul> |Exemple de chaîne de version retournée par `Diagnostics.hostVersion` Outlook sur Mac : `15.0 (140325)` |
| Propriétés personnalisées d’un élément | Si le réseau tombe en panne, un complément peut toujours accéder aux propriétés personnalisées mises en cache. | Étant donné qu’Outlook sur Mac ne met pas en cache les propriétés personnalisées, si le réseau tombe en panne, les compléments ne peuvent pas y accéder. |
| Détails des pièces jointes | Le type de contenu et les noms des pièces jointes dans un objet [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) dépendent du type de client :<ul><li>Exemple JSON de `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` does not contain any filename extension. As an example, if the attachment is a message that has the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Exemple JSON de `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` always includes a filename extension. Attachments that are mail items have a .eml extension, and appointments have a .ics extension. As an example, if an attachment is an email with the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity.eml"`.<p>**REMARQUE** : si un fichier est joint par programmation (par exemple, par le biais d’un complément) sans extension, `AttachmentDetails.name` ne contient pas l’extension dans le nom de fichier.</p></li></ul> |
| Chaîne représentant le fuseau horaire dans les propriétés `dateTimeCreated` et `dateTimeModified` |Par exemple : `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Par exemple : `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Précision horaire de `dateTimeCreated` et `dateTimeModified` | Si un complément utilise le code suivant, la précision est de l’ordre de la milliseconde.<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| La précision peut seulement atteindre une seconde. |

## <a name="add-in-support-in-outlook-on-new-mac-ui"></a>Prise en charge des compléments dans Outlook sur la nouvelle interface utilisateur Mac

Les compléments Outlook sont désormais pris en charge dans la nouvelle interface utilisateur Mac (disponible à partir d’Outlook version 16.38.506). Pour connaître les ensembles de conditions requises actuellement pris en charge dans la nouvelle interface utilisateur Mac, consultez la [prise en charge du client de l’ensemble de conditions requises de l’API Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support).

Pour en savoir plus sur la nouvelle interface utilisateur Mac, consultez [La nouvelle Outlook pour Mac](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439).

Vous pouvez déterminer la version de l’interface utilisateur sur laquelle vous utilisez, comme suit :

**Interface utilisateur classique**

![Interface utilisateur classique sur Mac.](../images/outlook-on-mac-classic.png)

**Nouvelle interface utilisateur**

![Nouvelle interface utilisateur sur Mac.](../images/outlook-on-mac-new.png)
