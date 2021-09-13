---
title: Comparer Outlook prise en charge des Outlook sur Mac
description: Découvrez comment la prise en charge des Outlook sur Mac se compare à d’Outlook clients.
ms.date: 07/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 9cb9d80d1eff89919b3c33a68b42eb61eca27af3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153172"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Comparer Outlook prise en charge des Outlook sur Mac avec d’autres clients Outlook client

Vous pouvez créer et exécuter un Outlook de la même manière dans Outlook sur Mac que dans les autres clients, notamment Outlook sur le web, Windows, iOS et Android, sans personnaliser le JavaScript pour chaque client. Les mêmes appels du add-in vers l’API JavaScript Office fonctionnent généralement de la même manière, à l’exception des domaines décrits dans le tableau suivant.

Pour plus d'informations, voir [Déployer et installer des compléments Outlook à des fins de test](testing-and-tips.md).

Pour plus d’informations sur la nouvelle prise en charge de l’interface utilisateur, voir la prise en charge des Outlook [sur la nouvelle interface utilisateur Mac.](#add-in-support-in-outlook-on-new-mac-ui-preview)

| Zone | Outlook sur le web, Windows et appareils mobiles | Outlook sur Mac |
|:-----|:-----|:-----|
| Versions d’office.js et du schéma de manifeste des Compléments Office pris en charge | Toutes les API dans Office.js et le schéma version 1.1. | Toutes les API dans Office.js et le schéma version 1.1.<br><br>**REMARQUE**: dans Outlook mac, seule la build 16.35.308 ou ultérieure prend en charge l’enregistrement d’une réunion. Sinon, la `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition. Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745). |
| Instances d’une série de rendez-vous périodiques | <ul><li>Peut obtenir l’ID d’élément et d’autres propriétés d’un rendez-vous principal ou d’une instance de rendez-vous d’une série périodique.</li><li>peut utiliser [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour afficher une instance ou le masque d’une série périodique.</li></ul> | <ul><li>Peut obtenir l’ID d’élément et d’autres propriétés du rendez-vous principal, mais pas ceux d’une instance d’une série périodique.</li><li>Peut afficher le rendez-vous principal d’une série périodique. Sans l’ID d’élément, ne peut pas afficher une instance d’une série périodique.</li></ul> |
| Type de destinataire d’un participant de rendez-vous | Peut utiliser [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipientType) pour identifier le type de destinataire d’un participant. | `EmailAddressDetails.recipientType` Renvoie `undefined` pour les participants à un rendez-vous. |
| Chaîne de version de l’application cliente | Le format de la chaîne de version renvoyée par [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostVersion) dépend du type réel de client. Par exemple :<ul><li>Outlook sur Windows :`15.0.4454.1002`</li><li>Outlook sur le web :`15.0.918.2`</li></ul> |Exemple de chaîne de version renvoyée par `Diagnostics.hostVersion` la Outlook sur Mac :`15.0 (140325)` |
| Propriétés personnalisées d’un élément | Si le réseau tombe en panne, un complément peut toujours accéder aux propriétés personnalisées mises en cache. | Comme Outlook mac ne met pas en cache les propriétés personnalisées, si le réseau est en panne, les macros ne pourront pas y accéder. |
| Détails des pièces jointes | Le type de contenu et les noms des pièces jointes dans un [objet AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) dépendent du type de client :<ul><li>Exemple JSON de `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` ne contient aucune extension de nom de fichier. Par exemple, si la pièce jointe est un message dont l’objet est « RE: Summer activity », l’objet JSON qui représente le nom de la pièce jointe serait `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Exemple JSON de `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` inclut toujours une extension de nom de fichier. Les pièces jointes qui sont des éléments de messagerie ont une extension .eml et les rendez-vous ont une extension .ics. Par exemple, si une pièce jointe est un message électronique dont l’objet est « RE: Summer activity », l’objet JSON qui représente le nom de pièce jointe sera `"name": "RE: Summer activity.eml"`<p>**REMARQUE** : si un fichier est joint par programmation (par exemple, par le biais d’un complément) sans extension, `AttachmentDetails.name` ne contient pas l’extension dans le nom de fichier.</p></li></ul> |
| Chaîne représentant le fuseau horaire dans les propriétés `dateTimeCreated` et `dateTimeModified` |Par exemple : `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Par exemple : `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Précision horaire de `dateTimeCreated` et `dateTimeModified` | Si un complément utilise le code suivant, la précision est de l’ordre de la milliseconde.<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| La précision peut seulement atteindre une seconde. |

## <a name="add-in-support-in-outlook-on-new-mac-ui-preview"></a>Prise en charge des Outlook sur la nouvelle interface utilisateur Mac (prévisualisation)

Outlook sont désormais pris en charge sur la nouvelle interface utilisateur Mac (prévisualisation), jusqu’à l’ensemble de conditions requises 1.8. Toutefois, les ensembles de conditions requises et les fonctionnalités suivants ne **sont pas encore** pris en charge.

- Ensembles de conditions requises de l’API 1.9 et 1.10

Nous vous encourageons à prévisualiser Outlook la nouvelle interface utilisateur Mac, disponible à partir de la version 16.38.506. Pour en savoir plus sur la façon de l’essayer, voir Outlook pour Mac - Notes de publication pour les [builds Insider Fast.](https://support.microsoft.com/office/d6347358-5613-433e-a49e-a9a0e8e0462a)

Vous pouvez déterminer la version de l’interface utilisateur sur laquelle vous vous trouverez, comme suit :

**Interface utilisateur actuelle**

![Interface utilisateur actuelle sur Mac.](../images/outlook-on-mac-classic.png)

**Nouvelle interface utilisateur (aperçu)**

![Nouvelle interface utilisateur en prévisualisation sur Mac.](../images/outlook-on-mac-new.png)
