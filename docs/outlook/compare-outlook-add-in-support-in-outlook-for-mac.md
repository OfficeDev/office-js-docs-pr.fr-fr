---
title: Comparer la prise en charge des compléments Outlook dans Outlook sur Mac
description: Découvrez comment la prise en charge des compléments dans Outlook sur Mac est comparée à celle des autres clients Outlook.
ms.date: 06/04/2020
localization_priority: Normal
ms.openlocfilehash: a1eb51ed5b8fa51283b738bc7522b1cf4eb16169
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608970"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Comparaison de la prise en charge des compléments Outlook dans Outlook sur Mac avec d’autres clients Outlook

Vous pouvez créer et exécuter un complément Outlook de la même manière dans Outlook sur Mac que dans les autres clients, y compris Outlook sur le Web, Windows, iOS et Android, sans personnaliser le code JavaScript pour chaque client. Les mêmes appels à partir du complément vers l’API JavaScript Office fonctionnent généralement de la même manière, à l’exception des zones décrites dans le tableau suivant.

Pour plus d'informations, voir [Déployer et installer des compléments Outlook à des fins de test](testing-and-tips.md).

Pour plus d’informations sur la prise en charge de la nouvelle interface utilisateur sur Mac, consultez la rubrique [New Outlook sur Mac](#new-outlook-on-mac-preview).

| Domaine | Outlook sur le Web, Windows et les appareils mobiles | Outlook sur Mac |
|:-----|:-----|:-----|
| Versions d’office.js et du schéma de manifeste des Compléments Office pris en charge | Toutes les API dans Office.js et le schéma version 1.1. | Toutes les API dans Office.js et le schéma version 1.1.<br><br>**Remarque**: dans Outlook sur Mac, seul Build 16.35.308 ou version ultérieure prend en charge l’enregistrement d’une réunion. Dans le cas contraire, la `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition. Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745). |
| Instances d’une série de rendez-vous périodiques | <ul><li>Peut obtenir l’ID d’élément et d’autres propriétés d’un rendez-vous principal ou d’une instance de rendez-vous d’une série périodique.</li><li>peut utiliser [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour afficher une instance ou le masque d’une série périodique.</li></ul> | <ul><li>Peut obtenir l’ID d’élément et d’autres propriétés du rendez-vous principal, mais pas ceux d’une instance d’une série périodique.</li><li>Peut afficher le rendez-vous principal d’une série périodique. Sans l’ID d’élément, ne peut pas afficher une instance d’une série périodique.</li></ul> |
| Type de destinataire d’un participant de rendez-vous | Peut utiliser [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) pour identifier le type de destinataire d’un participant. | `EmailAddressDetails.recipientType` Renvoie `undefined` pour les participants à un rendez-vous. |
| Chaîne de version du client hôte | Le format de la chaîne de version renvoyée par [Diagnostics. hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) dépend du type de client réel. Par exemple :<ul><li>Outlook sur Windows :`15.0.4454.1002`</li><li>Outlook sur le Web :`15.0.918.2`</li></ul> |Exemple de la chaîne de version renvoyée par `Diagnostics.hostVersion` sur Outlook sur Mac :`15.0 (140325)` |
| Propriétés personnalisées d’un élément | Si le réseau tombe en panne, un complément peut toujours accéder aux propriétés personnalisées mises en cache. | Étant donné qu’Outlook sur Mac ne met pas en cache les propriétés personnalisées, si le réseau tombe en panne, les compléments ne pourront pas y accéder. |
| Détails des pièces jointes | Le type de contenu et les noms de pièces jointes dans un objet [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) dépendent du type de client :<ul><li>Exemple JSON de `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` ne contient aucune extension de nom de fichier. Par exemple, si la pièce jointe est un message dont l’objet est « RE: Summer activity », l’objet JSON qui représente le nom de la pièce jointe serait `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Exemple JSON de `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` inclut toujours une extension de nom de fichier. Les pièces jointes qui sont des éléments de messagerie ont une extension .eml et les rendez-vous ont une extension .ics. Par exemple, si une pièce jointe est un message électronique dont l’objet est « RE: Summer activity », l’objet JSON qui représente le nom de pièce jointe sera `"name": "RE: Summer activity.eml"`<p>**REMARQUE** : si un fichier est joint par programmation (par exemple, par le biais d’un complément) sans extension, `AttachmentDetails.name` ne contient pas l’extension dans le nom de fichier.</p></li></ul> |
| Chaîne représentant le fuseau horaire dans les propriétés `dateTimeCreated` et `dateTimeModified` |Par exemple : `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Par exemple : `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Précision horaire de `dateTimeCreated` et `dateTimeModified` | Si un complément utilise le code suivant, la précision est de l’ordre de la milliseconde :<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| La précision peut seulement atteindre une seconde. |

## <a name="new-outlook-on-mac-preview"></a>Nouvelle version d’Outlook sur Mac (aperçu)

Les compléments Outlook sont désormais pris en charge dans la nouvelle interface utilisateur Mac, jusqu’à l’ensemble de conditions requises 1,6. Toutefois, les ensembles de conditions requises et les fonctionnalités suivantes ne sont **pas** encore pris en charge.

1. Ensembles de conditions requises de l’API 1,7 et 1,8
1. Volet Office épinglables, `ItemChanged` événement
1. Compléments contextuels
1. En envoi
1. Prise en charge des dossiers partagés
1. `saveAsync`lors de la composition d’une réunion
1. Authentification unique (SSO)

Nous vous invitons à prévisualiser la nouvelle version d’Outlook sur Mac, disponible à partir de la version 16.38.506. Pour en savoir plus sur la façon de le tester, consultez la rubrique [relative aux notes de publication d’Outlook pour Mac pour les générations rapides Insiders](https://support.microsoft.com/office/d6347358-5613-433e-a49e-a9a0e8e0462a).

Vous pouvez déterminer la version de l’interface utilisateur sur laquelle vous vous trouvez, comme suit.

**Interface utilisateur actuelle**

&nbsp;&nbsp;&nbsp;&nbsp;![Interface utilisateur actuelle sur Mac](../images/outlook-on-mac-classic.png)

**Nouvelle interface utilisateur (aperçu)**

&nbsp;&nbsp;&nbsp;&nbsp;![Nouvelle interface utilisateur en mode aperçu sur Mac](../images/outlook-on-mac-new.png)