---
title: Compléments Outlook pour Outlook Mobile
description: Les compléments mobiles Outlook sont pris en charge sur tous les comptes d’entreprise Microsoft 365 et les comptes Outlook.com.
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: ca09ba550d8d2ed6e9003e85a8d042f413a6ab52
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607561"
---
# <a name="add-ins-for-outlook-mobile"></a>Compléments pour Outlook Mobile

Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.

Les compléments mobiles Outlook sont pris en charge sur tous les comptes d’entreprise Microsoft 365 et les comptes Outlook.com. Toutefois, la prise en charge n’est actuellement pas disponible sur les comptes Gmail.

**Exemple de volet Office dans Outlook sur iOS**

![Capture d’écran d’un volet Office dans Outlook sur iOS.](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Exemple de volet Office dans Outlook sur Android**

![Capture d’écran d’un volet Office dans Outlook sur Android.](../images/outlook-mobile-addin-taskpane-android.png)

## <a name="whats-different-on-mobile"></a>Qu’est-ce qui est différent sur mobile ?

- The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.
  - Le complément **DOIT** respecter les [instructions concernant l’interface utilisateur](outlook-addin-design.md).
  - Le scénario du complément **DOIT** [être pertinent sur mobile](#what-makes-a-good-scenario-for-mobile-add-ins).

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

- En général, seul le mode lecture des messages est pris en charge pour l’instant. Cela signifie qu’il `MobileMessageReadCommandSurface` s’agit du seul [Point d’extension](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface) que vous devez déclarer dans la section mobile de votre manifeste. Toutefois, il existe quelques exceptions :
  1. Le mode Organisateur de rendez-vous est pris en charge pour les compléments intégrés du fournisseur de réunion en ligne qui déclarent à la place le [point d’extension MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface). Pour plus d’informations sur ce scénario, consultez l’article [Créer un complément mobile Outlook pour un fournisseur de réunions en ligne](online-meeting.md) .
  1. Le mode Participant au rendez-vous est pris en charge pour les compléments intégrés créés par les fournisseurs d’applications CRM (Customer Relationship Management). Ces compléments doivent à la place déclarer le [point d’extension MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee). Pour plus d’informations sur ce scénario, consultez les [notes de rendez-vous du journal d’une application externe dans les compléments mobiles Outlook](mobile-log-appointments.md) .

- The [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).

- Lorsque vous soumettez votre complément dans le magasin avec l’élément [MobileFormFactor](/javascript/api/manifest/mobileformfactor) dans le manifeste, vous devez accepter notre addendum pour les développeurs de compléments sur iOS, et envoyer votre ID de développeur Apple pour vérification.

- Enfin, votre manifeste devra déclarer l’élément `MobileFormFactor`, et inclure les types de [contrôles](/javascript/api/manifest/control) et de [tailles d’icône](/javascript/api/manifest/icon) corrects.

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>Qu’est-ce qu’un bon scénario pour les compléments mobiles ?

Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.

Voici quelques exemples de scénarios pertinents dans Outlook Mobile.

- The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.

- The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.

**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur iOS**

![GIF animé montrant l’interaction de l’utilisateur avec un complément Outlook Mobile sur iOS.](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur Android**

![GIF animé montrant l’interaction utilisateur avec un complément Outlook Mobile sur Android.](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>Test de vos compléments sur mobile

Pour tester un complément sur Outlook Mobile, [commencez par charger un complément](sideload-outlook-add-ins-for-testing.md) sur un compte Microsoft 365 ou Outlook.com sur le web, Windows ou Mac. Assurez-vous que votre manifeste est correctement mis en forme pour contenir `MobileFormFactor` ou qu’il ne se charge pas dans votre client Outlook sur mobile.

After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.

La résolution des problèmes sur les appareils mobiles peut être difficile, car vous n’avez peut-être pas les outils auxquels vous êtes habitué. Toutefois, une option de résolution des problèmes sur iOS consiste à utiliser Fiddler (consultez [ce tutoriel sur son utilisation avec un appareil iOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).

> [!NOTE]
> Les Outlook sur le web modernes sur les smartphones iPhone et Android ne sont plus nécessaires ni disponibles pour tester les compléments Outlook. En outre, les compléments ne sont pas pris en charge dans Outlook sur Android, sur iOS et le web mobile moderne avec des comptes Exchange locaux. Certains appareils iOS prennent toujours en charge les compléments lors de l’utilisation de comptes Exchange locaux avec des Outlook sur le web classiques. Pour plus d’informations sur les appareils pris en charge, consultez [Configuration requise pour l’exécution des compléments Office](../concepts/requirements-for-running-office-add-ins.md#client-requirements-non-windows-smartphone-and-tablet).

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment :

- [ajouter la prise en charge mobile au manifeste de votre complément](add-mobile-support.md),
- [concevoir une expérience mobile exceptionnelle pour votre complément](outlook-addin-design.md),
- [obtenir un jeton d’accès et appeler des API REST Outlook](use-rest-api.md) à partir de votre complément.
