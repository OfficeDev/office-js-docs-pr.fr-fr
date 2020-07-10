---
title: Compléments Outlook pour Outlook Mobile
description: Les compléments Outlook Mobile sont pris en charge sur tous les comptes professionnels de Microsoft 365, les comptes Outlook.com et le support sont bientôt disponibles dans les comptes gmail.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 34fbb01d596c4da38fe81438088cd71d8c7e152a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093895"
---
# <a name="add-ins-for-outlook-mobile"></a>Compléments pour Outlook Mobile

Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.

Les compléments Outlook Mobile sont pris en charge sur tous les comptes professionnels de Microsoft 365, les comptes Outlook.com et le support sont bientôt disponibles dans les comptes gmail.

**Exemple de volet Office dans Outlook sur iOS**

![Capture d’écran d’un volet Office dans Outlook sur iOS](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Exemple de volet Office dans Outlook sur Android**

![Capture d’écran d’un volet Office dans Outlook sur Android](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> Les compléments ne fonctionnent pas dans la version moderne d’Outlook dans un navigateur mobile. Pour plus d’informations, consultez [la rubrique Outlook sur votre navigateur mobile est en cours de mise à niveau](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).

## <a name="whats-different-on-mobile"></a>Qu’est-ce qui est différent sur mobile ?

- The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.
    - Le complément **DOIT** respecter les [instructions concernant l’interface utilisateur](outlook-addin-design.md).
    - Le scénario du complément **DOIT** [être pertinent sur mobile](#what-makes-a-good-scenario-for-mobile-add-ins).

- En règle générale, seul le mode lecture de message est pris en charge pour le moment. Cela signifie qu’il `MobileMessageReadCommandSurface` s’agit du seul [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) que vous devez déclarer dans la section mobile de votre manifeste. Toutefois, le mode organisateur de rendez-vous est pris en charge pour les compléments intégrés au fournisseur de réunions en ligne qui déclarent le [point d’extension MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview). Pour plus d’informations sur ce scénario, reportez-vous à l’article [créer un complément Outlook Mobile pour un fournisseur de réunion en ligne](online-meeting.md) .

- The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).

- Lorsque vous soumettez votre complément dans le magasin avec l’élément [MobileFormFactor](../reference/manifest/mobileformfactor.md) dans le manifeste, vous devez accepter notre addendum pour les développeurs de compléments sur iOS, et envoyer votre ID de développeur Apple pour vérification.

- Enfin, votre manifeste devra déclarer l’élément `MobileFormFactor`, et inclure les types de [contrôles](../reference/manifest/control.md) et de [tailles d’icône](../reference/manifest/icon.md) corrects.

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>Qu’est-ce qu’un bon scénario pour les compléments mobiles ?

Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.

Voici quelques exemples de scénarios pertinents dans Outlook Mobile.

- The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.

- The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.

**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur iOS**

![Image GIF animée montrant l’interaction d’un utilisateur avec un complément Outlook Mobile sur iOS](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur Android**

![Image GIF animée montrant l’interaction d’un utilisateur avec un complément Outlook Mobile sur Android](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>Test de vos compléments sur mobile

To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.

After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.

Le dépannage sur mobile peut être difficile dans la mesure où vous ne disposez pas des outils que vous utilisez. Toutefois, une option de résolution des problèmes sur iOS consiste à utiliser Fiddler (consultez [ce didacticiel sur son utilisation avec un appareil iOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment :

- [ajouter la prise en charge mobile au manifeste de votre complément](add-mobile-support.md),
- [concevoir une expérience mobile exceptionnelle pour votre complément](outlook-addin-design.md),
- [obtenir un jeton d’accès et appeler des API REST Outlook](use-rest-api.md) à partir de votre complément.
