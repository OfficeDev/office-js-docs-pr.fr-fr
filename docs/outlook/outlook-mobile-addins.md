---
title: Compléments Outlook pour Outlook Mobile
description: Outlook pour appareils mobiles sont pris en charge sur tous les comptes Microsoft 365 entreprise et Outlook.com.
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 90e88b3b3596f2b11718b9fcf1af7402d7594fe7
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64483991"
---
# <a name="add-ins-for-outlook-mobile"></a>Compléments pour Outlook Mobile

Les compléments fonctionnent désormais sur Outlook Mobile, avec les mêmes API que celles disponibles pour d’autres points de terminaison Outlook. Si vous avez déjà créé un complément pour Outlook, il est facile de le faire fonctionner sur Outlook Mobile.

Outlook pour appareils mobiles sont pris en charge sur tous les comptes Microsoft 365 entreprise et Outlook.com. Toutefois, la prise en charge n’est pas disponible actuellement sur les comptes Gmail.

**Exemple de volet Office dans Outlook sur iOS**

![Capture d’écran d’un volet De tâches Outlook sur iOS.](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Exemple de volet Office dans Outlook sur Android**

![Capture d’écran d’un volet Des tâches Outlook sur Android.](../images/outlook-mobile-addin-taskpane-android.png)

## <a name="whats-different-on-mobile"></a>Qu’est-ce qui est différent sur mobile ?

- La taille réduite et la rapidité des interactions compliquent la conception pour les environnements mobiles. Pour garantir la qualité des expériences pour nos clients, nous définissons des critères de validation stricts qui doivent être respectés par un complément qui déclare prendre en charge les environnements mobiles pour être approuvé dans AppSource.
  - Le complément **DOIT** respecter les [instructions concernant l’interface utilisateur](outlook-addin-design.md).
  - Le scénario du complément **DOIT** [être pertinent sur mobile](#what-makes-a-good-scenario-for-mobile-add-ins).

- En règle générale, seul le mode lecture de message est pris en charge pour le moment. Cela signifie `MobileMessageReadCommandSurface` qu’il s’agit du seul [ExtensionPoint](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface) que vous devez déclarer dans la section mobile de votre manifeste. Toutefois, le mode Organisateur de rendez-vous est pris en charge pour les applications intégrées du fournisseur de réunions en ligne qui déclarent à la place le [point d’extension MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface). Pour plus [d’informations sur ce scénario](online-meeting.md), consultez l’article Créer un Outlook mobile pour un fournisseur de réunion en ligne.

- L’API [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) n’est pas prise en charge sur mobile dans la mesure où l’application mobile utilise les API REST pour communiquer avec le serveur. Si le serveur principal de votre application doit se connecter au serveur Exchange, vous pouvez utiliser le jeton de rappel pour émettre des appels d’API REST. Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](use-rest-api.md).

- Lorsque vous soumettez votre complément dans le magasin avec l’élément [MobileFormFactor](/javascript/api/manifest/mobileformfactor) dans le manifeste, vous devez accepter notre addendum pour les développeurs de compléments sur iOS, et envoyer votre ID de développeur Apple pour vérification.

- Enfin, votre manifeste devra déclarer l’élément `MobileFormFactor`, et inclure les types de [contrôles](/javascript/api/manifest/control) et de [tailles d’icône](/javascript/api/manifest/icon) corrects.

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>Qu’est-ce qu’un bon scénario pour les compléments mobiles ?

N’oubliez pas que la durée moyenne d’une session Outlook sur un téléphone est beaucoup plus courte que sur un PC. Cela signifie que votre complément doit être rapide et que le scénario doit permettre à l’utilisateur d’accéder à votre complément, d’en sortir et de traiter ses messages.

Voici quelques exemples de scénarios pertinents dans Outlook Mobile.

- Le complément apporte des informations précieuses dans Outlook et aide les utilisateurs à trier leurs messages et à y répondre correctement. Exemple : un complément CRM qui permet à l’utilisateur de voir les informations client et de partager des informations appropriées.

- Le complément apporte de la valeur ajoutée au contenu des messages de l’utilisateur en enregistrant les informations dans un système de suivi, de collaboration ou de type similaire. Exemple : un complément qui permet aux utilisateurs de transformer les messages électroniques en tâches afin de suivre des projets ou en demandes d’aide pour une équipe de support technique.

**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur iOS**

![Gif animé montrant l’interaction utilisateur avec un Outlook mobile sur iOS.](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur Android**

![Gif animé montrant l’interaction utilisateur avec un Outlook mobile sur Android.](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>Test de vos compléments sur mobile

Pour tester un compl?ment sur Outlook Mobile, chargez tout d’abord une version test d’un compl?ment vers un compte Microsoft 365 ou Outlook.com sur le web, Windows ou Mac.[](sideload-outlook-add-ins-for-testing.md) Assurez-vous que votre manifeste est correctement formaté `MobileFormFactor` pour contenir ou qu’il ne sera pas chargé dans votre client Outlook sur mobile.

Une fois que votre complément fonctionne, testez-le sur différentes tailles d’écran, y compris sur des téléphones et des tablettes. Vous devez vous assurer qu’il respecte les instructions d’accessibilité en matière de contraste, de taille de police et de couleur, et qu’il peut être utilisé avec un lecteur d’écran comme VoiceOver sur iOS ou TalkBack sur Android.

La résolution des problèmes sur les appareils mobiles peut être difficile, car vous n’avez peut-être pas les outils que vous avez l’habitude d’utiliser. Toutefois, une option de dépannage sur iOS consiste à utiliser Fiddler (consultez ce didacticiel sur son utilisation avec [un appareil iOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).

> [!NOTE]
> Les Outlook sur le web modernes sur iPhone et les smartphones Android ne sont plus nécessaires ou disponibles pour le test Outlook des applications. Pour plus d’informations sur les appareils pris en charge, voir [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md#client-requirements-non-windows-smartphone-and-tablet).

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment :

- [ajouter la prise en charge mobile au manifeste de votre complément](add-mobile-support.md),
- [concevoir une expérience mobile exceptionnelle pour votre complément](outlook-addin-design.md),
- [obtenir un jeton d’accès et appeler des API REST Outlook](use-rest-api.md) à partir de votre complément.
