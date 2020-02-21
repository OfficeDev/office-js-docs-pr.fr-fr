---
title: Compléments Outlook pour Outlook Mobile
description: Les compléments Outlook Mobile sont pris en charge sur tous les comptes commerciaux Office 365 et les comptes Outlook.com. Les comptes Gmail seront pris en charge très bientôt.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 7ede3165f40e644715dc488214e047f00dafbede
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166177"
---
# <a name="add-ins-for-outlook-mobile"></a>Compléments pour Outlook Mobile

Les compléments fonctionnent désormais sur Outlook Mobile, avec les mêmes API que celles disponibles pour d’autres points de terminaison Outlook. Si vous avez déjà créé un complément pour Outlook, il est facile de le faire fonctionner sur Outlook Mobile.

Les compléments Outlook Mobile sont pris en charge sur tous les comptes commerciaux Office 365 et les comptes Outlook.com. Les comptes Gmail seront pris en charge très bientôt.

**Exemple de volet Office dans Outlook sur iOS**

![Capture d’écran d’un volet Office dans Outlook sur iOS](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Exemple de volet Office dans Outlook sur Android**

![Capture d’écran d’un volet Office dans Outlook sur Android](../images/outlook-mobile-addin-taskpane-android.png)

## <a name="whats-different-on-mobile"></a>Qu’est-ce qui est différent sur mobile ?

- La taille réduite et la rapidité des interactions compliquent la conception pour les environnements mobiles. Pour garantir la qualité des expériences pour nos clients, nous définissons des critères de validation stricts qui doivent être respectés par un complément qui déclare prendre en charge les environnements mobiles pour être approuvé dans AppSource.
    - Le complément **DOIT** respecter les [instructions concernant l’interface utilisateur](outlook-addin-design.md).
    - Le scénario du complément **DOIT** [être pertinent sur mobile](#what-makes-a-good-scenario-for-mobile-add-ins).

- Seule la lecture du courrier est prise en charge pour l’instant. Cela signifie que `MobileMessageReadCommandSurface` est le seul élément [ExtensionPoint](../reference/manifest/extensionpoint.md) que vous devez déclarer dans la section mobile de votre manifeste.

- L’API [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) n’est pas prise en charge sur mobile dans la mesure où l’application mobile utilise les API REST pour communiquer avec le serveur. Si le serveur principal de votre application doit se connecter au serveur Exchange, vous pouvez utiliser le jeton de rappel pour émettre des appels d’API REST. Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](use-rest-api.md).

- Lorsque vous soumettez votre complément dans le magasin avec l’élément [MobileFormFactor](../reference/manifest/mobileformfactor.md) dans le manifeste, vous devez accepter notre addendum pour les développeurs de compléments sur iOS, et envoyer votre ID de développeur Apple pour vérification.

- Enfin, votre manifeste devra déclarer l’élément `MobileFormFactor`, et inclure les types de [contrôles](../reference/manifest/control.md) et de [tailles d’icône](../reference/manifest/icon.md) corrects.

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>Qu’est-ce qu’un bon scénario pour les compléments mobiles ?

N’oubliez pas que la durée moyenne d’une session Outlook sur un téléphone est beaucoup plus courte que sur un PC. Cela signifie que votre complément doit être rapide et que le scénario doit permettre à l’utilisateur d’accéder à votre complément, d’en sortir et de traiter ses messages.

Voici quelques exemples de scénarios pertinents dans Outlook Mobile.

- Le complément apporte des informations précieuses dans Outlook et aide les utilisateurs à trier leurs messages et à y répondre correctement. Exemple : un complément CRM qui permet à l’utilisateur de voir les informations client et de partager des informations appropriées.

- Le complément apporte de la valeur ajoutée au contenu des messages de l’utilisateur en enregistrant les informations dans un système de suivi, de collaboration ou de type similaire. Exemple : un complément qui permet aux utilisateurs de transformer les messages électroniques en tâches afin de suivre des projets ou en demandes d’aide pour une équipe de support technique.

**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur iOS**

![Image GIF animée montrant l’interaction d’un utilisateur avec un complément Outlook Mobile sur iOS](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur Android**

![Image GIF animée montrant l’interaction d’un utilisateur avec un complément Outlook Mobile sur Android](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>Test de vos compléments sur mobile

Pour tester un complément sur Outlook Mobile, vous pouvez charger de manière indépendante un complément sur un compte Office 365 ou Outlook.com. Dans Outlook sur le web, accédez à l’icône des paramètres représentée par un engrenage, puis choisissez **Gérer les intégrations** ou **Gérer les compléments**. Près de la partie supérieure, cliquez sur l’emplacement qui indique **Cliquez ici pour ajouter un complément personnalisé** et téléchargez votre manifeste. Vérifiez que votre manifeste est correctement mis en forme et qu’il contient `MobileFormFactor`, sinon il ne sera pas chargé.

Une fois que votre complément fonctionne, testez-le sur différentes tailles d’écran, y compris sur des téléphones et des tablettes. Vous devez vous assurer qu’il respecte les instructions d’accessibilité en matière de contraste, de taille de police et de couleur, et qu’il peut être utilisé avec un lecteur d’écran comme VoiceOver sur iOS ou TalkBack sur Android.

La résolution des problèmes sur mobile peut s’avérer difficile, car vous n’avez peut-être pas les outils auxquels vous êtes habitué. Pour résoudre les problèmes, vous pouvez [utiliser Vorlon.js](../testing/debug-office-add-ins-on-ipad-and-mac.md). Sinon, si vous avez déjà utilisé Fiddler, consultez [ce didacticiel sur son utilisation avec un appareil iOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices).

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment :

- [ajouter la prise en charge mobile au manifeste de votre complément](add-mobile-support.md),
- [concevoir une expérience mobile exceptionnelle pour votre complément](outlook-addin-design.md),
- [obtenir un jeton d’accès et appeler des API REST Outlook](use-rest-api.md) à partir de votre complément.
