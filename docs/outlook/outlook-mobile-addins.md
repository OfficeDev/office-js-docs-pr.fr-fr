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
# <a name="add-ins-for-outlook-mobile"></a><span data-ttu-id="662b8-103">Compléments pour Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="662b8-103">Add-ins for Outlook Mobile</span></span>

<span data-ttu-id="662b8-p101">Les compléments fonctionnent désormais sur Outlook Mobile, avec les mêmes API que celles disponibles pour d’autres points de terminaison Outlook. Si vous avez déjà créé un complément pour Outlook, il est facile de le faire fonctionner sur Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="662b8-p101">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span></span>

<span data-ttu-id="662b8-106">Les compléments Outlook Mobile sont pris en charge sur tous les comptes commerciaux Office 365 et les comptes Outlook.com. Les comptes Gmail seront pris en charge très bientôt.</span><span class="sxs-lookup"><span data-stu-id="662b8-106">Outlook mobile add-ins are supported on all Office 365 Commercial accounts, Outlook.com accounts, and support is coming soon to Gmail accounts.</span></span>

<span data-ttu-id="662b8-107">**Exemple de volet Office dans Outlook sur iOS**</span><span class="sxs-lookup"><span data-stu-id="662b8-107">**An example task pane in Outlook on iOS**</span></span>

![Capture d’écran d’un volet Office dans Outlook sur iOS](../images/outlook-mobile-addin-taskpane.png)

<br/>

<span data-ttu-id="662b8-109">**Exemple de volet Office dans Outlook sur Android**</span><span class="sxs-lookup"><span data-stu-id="662b8-109">**An example task pane in Outlook on Android**</span></span>

![Capture d’écran d’un volet Office dans Outlook sur Android](../images/outlook-mobile-addin-taskpane-android.png)

## <a name="whats-different-on-mobile"></a><span data-ttu-id="662b8-111">Qu’est-ce qui est différent sur mobile ?</span><span class="sxs-lookup"><span data-stu-id="662b8-111">What's different on mobile?</span></span>

- <span data-ttu-id="662b8-p102">La taille réduite et la rapidité des interactions compliquent la conception pour les environnements mobiles. Pour garantir la qualité des expériences pour nos clients, nous définissons des critères de validation stricts qui doivent être respectés par un complément qui déclare prendre en charge les environnements mobiles pour être approuvé dans AppSource.</span><span class="sxs-lookup"><span data-stu-id="662b8-p102">The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span></span>
    - <span data-ttu-id="662b8-114">Le complément **DOIT** respecter les [instructions concernant l’interface utilisateur](outlook-addin-design.md).</span><span class="sxs-lookup"><span data-stu-id="662b8-114">The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).</span></span>
    - <span data-ttu-id="662b8-115">Le scénario du complément **DOIT** [être pertinent sur mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span><span class="sxs-lookup"><span data-stu-id="662b8-115">The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span></span>

- <span data-ttu-id="662b8-p103">Seule la lecture du courrier est prise en charge pour l’instant. Cela signifie que `MobileMessageReadCommandSurface` est le seul élément [ExtensionPoint](../reference/manifest/extensionpoint.md) que vous devez déclarer dans la section mobile de votre manifeste.</span><span class="sxs-lookup"><span data-stu-id="662b8-p103">Only mail read is supported at this time. That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](../reference/manifest/extensionpoint.md) you should declare in the mobile section of your manifest.</span></span>

- <span data-ttu-id="662b8-p104">L’API [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) n’est pas prise en charge sur mobile dans la mesure où l’application mobile utilise les API REST pour communiquer avec le serveur. Si le serveur principal de votre application doit se connecter au serveur Exchange, vous pouvez utiliser le jeton de rappel pour émettre des appels d’API REST. Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="662b8-p104">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span></span>

- <span data-ttu-id="662b8-121">Lorsque vous soumettez votre complément dans le magasin avec l’élément [MobileFormFactor](../reference/manifest/mobileformfactor.md) dans le manifeste, vous devez accepter notre addendum pour les développeurs de compléments sur iOS, et envoyer votre ID de développeur Apple pour vérification.</span><span class="sxs-lookup"><span data-stu-id="662b8-121">When you submit your add-in to the store with [MobileFormFactor](../reference/manifest/mobileformfactor.md) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.</span></span>

- <span data-ttu-id="662b8-122">Enfin, votre manifeste devra déclarer l’élément `MobileFormFactor`, et inclure les types de [contrôles](../reference/manifest/control.md) et de [tailles d’icône](../reference/manifest/icon.md) corrects.</span><span class="sxs-lookup"><span data-stu-id="662b8-122">Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](../reference/manifest/control.md) and [icon sizes](../reference/manifest/icon.md) included.</span></span>

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a><span data-ttu-id="662b8-123">Qu’est-ce qu’un bon scénario pour les compléments mobiles ?</span><span class="sxs-lookup"><span data-stu-id="662b8-123">What makes a good scenario for mobile add-ins?</span></span>

<span data-ttu-id="662b8-p105">N’oubliez pas que la durée moyenne d’une session Outlook sur un téléphone est beaucoup plus courte que sur un PC. Cela signifie que votre complément doit être rapide et que le scénario doit permettre à l’utilisateur d’accéder à votre complément, d’en sortir et de traiter ses messages.</span><span class="sxs-lookup"><span data-stu-id="662b8-p105">Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span></span>

<span data-ttu-id="662b8-126">Voici quelques exemples de scénarios pertinents dans Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="662b8-126">Here are examples of scenarios that make sense in Outlook Mobile.</span></span>

- <span data-ttu-id="662b8-p106">Le complément apporte des informations précieuses dans Outlook et aide les utilisateurs à trier leurs messages et à y répondre correctement. Exemple : un complément CRM qui permet à l’utilisateur de voir les informations client et de partager des informations appropriées.</span><span class="sxs-lookup"><span data-stu-id="662b8-p106">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.</span></span>

- <span data-ttu-id="662b8-p107">Le complément apporte de la valeur ajoutée au contenu des messages de l’utilisateur en enregistrant les informations dans un système de suivi, de collaboration ou de type similaire. Exemple : un complément qui permet aux utilisateurs de transformer les messages électroniques en tâches afin de suivre des projets ou en demandes d’aide pour une équipe de support technique.</span><span class="sxs-lookup"><span data-stu-id="662b8-p107">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span></span>

<span data-ttu-id="662b8-131">**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur iOS**</span><span class="sxs-lookup"><span data-stu-id="662b8-131">**An example user interaction to create a Trello card from an email message on iOS**</span></span>

![Image GIF animée montrant l’interaction d’un utilisateur avec un complément Outlook Mobile sur iOS](../images/outlook-mobile-addin-interaction.gif)

<br/>

<span data-ttu-id="662b8-133">**Exemple d’interaction utilisateur pour créer une carte Trello à partir d’un message électronique sur Android**</span><span class="sxs-lookup"><span data-stu-id="662b8-133">**An example user interaction to create a Trello card from an email message on Android**</span></span>

![Image GIF animée montrant l’interaction d’un utilisateur avec un complément Outlook Mobile sur Android](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a><span data-ttu-id="662b8-135">Test de vos compléments sur mobile</span><span class="sxs-lookup"><span data-stu-id="662b8-135">Testing your add-ins on mobile</span></span>

<span data-ttu-id="662b8-p108">Pour tester un complément sur Outlook Mobile, vous pouvez charger de manière indépendante un complément sur un compte Office 365 ou Outlook.com. Dans Outlook sur le web, accédez à l’icône des paramètres représentée par un engrenage, puis choisissez **Gérer les intégrations** ou **Gérer les compléments**. Près de la partie supérieure, cliquez sur l’emplacement qui indique **Cliquez ici pour ajouter un complément personnalisé** et téléchargez votre manifeste. Vérifiez que votre manifeste est correctement mis en forme et qu’il contient `MobileFormFactor`, sinon il ne sera pas chargé.</span><span class="sxs-lookup"><span data-stu-id="662b8-p108">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span></span>

<span data-ttu-id="662b8-p109">Une fois que votre complément fonctionne, testez-le sur différentes tailles d’écran, y compris sur des téléphones et des tablettes. Vous devez vous assurer qu’il respecte les instructions d’accessibilité en matière de contraste, de taille de police et de couleur, et qu’il peut être utilisé avec un lecteur d’écran comme VoiceOver sur iOS ou TalkBack sur Android.</span><span class="sxs-lookup"><span data-stu-id="662b8-p109">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span></span>

<span data-ttu-id="662b8-p110">La résolution des problèmes sur mobile peut s’avérer difficile, car vous n’avez peut-être pas les outils auxquels vous êtes habitué. Pour résoudre les problèmes, vous pouvez [utiliser Vorlon.js](../testing/debug-office-add-ins-on-ipad-and-mac.md). Sinon, si vous avez déjà utilisé Fiddler, consultez [ce didacticiel sur son utilisation avec un appareil iOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices).</span><span class="sxs-lookup"><span data-stu-id="662b8-p110">Troubleshooting on mobile can be hard since you may not have the tools you're used to. One option for troubleshooting is to [use Vorlon.js](../testing/debug-office-add-ins-on-ipad-and-mac.md). Or, if you've used Fiddler before, check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices).</span></span>

## <a name="next-steps"></a><span data-ttu-id="662b8-144">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="662b8-144">Next steps</span></span>

<span data-ttu-id="662b8-145">Découvrez comment :</span><span class="sxs-lookup"><span data-stu-id="662b8-145">Learn how to:</span></span>

- <span data-ttu-id="662b8-146">[ajouter la prise en charge mobile au manifeste de votre complément](add-mobile-support.md),</span><span class="sxs-lookup"><span data-stu-id="662b8-146">[Add mobile support to your add-in's manifest](add-mobile-support.md).</span></span>
- <span data-ttu-id="662b8-147">[concevoir une expérience mobile exceptionnelle pour votre complément](outlook-addin-design.md),</span><span class="sxs-lookup"><span data-stu-id="662b8-147">[Design a great mobile experience for your add-in](outlook-addin-design.md).</span></span>
- <span data-ttu-id="662b8-148">[obtenir un jeton d’accès et appeler des API REST Outlook](use-rest-api.md) à partir de votre complément.</span><span class="sxs-lookup"><span data-stu-id="662b8-148">[Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.</span></span>
