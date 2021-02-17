---
title: Créer un application mobile Outlook pour un fournisseur de réunion en ligne
description: Explique comment configurer un application mobile Outlook pour un fournisseur de services de réunion en ligne.
ms.topic: article
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: fb98ddeeef8615476659a0abb798ea7901d81248
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270741"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a><span data-ttu-id="3df87-103">Créer un application mobile Outlook pour un fournisseur de réunion en ligne</span><span class="sxs-lookup"><span data-stu-id="3df87-103">Create an Outlook mobile add-in for an online-meeting provider</span></span>

<span data-ttu-id="3df87-104">La configuration d’une réunion en ligne est une expérience essentielle pour un utilisateur Outlook et il est facile de créer une réunion [Teams avec Outlook](/microsoftteams/teams-add-in-for-outlook) Mobile.</span><span class="sxs-lookup"><span data-stu-id="3df87-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="3df87-105">Toutefois, la création d’une réunion en ligne dans Outlook avec un service non-Microsoft peut être fastidieuse.</span><span class="sxs-lookup"><span data-stu-id="3df87-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="3df87-106">En implémentant cette fonctionnalité, les fournisseurs de services peuvent simplifier l’expérience de création de réunions en ligne pour les utilisateurs de leurs add-in Outlook.</span><span class="sxs-lookup"><span data-stu-id="3df87-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3df87-107">Cette fonctionnalité est uniquement prise en charge sur Android et iOS avec un abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="3df87-107">This feature is only supported on Android and iOS with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="3df87-108">Dans cet article, vous allez apprendre à configurer votre application mobile Outlook pour permettre aux utilisateurs d’organiser et de participer à une réunion à l’aide de votre service de réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="3df87-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="3df87-109">Tout au long de cet article, nous allons utiliser un fournisseur fictif de services de réunion en ligne, « Contoso ».</span><span class="sxs-lookup"><span data-stu-id="3df87-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="3df87-110">Configuration de votre environnement</span><span class="sxs-lookup"><span data-stu-id="3df87-110">Set up your environment</span></span>

<span data-ttu-id="3df87-111">Terminez [le démarrage rapide d’Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de compl?ment avec le générateur Yeoman pour les compl?ments Office.</span><span class="sxs-lookup"><span data-stu-id="3df87-111">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="3df87-112">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="3df87-112">Configure the manifest</span></span>

<span data-ttu-id="3df87-113">Pour permettre aux utilisateurs de créer des réunions en ligne avec votre application, vous devez configurer le [point d’extension MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) dans le manifeste sous l’élément `MobileFormFactor` parent.</span><span class="sxs-lookup"><span data-stu-id="3df87-113">To enable users to create online meetings with your add-in, you must configure the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="3df87-114">Les autres facteurs de forme ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="3df87-114">Other form factors are not supported.</span></span>

1. <span data-ttu-id="3df87-115">Dans votre éditeur de code, ouvrez le projet de démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="3df87-115">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="3df87-116">Ouvrez **lemanifest.xml** situé à la racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="3df87-116">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="3df87-117">Sélectionnez l’intégralité du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le `<VersionOverrides>` par le code XML suivant.</span><span class="sxs-lookup"><span data-stu-id="3df87-117">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription"></Description>
    <Requirements>
      <bt:Sets>
        <bt:Set Name="Mailbox" MinVersion="1.3"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeGroup">
                <Label resid="residDescription"/>
                <Control xsi:type="Button" id="insertMeetingButton">
                  <Label resid="residLabel"/>
                  <Supertip>
                    <Title resid="residLabel"/>
                    <Description resid="residTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="64" resid="icon-64"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>insertContosoMeeting</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>

        <MobileFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <Control xsi:type="MobileButton" id="insertMeetingButton">
              <Label resid="residLabel"/>
              <Icon>
                <bt:Image size="25" scale="1" resid="icon-16"/>
                <bt:Image size="25" scale="2" resid="icon-16"/>
                <bt:Image size="25" scale="3" resid="icon-16"/>

                <bt:Image size="32" scale="1" resid="icon-32"/>
                <bt:Image size="32" scale="2" resid="icon-32"/>
                <bt:Image size="32" scale="3" resid="icon-32"/>

                <bt:Image size="48" scale="1" resid="icon-48"/>
                <bt:Image size="48" scale="2" resid="icon-48"/>
                <bt:Image size="48" scale="3" resid="icon-48"/>
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
        <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
        <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
        <bt:Image id="icon-64" DefaultValue="https://contoso.com/assets/icon-64.png"/>
        <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residDescription" DefaultValue="Contoso meeting"/>
        <bt:String id="residLabel" DefaultValue="Add a contoso meeting"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residTooltip" DefaultValue="Add a contoso meeting to this appointment."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="3df87-118">Pour en savoir plus sur les manifestes pour les applications [Outlook,](manifests.md) consultez les manifestes des applications Outlook et ajoutez la prise en charge des commandes de module pour [Outlook Mobile.](add-mobile-support.md)</span><span class="sxs-lookup"><span data-stu-id="3df87-118">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="3df87-119">Implémenter l’ajout de détails de réunion en ligne</span><span class="sxs-lookup"><span data-stu-id="3df87-119">Implement adding online meeting details</span></span>

<span data-ttu-id="3df87-120">Dans cette section, découvrez comment votre script de add-in peut mettre à jour la réunion d’un utilisateur pour inclure les détails de la réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="3df87-120">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

1. <span data-ttu-id="3df87-121">À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="3df87-121">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="3df87-122">Remplacez tout le contenu du **fichiercommands.js** par le javaScript suivant.</span><span class="sxs-lookup"><span data-stu-id="3df87-122">Replace the entire content of the **commands.js** file with the following JavaScript.</span></span>

    ```js
    // 1. How to construct online meeting details.
    // Not shown: How to get the meeting organizer's ID and other details from your service.
    const newBody = '<br>' +
        '<a href="https://contoso.com/meeting?id=123456789" target="_blank">Join Contoso meeting</a>' +
        '<br><br>' +
        'Phone Dial-in: +1(123)456-7890' +
        '<br><br>' +
        'Meeting ID: 123 456 789' +
        '<br><br>' +
        'Want to test your video connection?' +
        '<br><br>' +
        '<a href="https://contoso.com/testmeeting" target="_blank">Join test meeting</a>' +
        '<br><br>';

    var mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define a UI-less function named `insertContosoMeeting` (referenced in the manifest)
    //    to update the meeting body with the online meeting details.
    function insertContosoMeeting(event) {
        // Get HTML body from the client.
        mailboxItem.body.getAsync("html",
            { asyncContext: event },
            function (getBodyResult) {
                if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    updateBody(getBodyResult.asyncContext, getBodyResult.value);
                } else {
                    console.error("Failed to get HTML body.");
                    getBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    // 3. How to implement a supporting function `updateBody`
    //    that appends the online meeting details to the current body of the meeting.
    function updateBody(event, existingBody) {
        // Append new body to the existing body.
        mailboxItem.body.setAsync(existingBody + newBody,
            { asyncContext: event, coercionType: "html" },
            function (setBodyResult) {
                if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    setBodyResult.asyncContext.completed({ allowEvent: true });
                } else {
                    console.error("Failed to set HTML body.");
                    setBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    function getGlobal() {
      return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
    }

    const g = getGlobal();

    // The add-in command functions need to be available in global scope.
    g.insertContosoMeeting = insertContosoMeeting;
    ```

## <a name="testing-and-validation"></a><span data-ttu-id="3df87-123">Test et validation</span><span class="sxs-lookup"><span data-stu-id="3df87-123">Testing and validation</span></span>

<span data-ttu-id="3df87-124">Suivez les instructions habituelles [pour tester et valider votre add-in.](testing-and-tips.md)</span><span class="sxs-lookup"><span data-stu-id="3df87-124">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="3df87-125">Après [le chargement d’une](sideload-outlook-add-ins-for-testing.md) version latéral dans Outlook sur le web, Windows ou Mac, redémarrez Outlook sur votre appareil mobile Android.</span><span class="sxs-lookup"><span data-stu-id="3df87-125">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device.</span></span> <span data-ttu-id="3df87-126">(Android est le seul client pris en charge pour le moment.) Ensuite, sur un nouvel écran de réunion, vérifiez que le basculement Microsoft Teams ou Skype est remplacé par le vôtre.</span><span class="sxs-lookup"><span data-stu-id="3df87-126">(Android is the only supported client for now.) Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="3df87-127">Créer une interface utilisateur de réunion</span><span class="sxs-lookup"><span data-stu-id="3df87-127">Create meeting UI</span></span>

<span data-ttu-id="3df87-128">En tant qu’organisateur de réunion, vous devez voir des écrans semblables aux trois images suivantes lorsque vous créez une réunion.</span><span class="sxs-lookup"><span data-stu-id="3df87-128">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="3df87-129">Capture d’écran de la création d’un écran de réunion sur [ ![ Android - Contoso](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) désaggle la capture d’écran de création de réunion sur Android - Chargement de la capture d’écran bascule [ ![ Contoso](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) de création d’un écran de réunion sur [ ![ Android - Contoso](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox) bascule sur</span><span class="sxs-lookup"><span data-stu-id="3df87-129">[![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="3df87-130">Rejoindre l’interface utilisateur de réunion</span><span class="sxs-lookup"><span data-stu-id="3df87-130">Join meeting UI</span></span>

<span data-ttu-id="3df87-131">En tant que participant à la réunion, vous devez voir un écran semblable à l’image suivante lorsque vous affichez la réunion.</span><span class="sxs-lookup"><span data-stu-id="3df87-131">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="3df87-132">[![Capture d’écran de l’écran participer à une réunion sur Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="3df87-132">[![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3df87-133">Si le lien Rejoindre  n’est pas disponible, il se peut que le modèle de réunion en ligne de votre service ne soit pas inscrit sur nos serveurs.</span><span class="sxs-lookup"><span data-stu-id="3df87-133">If you don't see the **Join** link, it may be that the online-meeting template for your service is not registered on our servers.</span></span> <span data-ttu-id="3df87-134">Pour plus [d’informations, consultez](#register-your-online-meeting-template) la section Inscrire votre modèle de réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="3df87-134">See the [Register your online-meeting template](#register-your-online-meeting-template) section for details.</span></span>

## <a name="register-your-online-meeting-template"></a><span data-ttu-id="3df87-135">Inscrire votre modèle de réunion en ligne</span><span class="sxs-lookup"><span data-stu-id="3df87-135">Register your online-meeting template</span></span>

<span data-ttu-id="3df87-136">Si vous souhaitez inscrire le modèle de réunion en ligne pour votre service, vous pouvez créer un problème GitHub avec les détails.</span><span class="sxs-lookup"><span data-stu-id="3df87-136">If you would like to register the online-meeting template for your service, you can create a GitHub issue with the details.</span></span> <span data-ttu-id="3df87-137">Après cela, nous vous contacterons pour coordonner la chronologie d’inscription.</span><span class="sxs-lookup"><span data-stu-id="3df87-137">After that, we'll contact you to coordinate registration timeline.</span></span>

1. <span data-ttu-id="3df87-138">Go to the **Feedback** section at the end of this article.</span><span class="sxs-lookup"><span data-stu-id="3df87-138">Go to the **Feedback** section at the end of this article.</span></span>
1. <span data-ttu-id="3df87-139">Appuyez sur **le lien Cette page.**</span><span class="sxs-lookup"><span data-stu-id="3df87-139">Press the **This page** link.</span></span>
1. <span data-ttu-id="3df87-140">Définissez **le titre** du nouveau problème sur « Enregistrer le modèle de réunion en ligne pour mon service », en remplaçant par le nom de votre `my-service` service.</span><span class="sxs-lookup"><span data-stu-id="3df87-140">Set the **Title** of the new issue to "Register the online-meeting template for my-service", replacing `my-service` with your service name.</span></span>
1. <span data-ttu-id="3df87-141">Dans le corps du problème, remplacez la chaîne « [Entrez vos commentaires ici] » par la chaîne que vous avez définie dans la variable ou une variable similaire de la section Implémenter l’ajout de détails de réunion en ligne plus haut dans cet `newBody` article. [](#implement-adding-online-meeting-details)</span><span class="sxs-lookup"><span data-stu-id="3df87-141">In the issue body, replace the string "[Enter feedback here]" with the string you set in the `newBody` or similar variable from the [Implement adding online meeting details](#implement-adding-online-meeting-details) section earlier in this article.</span></span>
1. <span data-ttu-id="3df87-142">Cliquez **sur Envoyer un nouveau problème.**</span><span class="sxs-lookup"><span data-stu-id="3df87-142">Click **Submit new issue**.</span></span>

![Capture d’écran du nouvel écran de problème GitHub avec un exemple de contenu Contoso](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a><span data-ttu-id="3df87-144">API disponibles</span><span class="sxs-lookup"><span data-stu-id="3df87-144">Available APIs</span></span>

<span data-ttu-id="3df87-145">Les API suivantes sont disponibles pour cette fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="3df87-145">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="3df87-146">API d’organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="3df87-146">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="3df87-147">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Objet](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="3df87-147">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="3df87-148">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Heure](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="3df87-148">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="3df87-149">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Heure](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="3df87-149">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="3df87-150">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="3df87-150">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="3df87-151">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="3df87-151">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="3df87-152">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="3df87-152">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="3df87-153">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span><span class="sxs-lookup"><span data-stu-id="3df87-153">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="3df87-154">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="3df87-154">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="3df87-155">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="3df87-155">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span></span>
- <span data-ttu-id="3df87-156">Gérer le flux d’th</span><span class="sxs-lookup"><span data-stu-id="3df87-156">Handle auth flow</span></span>
  - [<span data-ttu-id="3df87-157">API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="3df87-157">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="3df87-158">Restrictions</span><span class="sxs-lookup"><span data-stu-id="3df87-158">Restrictions</span></span>

<span data-ttu-id="3df87-159">Plusieurs restrictions s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="3df87-159">Several restrictions apply.</span></span>

- <span data-ttu-id="3df87-160">Applicable uniquement aux fournisseurs de services de réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="3df87-160">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="3df87-161">Seuls les add-ins installés par l’administrateur apparaissent sur l’écran de composition de réunion, remplaçant l’option Teams ou Skype par défaut.</span><span class="sxs-lookup"><span data-stu-id="3df87-161">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="3df87-162">Les add-ins installés par l’utilisateur ne s’activent pas.</span><span class="sxs-lookup"><span data-stu-id="3df87-162">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="3df87-163">L’icône du add-in doit être en échelles de gris à l’aide de code hexas ou de son équivalent `#919191` dans [d’autres formats de couleur.](https://convertingcolors.com/hex-color-919191.html)</span><span class="sxs-lookup"><span data-stu-id="3df87-163">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="3df87-164">Une seule commande sans interface utilisateur est prise en charge en mode Organisateur de rendez-vous (composition).</span><span class="sxs-lookup"><span data-stu-id="3df87-164">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>

## <a name="see-also"></a><span data-ttu-id="3df87-165">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3df87-165">See also</span></span>

- [<span data-ttu-id="3df87-166">Compléments pour Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="3df87-166">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="3df87-167">Ajouter la prise en charge des commandes de add-in pour Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="3df87-167">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
