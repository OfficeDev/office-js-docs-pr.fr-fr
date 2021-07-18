---
title: Créer un Outlook mobile pour un fournisseur de réunion en ligne
description: Explique comment configurer un Outlook mobile pour un fournisseur de services de réunion en ligne.
ms.topic: article
ms.date: 07/09/2021
localization_priority: Normal
ms.openlocfilehash: f0f9b69c2b8b515df3829ca3ba0714393df79fd1
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455501"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a><span data-ttu-id="0ecda-103">Créer un Outlook mobile pour un fournisseur de réunion en ligne</span><span class="sxs-lookup"><span data-stu-id="0ecda-103">Create an Outlook mobile add-in for an online-meeting provider</span></span>

<span data-ttu-id="0ecda-104">La configuration d’une réunion en ligne est une expérience essentielle pour un utilisateur Outlook et il est facile de créer une réunion Teams avec Outlook [mobile.](/microsoftteams/teams-add-in-for-outlook)</span><span class="sxs-lookup"><span data-stu-id="0ecda-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="0ecda-105">Toutefois, la création d’une réunion en Outlook avec un service non-Microsoft peut être fastidieuse.</span><span class="sxs-lookup"><span data-stu-id="0ecda-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="0ecda-106">En implémentant cette fonctionnalité, les fournisseurs de services peuvent simplifier l’expérience de création de réunions en ligne pour Outlook utilisateurs de leur application.</span><span class="sxs-lookup"><span data-stu-id="0ecda-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0ecda-107">Cette fonctionnalité est uniquement prise en charge sur Android et iOS avec Microsoft 365 abonnement.</span><span class="sxs-lookup"><span data-stu-id="0ecda-107">This feature is only supported on Android and iOS with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="0ecda-108">Dans cet article, vous allez apprendre à configurer votre Outlook pour permettre aux utilisateurs d’organiser et de participer à une réunion à l’aide de votre service de réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="0ecda-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="0ecda-109">Tout au long de cet article, nous allons utiliser un fournisseur fictif de services de réunion en ligne, « Contoso ».</span><span class="sxs-lookup"><span data-stu-id="0ecda-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="0ecda-110">Configuration de votre environnement</span><span class="sxs-lookup"><span data-stu-id="0ecda-110">Set up your environment</span></span>

<span data-ttu-id="0ecda-111">[Complétez Outlook démarrage](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rapide qui crée un projet de compl?ment avec le générateur Yeoman pour Office compl?ments.</span><span class="sxs-lookup"><span data-stu-id="0ecda-111">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="0ecda-112">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="0ecda-112">Configure the manifest</span></span>

<span data-ttu-id="0ecda-113">Pour permettre aux utilisateurs de créer des réunions en ligne avec votre application, vous devez configurer le [point d’extension MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) dans le manifeste sous l’élément `MobileFormFactor` parent.</span><span class="sxs-lookup"><span data-stu-id="0ecda-113">To enable users to create online meetings with your add-in, you must configure the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="0ecda-114">Les autres facteurs de forme ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="0ecda-114">Other form factors are not supported.</span></span>

1. <span data-ttu-id="0ecda-115">Dans votre éditeur de code, ouvrez le projet de démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="0ecda-115">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="0ecda-116">Ouvrez **lemanifest.xml** situé à la racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="0ecda-116">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="0ecda-117">Sélectionnez l’intégralité du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le `<VersionOverrides>` par le code XML suivant.</span><span class="sxs-lookup"><span data-stu-id="0ecda-117">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="0ecda-118">Pour en savoir plus sur les manifestes des Outlook, voir les [manifestes](manifests.md) de Outlook et ajouter la prise en charge des commandes de Outlook [Mobile.](add-mobile-support.md)</span><span class="sxs-lookup"><span data-stu-id="0ecda-118">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="0ecda-119">Implémenter l’ajout de détails de réunion en ligne</span><span class="sxs-lookup"><span data-stu-id="0ecda-119">Implement adding online meeting details</span></span>

<span data-ttu-id="0ecda-120">Dans cette section, découvrez comment votre script de add-in peut mettre à jour la réunion d’un utilisateur pour inclure les détails de la réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="0ecda-120">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

1. <span data-ttu-id="0ecda-121">À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="0ecda-121">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="0ecda-122">Remplacez tout le contenu du **fichiercommands.js** par le javaScript suivant.</span><span class="sxs-lookup"><span data-stu-id="0ecda-122">Replace the entire content of the **commands.js** file with the following JavaScript.</span></span>

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

## <a name="testing-and-validation"></a><span data-ttu-id="0ecda-123">Test et validation</span><span class="sxs-lookup"><span data-stu-id="0ecda-123">Testing and validation</span></span>

<span data-ttu-id="0ecda-124">Suivez les instructions habituelles [pour tester et valider votre add-in.](testing-and-tips.md)</span><span class="sxs-lookup"><span data-stu-id="0ecda-124">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="0ecda-125">Après [le chargement de](sideload-outlook-add-ins-for-testing.md) version Outlook sur le web, Windows ou Mac, redémarrez Outlook sur votre appareil mobile Android ou iOS.</span><span class="sxs-lookup"><span data-stu-id="0ecda-125">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android or iOS mobile device.</span></span> <span data-ttu-id="0ecda-126">Ensuite, sur un nouvel écran de réunion, vérifiez que le Microsoft Teams ou Skype bascule est remplacé par le vôtre.</span><span class="sxs-lookup"><span data-stu-id="0ecda-126">Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="0ecda-127">Créer une interface utilisateur de réunion</span><span class="sxs-lookup"><span data-stu-id="0ecda-127">Create meeting UI</span></span>

<span data-ttu-id="0ecda-128">En tant qu’organisateur de réunion, vous devez voir des écrans semblables aux trois images suivantes lorsque vous créez une réunion.</span><span class="sxs-lookup"><span data-stu-id="0ecda-128">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="0ecda-129">[![Écran créer une réunion sur Android - Contoso bascule.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="0ecda-129">[![The create meeting screen on Android - Contoso toggle off.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)</span></span> <span data-ttu-id="0ecda-130">[![Écran Créer une réunion sur Android : chargement du basculement Contoso.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="0ecda-130">[![The create meeting screen on Android - loading Contoso toggle.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)</span></span> <span data-ttu-id="0ecda-131">[![Écran créer une réunion sur Android - Contoso bascule.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="0ecda-131">[![The create meeting screen on Android - Contoso toggle on.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="0ecda-132">Rejoindre l’interface utilisateur de réunion</span><span class="sxs-lookup"><span data-stu-id="0ecda-132">Join meeting UI</span></span>

<span data-ttu-id="0ecda-133">En tant que participant à la réunion, vous devez voir un écran semblable à l’image suivante lorsque vous visualisez la réunion.</span><span class="sxs-lookup"><span data-stu-id="0ecda-133">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="0ecda-134">[![Capture d’écran de l’écran participer à une réunion sur Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="0ecda-134">[![Screenshot of join meeting screen on Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0ecda-135">Si le lien Rejoindre  n’est pas disponible, il se peut que le modèle de réunion en ligne de votre service ne soit pas inscrit sur nos serveurs.</span><span class="sxs-lookup"><span data-stu-id="0ecda-135">If you don't see the **Join** link, it may be that the online-meeting template for your service is not registered on our servers.</span></span> <span data-ttu-id="0ecda-136">Pour plus [d’informations, consultez](#register-your-online-meeting-template) la section Inscrire votre modèle de réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="0ecda-136">See the [Register your online-meeting template](#register-your-online-meeting-template) section for details.</span></span>

## <a name="register-your-online-meeting-template"></a><span data-ttu-id="0ecda-137">Inscrire votre modèle de réunion en ligne</span><span class="sxs-lookup"><span data-stu-id="0ecda-137">Register your online-meeting template</span></span>

<span data-ttu-id="0ecda-138">Si vous souhaitez inscrire le modèle de réunion en ligne pour votre service, vous pouvez créer un GitHub avec les détails.</span><span class="sxs-lookup"><span data-stu-id="0ecda-138">If you would like to register the online-meeting template for your service, you can create a GitHub issue with the details.</span></span> <span data-ttu-id="0ecda-139">Après cela, nous vous contacterons pour coordonner la chronologie d’inscription.</span><span class="sxs-lookup"><span data-stu-id="0ecda-139">After that, we'll contact you to coordinate registration timeline.</span></span>

1. <span data-ttu-id="0ecda-140">Go to the **Feedback** section at the end of this article.</span><span class="sxs-lookup"><span data-stu-id="0ecda-140">Go to the **Feedback** section at the end of this article.</span></span>
1. <span data-ttu-id="0ecda-141">Appuyez sur **le lien Cette page.**</span><span class="sxs-lookup"><span data-stu-id="0ecda-141">Press the **This page** link.</span></span>
1. <span data-ttu-id="0ecda-142">Définissez **le titre** du nouveau problème sur « Enregistrer le modèle de réunion en ligne pour mon service », en remplaçant par votre nom de `my-service` service.</span><span class="sxs-lookup"><span data-stu-id="0ecda-142">Set the **Title** of the new issue to "Register the online-meeting template for my-service", replacing `my-service` with your service name.</span></span>
1. <span data-ttu-id="0ecda-143">Dans le corps du problème, remplacez la chaîne « [Entrez les commentaires ici] » par la chaîne que vous avez définie dans la variable ou une variable similaire de la section Implémenter l’ajout de détails de réunion en ligne plus haut dans cet `newBody` article. [](#implement-adding-online-meeting-details)</span><span class="sxs-lookup"><span data-stu-id="0ecda-143">In the issue body, replace the string "[Enter feedback here]" with the string you set in the `newBody` or similar variable from the [Implement adding online meeting details](#implement-adding-online-meeting-details) section earlier in this article.</span></span>
1. <span data-ttu-id="0ecda-144">Cliquez **sur Envoyer un nouveau problème.**</span><span class="sxs-lookup"><span data-stu-id="0ecda-144">Click **Submit new issue**.</span></span>

![Capture d’écran du nouveau GitHub’écran de problème avec l’exemple de contenu Contoso.](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a><span data-ttu-id="0ecda-146">API disponibles</span><span class="sxs-lookup"><span data-stu-id="0ecda-146">Available APIs</span></span>

<span data-ttu-id="0ecda-147">Les API suivantes sont disponibles pour cette fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="0ecda-147">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="0ecda-148">API d’organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="0ecda-148">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="0ecda-149">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span><span class="sxs-lookup"><span data-stu-id="0ecda-149">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="0ecda-150">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Heure](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="0ecda-150">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="0ecda-151">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="0ecda-151">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="0ecda-152">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="0ecda-152">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="0ecda-153">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="0ecda-153">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="0ecda-154">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="0ecda-154">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="0ecda-155">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Heure](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="0ecda-155">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="0ecda-156">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Objet](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="0ecda-156">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="0ecda-157">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="0ecda-157">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span></span>
- <span data-ttu-id="0ecda-158">Gérer le flux d’th</span><span class="sxs-lookup"><span data-stu-id="0ecda-158">Handle auth flow</span></span>
  - [<span data-ttu-id="0ecda-159">API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="0ecda-159">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="0ecda-160">Restrictions</span><span class="sxs-lookup"><span data-stu-id="0ecda-160">Restrictions</span></span>

<span data-ttu-id="0ecda-161">Plusieurs restrictions s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="0ecda-161">Several restrictions apply.</span></span>

- <span data-ttu-id="0ecda-162">Applicable uniquement aux fournisseurs de services de réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="0ecda-162">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="0ecda-163">Seuls les add-ins installés par l’administrateur apparaissent sur l’écran de composition de la réunion, remplaçant l’option Teams ou Skype par défaut.</span><span class="sxs-lookup"><span data-stu-id="0ecda-163">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="0ecda-164">Les add-ins installés par l’utilisateur ne s’activent pas.</span><span class="sxs-lookup"><span data-stu-id="0ecda-164">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="0ecda-165">L’icône du add-in doit être en échelles de gris à l’aide de code hexas ou de son équivalent `#919191` dans [d’autres formats de couleur.](https://convertingcolors.com/hex-color-919191.html)</span><span class="sxs-lookup"><span data-stu-id="0ecda-165">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="0ecda-166">Une seule commande sans interface utilisateur est prise en charge en mode Organisateur de rendez-vous (composition).</span><span class="sxs-lookup"><span data-stu-id="0ecda-166">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>
- <span data-ttu-id="0ecda-167">Le add-in doit mettre à jour les détails de la réunion dans le formulaire de rendez-vous dans le délai d’une minute.</span><span class="sxs-lookup"><span data-stu-id="0ecda-167">The add-in should update the meeting details in the appointment form within the one-minute timeout period.</span></span> <span data-ttu-id="0ecda-168">Toutefois, tout temps passé dans une boîte de dialogue où le module a été ouvert pour authentification, etc. est exclu du délai d’attente.</span><span class="sxs-lookup"><span data-stu-id="0ecda-168">However, any time spent in a dialog box the add-in opened for authentication, etc. is excluded from the timeout period.</span></span>

## <a name="see-also"></a><span data-ttu-id="0ecda-169">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0ecda-169">See also</span></span>

- [<span data-ttu-id="0ecda-170">Compléments pour Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="0ecda-170">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="0ecda-171">Ajouter la prise en charge des commandes de Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="0ecda-171">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
