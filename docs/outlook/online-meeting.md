---
title: Créer un complément Outlook Mobile pour un fournisseur de réunions en ligne (aperçu)
description: Explique comment configurer un complément Outlook Mobile pour un fournisseur de services en ligne.
ms.topic: article
ms.date: 04/13/2020
localization_priority: Normal
ms.openlocfilehash: 6a9d484bb74f238c0c62e689c66afaeb284eec2d
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241090"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider-preview"></a><span data-ttu-id="28e17-103">Créer un complément Outlook Mobile pour un fournisseur de réunions en ligne (aperçu)</span><span class="sxs-lookup"><span data-stu-id="28e17-103">Create an Outlook mobile add-in for an online-meeting provider (preview)</span></span>

<span data-ttu-id="28e17-104">La configuration d’une réunion en ligne est une expérience de base pour un utilisateur d’Outlook, et il est facile de [créer une réunion teams avec Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span><span class="sxs-lookup"><span data-stu-id="28e17-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="28e17-105">Toutefois, la création d’une réunion en ligne dans Outlook avec un service non-Microsoft peut être lourde.</span><span class="sxs-lookup"><span data-stu-id="28e17-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="28e17-106">En implémentant cette fonctionnalité, les fournisseurs de services peuvent rationaliser l’expérience de création des réunions en ligne pour leurs utilisateurs des compléments Outlook.</span><span class="sxs-lookup"><span data-stu-id="28e17-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!NOTE]
> <span data-ttu-id="28e17-107">Cette fonctionnalité est uniquement prise en charge en [Aperçu](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur Android avec un abonnement Office 365.</span><span class="sxs-lookup"><span data-stu-id="28e17-107">This feature is only supported in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="28e17-108">Dans cet article, vous apprendrez à configurer votre complément Outlook Mobile pour permettre aux utilisateurs d’organiser et de participer à une réunion à l’aide de votre service de réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="28e17-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="28e17-109">Tout au long de cet article, nous allons utiliser un fournisseur de services de réunion en ligne fictif, « contoso ».</span><span class="sxs-lookup"><span data-stu-id="28e17-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="28e17-110">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="28e17-110">Configure the manifest</span></span>

<span data-ttu-id="28e17-111">Pour permettre aux utilisateurs de créer des réunions en ligne avec votre complément, vous devez configurer `MobileOnlineMeetingCommandSurface` le point d’extension dans le manifeste sous l' `MobileFormFactor`élément parent.</span><span class="sxs-lookup"><span data-stu-id="28e17-111">To enable users to create online meetings with your add-in, you must configure the `MobileOnlineMeetingCommandSurface` extension point in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="28e17-112">Les autres facteurs de forme ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="28e17-112">Other form factors are not supported.</span></span>

<span data-ttu-id="28e17-113">L’exemple suivant montre un exemple du manifeste qui inclut l’élément `MobileFormFactor` et `MobileOnlineMeetingCommandSurface` le point d’extension.</span><span class="sxs-lookup"><span data-stu-id="28e17-113">The following example shows a sample of the manifest that includes the `MobileFormFactor` element and `MobileOnlineMeetingCommandSurface` extension point.</span></span>

```xml
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <MobileFormFactor>
          <FunctionFile resid="residMobileFuncUrl" />
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <!-- Configure selected extension point. -->
            <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
              <Label resid="residUILessButton0Name" />
              <Icon>
                <bt:Image resid="UiLessIcon" size="25" scale="1" />
                <bt:Image resid="UiLessIcon" size="25" scale="2" />
                <bt:Image resid="UiLessIcon" size="25" scale="3" />
                <bt:Image resid="UiLessIcon" size="32" scale="1" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="1" />
                <bt:Image resid="UiLessIcon" size="48" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="3" />
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="28e17-114">Implémenter l’ajout des détails de la réunion en ligne</span><span class="sxs-lookup"><span data-stu-id="28e17-114">Implement adding online meeting details</span></span>

<span data-ttu-id="28e17-115">Dans cette section, Découvrez comment votre script de complément peut mettre à jour la réunion d’un utilisateur pour inclure les détails de la réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="28e17-115">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

<span data-ttu-id="28e17-116">L’exemple suivant montre comment construire les détails de la réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="28e17-116">The following example shows how you construct online meeting details.</span></span> <span data-ttu-id="28e17-117">Non affiché indique comment obtenir l’ID de l’organisateur de la réunion et d’autres détails à partir de votre service.</span><span class="sxs-lookup"><span data-stu-id="28e17-117">Not shown is how to get the meeting organizer's ID and other details from your service.</span></span>

```js
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
```

<span data-ttu-id="28e17-118">L’exemple suivant montre comment définir une fonction sans interface utilisateur nommée `insertContosoMeeting` référencée dans le manifeste pour mettre à jour le corps de la réunion avec les détails de la réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="28e17-118">The following example shows how to define a UI-less function named `insertContosoMeeting` referenced in the manifest to update the meeting body with the online meeting details.</span></span>

```js
var mailboxItem;

// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

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
```

<span data-ttu-id="28e17-119">L’exemple suivant illustre une implémentation de la fonction `updateBody` de prise en charge utilisée dans l’exemple précédent qui ajoute les détails de réunion en ligne au corps de la réunion.</span><span class="sxs-lookup"><span data-stu-id="28e17-119">The following example shows an implementation of the supporting function `updateBody` used in the previous example that appends the online meeting details to the current body of the meeting.</span></span>

```js
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
```

## <a name="testing-and-validation"></a><span data-ttu-id="28e17-120">Test et validation</span><span class="sxs-lookup"><span data-stu-id="28e17-120">Testing and validation</span></span>

<span data-ttu-id="28e17-121">Suivez les instructions habituelles pour [tester et valider votre complément](testing-and-tips.md).</span><span class="sxs-lookup"><span data-stu-id="28e17-121">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="28e17-122">Après avoir [chargement](sideload-outlook-add-ins-for-testing.md) dans Outlook sur le Web, Windows ou Mac, redémarrez Outlook sur votre appareil mobile Android (Android est le seul client pris en charge pour l’instant).</span><span class="sxs-lookup"><span data-stu-id="28e17-122">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device (Android is the only supported client for now).</span></span> <span data-ttu-id="28e17-123">Ensuite, dans un nouvel écran de réunion, vérifiez que le bouton bascule Microsoft teams ou Skype est remplacé par le vôtre.</span><span class="sxs-lookup"><span data-stu-id="28e17-123">Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="28e17-124">Créer une interface utilisateur de réunion</span><span class="sxs-lookup"><span data-stu-id="28e17-124">Create meeting UI</span></span>

<span data-ttu-id="28e17-125">En tant qu’organisateur de la réunion, vous devez voir des écrans semblables aux trois images suivantes lors de la création d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="28e17-125">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="28e17-126">[capture d’écran de la boîte de création de l’écran de réunion sur Android-contoso désactiver la capture d’écran de créer un écran de réunion sur Android-chargement de la capture d’écran de la création de la réunion sur Android-contoso-activer/désactiver ![](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [ ![](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [ ![](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="28e17-126">[![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="28e17-127">Interface utilisateur joindre une réunion</span><span class="sxs-lookup"><span data-stu-id="28e17-127">Join meeting UI</span></span>

<span data-ttu-id="28e17-128">En tant que participant à la réunion, vous devriez voir un écran semblable à l’image suivante lorsque vous affichez la réunion.</span><span class="sxs-lookup"><span data-stu-id="28e17-128">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="28e17-129">[![capture d’écran de l’écran de participation à une réunion sur Android](../images/outlook-android-join-online-meeting-view.png)](../images/outlook-android-join-online-meeting-view-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="28e17-129">[![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view.png)](../images/outlook-android-join-online-meeting-view-expanded.png#lightbox)</span></span>

## <a name="available-apis"></a><span data-ttu-id="28e17-130">API disponibles</span><span class="sxs-lookup"><span data-stu-id="28e17-130">Available APIs</span></span>

<span data-ttu-id="28e17-131">Les API suivantes sont disponibles pour cette fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="28e17-131">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="28e17-132">API d’organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="28e17-132">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="28e17-133">[Office. Context. Mailbox. Item. Subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="28e17-133">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="28e17-134">[Office. Context. Mailbox. Item. Start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="28e17-134">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="28e17-135">[Office. Context. Mailbox. Item. end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([heure](/javascript/api/outlook/office.time?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="28e17-135">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="28e17-136">[Office. Context. Mailbox. Item. Location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([emplacement](/javascript/api/outlook/office.location?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="28e17-136">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="28e17-137">[Office. Context. Mailbox. Item. optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="28e17-137">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="28e17-138">[Office. Context. Mailbox. Item. requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="28e17-138">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="28e17-139">[Office. Context. Mailbox. Item. Body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body. getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body. setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))</span><span class="sxs-lookup"><span data-stu-id="28e17-139">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="28e17-140">[Office. Context. Mailbox. Item. loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="28e17-140">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="28e17-141">[Office. Context. roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="28e17-141">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))</span></span>
- <span data-ttu-id="28e17-142">Gérer le flux d’authentification</span><span class="sxs-lookup"><span data-stu-id="28e17-142">Handle auth flow</span></span>
  - [<span data-ttu-id="28e17-143">API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="28e17-143">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="28e17-144">Restrictions</span><span class="sxs-lookup"><span data-stu-id="28e17-144">Restrictions</span></span>

<span data-ttu-id="28e17-145">Plusieurs restrictions s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="28e17-145">Several restrictions apply.</span></span>

- <span data-ttu-id="28e17-146">Applicable uniquement aux fournisseurs de service de réunion en ligne.</span><span class="sxs-lookup"><span data-stu-id="28e17-146">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="28e17-147">Actuellement en aperçu, cette fonctionnalité ne doit pas être utilisée dans les compléments de production.</span><span class="sxs-lookup"><span data-stu-id="28e17-147">Currently in preview so this feature shouldn't be used in production add-ins.</span></span>
- <span data-ttu-id="28e17-148">À présent, Android est le seul client pris en charge.</span><span class="sxs-lookup"><span data-stu-id="28e17-148">At present, Android is the only supported client.</span></span> <span data-ttu-id="28e17-149">Le support sur iOS sera bientôt disponible.</span><span class="sxs-lookup"><span data-stu-id="28e17-149">Support on iOS is coming soon.</span></span>
- <span data-ttu-id="28e17-150">Seuls les compléments installés par l’administrateur apparaissent sur l’écran de composition de la réunion et remplacent l’option teams ou Skype par défaut.</span><span class="sxs-lookup"><span data-stu-id="28e17-150">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="28e17-151">Les compléments installés par l’utilisateur ne peuvent pas être activés.</span><span class="sxs-lookup"><span data-stu-id="28e17-151">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="28e17-152">L’icône du complément doit être en nuances de gris à l' `#919191` aide de code hexadécimal ou de son équivalent dans d' [autres formats de couleur](https://convertingcolors.com/hex-color-919191.html).</span><span class="sxs-lookup"><span data-stu-id="28e17-152">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="28e17-153">Une seule commande sans interface utilisateur est prise en charge dans le mode organisateur de rendez-vous (composition).</span><span class="sxs-lookup"><span data-stu-id="28e17-153">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>

## <a name="see-also"></a><span data-ttu-id="28e17-154">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="28e17-154">See also</span></span>

- [<span data-ttu-id="28e17-155">Compléments pour Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="28e17-155">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="28e17-156">Ajouter la prise en charge des commandes de complément pour Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="28e17-156">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
