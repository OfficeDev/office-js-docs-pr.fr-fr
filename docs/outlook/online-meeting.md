---
title: Créer un complément mobile Outlook pour un fournisseur de réunions en ligne
description: Explique comment configurer un complément mobile Outlook pour un fournisseur de services de réunion en ligne.
ms.topic: article
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: ada539c323ce6fdf578c26c221eb6090280b5a0a
ms.sourcegitcommit: 5773c76912cdb6f0c07a932ccf07fc97939f6aa1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2022
ms.locfileid: "65244862"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a>Créer un complément mobile Outlook pour un fournisseur de réunions en ligne

La configuration d’une réunion en ligne est une expérience essentielle pour un utilisateur Outlook, et il est facile de [créer une réunion Teams avec Outlook](/microsoftteams/teams-add-in-for-outlook) mobile. Toutefois, la création d’une réunion en ligne dans Outlook avec un service non Microsoft peut être fastidieuse. En implémentant cette fonctionnalité, les fournisseurs de services peuvent simplifier l’expérience de création de réunions en ligne pour leurs utilisateurs de complément Outlook.

> [!IMPORTANT]
> Cette fonctionnalité est uniquement prise en charge sur Android et iOS avec un abonnement Microsoft 365.

Dans cet article, vous allez apprendre à configurer votre complément mobile Outlook pour permettre aux utilisateurs d’organiser et de participer à une réunion à l’aide de votre service de réunion en ligne. Tout au long de cet article, nous allons utiliser un fournisseur de services de réunion en ligne fictif, « Contoso ».

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Terminez le [Outlook démarrage rapide](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le générateur Yeoman pour Office compléments.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour permettre aux utilisateurs de créer des réunions en ligne avec votre complément, vous devez configurer le [point d’extension MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) dans le manifeste sous l’élément `MobileFormFactor`parent. D’autres facteurs de forme ne sont pas pris en charge.

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

1. Sélectionnez l’intégralité `<VersionOverrides>` du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant.

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
> Pour en savoir plus sur les manifestes pour Outlook compléments, consultez [Outlook manifestes de complément](manifests.md) et [ajouter la prise en charge des commandes de complément pour Outlook Mobile](add-mobile-support.md).

## <a name="implement-adding-online-meeting-details"></a>Implémenter l’ajout de détails de réunion en ligne

Dans cette section, découvrez comment votre script de complément peut mettre à jour la réunion d’un utilisateur pour inclure les détails de la réunion en ligne.

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.

1. Remplacez l’intégralité du contenu du fichier **commands.js** par le code JavaScript suivant.

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

    // 2. How to define and register a UI-less function named `insertContosoMeeting` (referenced in the manifest)
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
    // Register the function.
    Office.actions.associate("insertContosoMeeting", insertContosoMeeting);

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
    ```

## <a name="testing-and-validation"></a>Test et validation

Suivez les instructions habituelles pour [tester et valider votre complément](testing-and-tips.md). Après [le chargement indépendant](sideload-outlook-add-ins-for-testing.md) dans Outlook sur le web, Windows ou Mac, redémarrez Outlook sur votre appareil mobile Android ou iOS. Ensuite, sur un nouvel écran de réunion, vérifiez que le bouton bascule Microsoft Teams ou Skype est remplacé par le vôtre.

### <a name="create-meeting-ui"></a>Créer une interface utilisateur de réunion

En tant qu’organisateur de réunion, vous devez voir des écrans similaires aux trois images suivantes lorsque vous créez une réunion.

[![Écran de création de réunion sur Android - Contoso désactivé.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![Écran de création d’une réunion sur Android - chargement du bouton bascule Contoso.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![Écran de création de réunion sur Android - Contoso activé.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>Participer à l’interface utilisateur de la réunion

En tant que participant à la réunion, vous devez voir un écran similaire à l’image suivante lorsque vous affichez la réunion.

[![Capture d’écran de l’écran de participation à la réunion sur Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> Si vous ne voyez pas le lien **Joindre** , il se peut que le modèle de réunion en ligne de votre service ne soit pas inscrit sur nos serveurs. Pour plus d’informations, consultez la section [Inscrire votre modèle de réunion en ligne](#register-your-online-meeting-template) .

## <a name="register-your-online-meeting-template"></a>Inscrire votre modèle de réunion en ligne

Si vous souhaitez inscrire le modèle de réunion en ligne pour votre service, vous pouvez créer un problème de GitHub avec les détails. Après cela, nous vous contacterons pour coordonner la chronologie de l’inscription.

1. Accédez à la section **Commentaires** à la fin de cet article.
1. Appuyez sur le lien **de cette page** .
1. Définissez le **titre** du nouveau problème sur « Inscrire le modèle de réunion en ligne pour mon service », en remplacement de `my-service` votre nom de service.
1. Dans le corps du problème, remplacez la chaîne « [Entrez les commentaires ici] » par la chaîne que vous avez définie dans la `newBody` variable ou une variable similaire de la section [Implémenter l’ajout de détails de réunion en ligne](#implement-adding-online-meeting-details) plus haut dans cet article.
1. Cliquez sur **Envoyer un nouveau problème**.

![capture d’écran du nouvel écran de problème GitHub avec l’exemple de contenu Contoso.](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>API disponibles

Les API suivantes sont disponibles pour cette fonctionnalité.

- API d’organisateur de rendez-vous
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-body-member) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-getasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-setasync-member(1)))
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-end-member) ([heure](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-location-member) ([Emplacement](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) ([Destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) ([Destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-start-member) ([heure](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-subject-member) ([Objet](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))
  - [Office.context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))
- Gérer le flux d’authentification
  - [API de boîte de dialogue](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>Restrictions

Plusieurs restrictions s’appliquent.

- Applicable uniquement aux fournisseurs de services de réunion en ligne.
- Seuls les compléments installés par l’administrateur s’affichent sur l’écran de composition de la réunion, en remplaçant l’option par défaut Teams ou Skype. Les compléments installés par l’utilisateur ne sont pas activés.
- L’icône de complément doit être en nuances de gris à l’aide de code `#919191` hexadécimal ou de son équivalent dans [d’autres formats de couleur](https://convertingcolors.com/hex-color-919191.html).
- Une seule commande sans interface utilisateur est prise en charge en mode Organisateur de rendez-vous (composition).
- Le complément doit mettre à jour les détails de la réunion dans le formulaire de rendez-vous dans le délai d’expiration d’une minute. Toutefois, tout temps passé dans une boîte de dialogue que le complément a ouvert pour l’authentification, etc., est exclu du délai d’expiration.

## <a name="see-also"></a>Voir aussi

- [Compléments pour Outlook Mobile](outlook-mobile-addins.md)
- [Ajouter la prise en charge des commandes de complément pour Outlook Mobile](add-mobile-support.md)
