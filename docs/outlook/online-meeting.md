---
title: Créer un complément Outlook pour un fournisseur de réunions en ligne
description: Explique comment configurer un complément Outlook pour un fournisseur de services de réunion en ligne.
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c2cdb9f6369fd851a13fe45df132482b0ccdc0e
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767181"
---
# <a name="create-an-outlook-add-in-for-an-online-meeting-provider"></a>Créer un complément Outlook pour un fournisseur de réunions en ligne

La configuration d’une réunion en ligne est une expérience de base pour un utilisateur Outlook, et il est facile de [créer une réunion Teams avec Outlook](/microsoftteams/teams-add-in-for-outlook). Toutefois, la création d’une réunion en ligne dans Outlook avec un service non-Microsoft peut s’avérer fastidieuse. En implémentant cette fonctionnalité, les fournisseurs de services peuvent simplifier l’expérience de création et de participation à des réunions en ligne pour leurs utilisateurs de complément Outlook.

> [!IMPORTANT]
> Cette fonctionnalité est prise en charge dans Outlook sur le web, Windows, Mac, Android et iOS avec un abonnement Microsoft 365.

Dans cet article, vous allez apprendre à configurer votre complément Outlook pour permettre aux utilisateurs d’organiser et de rejoindre une réunion à l’aide de votre service de réunion en ligne. Tout au long de cet article, nous allons utiliser un fournisseur de services de réunion en ligne fictif, « Contoso ».

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Suivez le [guide de démarrage rapide Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le générateur Yeoman pour les compléments Office.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour permettre aux utilisateurs de créer des réunions en ligne avec votre complément, vous devez configurer le manifeste. Le balisage diffère selon deux variables :

- Type de plateforme cible ; mobile ou non mobile.
- Type de manifeste ; Manifeste XML ou [Teams pour compléments Office (préversion).](../develop/json-manifest-overview.md)

Si votre complément utilise un manifeste XML et que le complément n’est pris en charge que dans Outlook sur le web, Windows et Mac, sélectionnez l’onglet **Windows, Mac, web** pour obtenir des conseils. Toutefois, si votre complément sera également pris en charge dans Outlook sur Android et iOS, sélectionnez l’onglet **Mobile** .

Si le complément utilise le manifeste Teams (préversion), sélectionnez l’onglet **Manifeste Teams (préversion pour les développeurs).**

> [!IMPORTANT]
> Les fournisseurs de réunions en ligne ne sont pas encore pris en charge pour le manifeste Teams (préversion). Nous travaillons à fournir ce support bientôt.

# <a name="windows-mac-web"></a>[Windows, Mac, web](#tab/non-mobile)

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide Outlook que vous avez créé.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

1. Sélectionnez le nœud entier **\<VersionOverrides\>** (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant.

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

# <a name="mobile"></a>[Mobile](#tab/mobile)

Pour permettre aux utilisateurs de créer une réunion en ligne à partir de leur appareil mobile, le [point d’extension MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) est configuré dans le manifeste sous l’élément **\<MobileFormFactor\>** parent . Ce point d’extension n’est pas pris en charge dans d’autres facteurs de forme.

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide Outlook que vous avez créé.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

1. Sélectionnez le nœud entier **\<VersionOverrides\>** (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant.

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

# <a name="teams-manifest-developer-preview"></a>[Manifeste Teams (préversion pour les développeurs)](#tab/jsonmanifest)

> [!IMPORTANT]
> Les fournisseurs de réunions en ligne ne sont pas encore pris en charge pour le [manifeste Teams pour les compléments Office (préversion).](../develop/json-manifest-overview.md) Cet onglet est destiné à une utilisation ultérieure.

1. Ouvrez le fichier **manifest.json** .

1. Recherchez le *premier* objet dans le tableau « authorization.permissions.resourceSpecific » et définissez sa propriété « name » sur « MailboxItem.ReadWrite.User ». Cela doit ressembler à ceci lorsque vous avez terminé.

    ```json
    {
        "name": "MailboxItem.ReadWrite.User",
        "type": "Delegated"
    }
    ```

1. Dans le tableau « validDomains », remplacez l’URL par «https://contoso.com », qui est l’URL du fournisseur de réunion en ligne fictif. Le tableau doit ressembler à ceci lorsque vous avez terminé.

    ```json
    "validDomains": [
        "https://contoso.com"
    ],
    ```

1. Ajoutez l’objet suivant au tableau « extensions.runtimes ». Notez ce qui suit à propos de ce code.

   - La valeur « minVersion » de l’ensemble de conditions requises pour la boîte aux lettres est définie sur « 1.3 » afin que le runtime ne soit pas lancé sur les plateformes et les versions d’Office pour lesquelles cette fonctionnalité n’est pas prise en charge.
   - Le « id » du runtime est défini sur le nom descriptif « online_meeting_runtime ».
   - La propriété « code.page » est définie sur l’URL du fichier HTML sans interface utilisateur qui chargera la commande de fonction.
   - La propriété « lifetime » est définie sur « short », ce qui signifie que le runtime démarre lorsque le bouton de commande de fonction est sélectionné et s’arrête à la fin de la fonction. (Dans certains cas rares, le runtime s’arrête avant la fin du gestionnaire. Voir [Runtimes dans les compléments Office](../testing/runtimes.md).)
   - Il existe une action pour exécuter une fonction nommée « insertContosoMeeting ». Vous allez créer cette fonction dans une étape ultérieure.

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.3"
                }
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "id": "online_meeting_runtime",
        "type": "general",
        "code": {
            "page": "https://contoso.com/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "insertContosoMeeting",
                "type": "executeFunction",
                "displayName": "insertContosoMeeting"
            }
        ]
    }
    ```

1. Remplacez le tableau « extensions.ribbons » par ce qui suit. Notez les points suivants concernant ce balisage.

   - La valeur « minVersion » de l’ensemble de conditions requises pour la boîte aux lettres est définie sur « 1.3 » afin que les personnalisations du ruban n’apparaissent pas sur les plateformes et les versions d’Office où cette fonctionnalité n’est pas prise en charge.
   - Le tableau « contexts » spécifie que le ruban est disponible uniquement dans la fenêtre de l’organisateur des détails de la réunion.
   - Il y aura un groupe de contrôle personnalisé sous l’onglet du ruban par défaut (de la fenêtre de l’organisateur des détails de la réunion) nommé **Réunion Contoso**.
   - Le groupe aura un bouton intitulé **Ajouter une réunion Contoso**.
   - « actionId » du bouton a été défini sur « insertContosoMeeting », ce qui correspond à « id » de l’action que vous avez créée à l’étape précédente.

    ```json
    "ribbons": [
      {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.3"
                }
            ],
            "scopes": [
                "mail"
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "contexts": [
            "meetingDetailsOrganizer"
        ],
        "tabs": [
            {
                "builtInTabId": "TabDefault",
                "groups": [
                    {
                        "id": "apptComposeGroup",
                        "label": "Contoso meeting",
                        "controls": [
                            {
                                "id": "insertMeetingButton",
                                "type": "button",
                                "label": "Add a Contoso meeting",
                                "icons": [
                                    {
                                        "size": 16,
                                        "file": "icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "file": "icon-32.png"
                                    },
                                    {
                                        "size": 64,
                                        "file": "icon-64_02.png"
                                    },
                                    {
                                        "size": 80,
                                        "file": "icon-80.png"
                                    }
                                ],
                                "supertip": {
                                    "title": "Add a Contoso meeting",
                                    "description": "Add a Contoso meeting to this appointment."
                                },
                                "actionId": "insertContosoMeeting",
                            }
                        ]
                    }
                ]
            }
        ]
      }
    ]
    ```

---

> [!TIP]
> Pour en savoir plus sur les manifestes pour les compléments Outlook, voir [Manifestes de complément Outlook](manifests.md) et [Ajouter la prise en charge des commandes de complément pour Outlook Mobile](add-mobile-support.md).

## <a name="implement-adding-online-meeting-details"></a>Implémenter l’ajout de détails de réunion en ligne

Dans cette section, découvrez comment votre script de complément peut mettre à jour la réunion d’un utilisateur pour inclure les détails de la réunion en ligne. Les éléments suivants s’appliquent à toutes les plateformes prises en charge.

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.

1. Remplacez tout le contenu du fichier **commands.js** par le code JavaScript suivant.

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

    let mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define and register a function command named `insertContosoMeeting` (referenced in the manifest)
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

Suivez les instructions habituelles pour [tester et valider votre complément](testing-and-tips.md), puis [charger une version test](sideload-outlook-add-ins-for-testing.md) du manifeste dans Outlook sur le web, Windows ou Mac. Si votre complément prend également en charge les appareils mobiles, redémarrez Outlook sur votre appareil Android ou iOS après le chargement indépendant. Une fois le complément chargé, créez une réunion et vérifiez que le bouton bascule Microsoft Teams ou Skype est remplacé par le vôtre.

### <a name="create-meeting-ui"></a>Créer une interface utilisateur de réunion

En tant qu’organisateur de réunion, vous devez voir des écrans similaires aux trois images suivantes lorsque vous créez une réunion.

[![L’écran Créer une réunion sur Android avec le bouton bascule Contoso désactivé.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![L’écran Créer une réunion sur Android avec un bouton bascule Contoso de chargement.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![Écran créer une réunion sur Android avec le bouton bascule Contoso activé.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>Participer à l’interface utilisateur d’une réunion

En tant que participant à la réunion, vous devez voir un écran similaire à l’image suivante lorsque vous affichez la réunion.

[![Écran de participation à la réunion sur Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> Le bouton **Joindre** n’est pris en charge que dans Outlook sur le web, Mac, Android et iOS. Si vous voyez uniquement un lien de réunion, mais que vous ne voyez pas le bouton **Rejoindre** dans un client pris en charge, il se peut que le modèle de réunion en ligne de votre service ne soit pas inscrit sur nos serveurs. Pour plus d’informations, consultez la section [Inscrire votre modèle de réunion en ligne](#register-your-online-meeting-template) .

## <a name="register-your-online-meeting-template"></a>Inscrire votre modèle de réunion en ligne

L’inscription de votre complément de réunion en ligne est facultative. Il s’applique uniquement si vous souhaitez faire apparaître le bouton **Rejoindre** dans les réunions, en plus du lien de réunion. Une fois que vous avez développé votre complément de réunion en ligne et que vous souhaitez l’inscrire, créez un problème GitHub en suivant les conseils suivants. Nous vous contacterons pour coordonner une chronologie d’inscription.

> [!IMPORTANT]
> Le bouton **Joindre** n’est pris en charge que dans Outlook sur le web, Mac, Android et iOS.

1. Créez un [problème GitHub](https://github.com/OfficeDev/office-js/issues/new).
1. Définissez le **titre** du nouveau problème sur « Outlook : Inscrire le modèle de réunion en ligne pour my-service », en remplaçant par le nom de `my-service` votre service.
1. Dans le corps du problème, remplacez le texte existant par la chaîne que vous avez définie dans la `newBody` variable ou une variable similaire de la section [Implémenter l’ajout de détails de réunion en ligne](#implement-adding-online-meeting-details) plus haut dans cet article.
1. Cliquez sur **Envoyer un nouveau problème**.

![Un nouvel écran de problème GitHub avec un exemple de contenu Contoso.](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>API disponibles

Les API suivantes sont disponibles pour cette fonctionnalité.

- API d’organisateur de rendez-vous
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-body-member) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-getasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-setasync-member(1)))
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-end-member) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-location-member) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-start-member) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-subject-member) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))
  - [Office.context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))
- Gérer le flux d’authentification
  - [API de boîte de dialogue](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>Restrictions

Plusieurs restrictions s’appliquent.

- Applicable uniquement aux fournisseurs de services de réunion en ligne.
- Seuls les compléments installés par l’administrateur s’affichent sur l’écran de composition de la réunion, remplaçant l’option Teams ou Skype par défaut. Les compléments installés par l’utilisateur ne s’activent pas.
- L’icône de complément doit être en nuances de gris à l’aide du code `#919191` hexadécimal ou de son équivalent dans [d’autres formats de couleur](https://convertingcolors.com/hex-color-919191.html).
- Une seule commande de fonction est prise en charge en mode Organisateur de rendez-vous (composition).
- Le complément doit mettre à jour les détails de la réunion dans le formulaire de rendez-vous dans le délai d’expiration d’une minute. Toutefois, le temps passé dans une boîte de dialogue au complément ouvert pour l’authentification, par exemple, est exclu du délai d’expiration.

## <a name="see-also"></a>Voir aussi

- [Compléments pour Outlook Mobile](outlook-mobile-addins.md)
- [Ajout de la prise en charge des commandes de complément pour Outlook Mobile](add-mobile-support.md)
