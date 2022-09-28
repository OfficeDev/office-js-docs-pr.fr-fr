---
title: Configurer votre complément Outlook pour l’activation basée sur les événements
description: Découvrez comment configurer votre complément Outlook pour l’activation basée sur les événements.
ms.topic: article
ms.date: 09/21/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0e38f7e9c9d9f06ec7f427b12c04b30d6abf0112
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68093007"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>Configurer votre complément Outlook pour l’activation basée sur les événements

Sans la fonctionnalité d’activation basée sur les événements, un utilisateur doit lancer explicitement un complément pour effectuer ses tâches. Cette fonctionnalité permet à votre complément d’exécuter des tâches basées sur certains événements, en particulier pour les opérations qui s’appliquent à chaque élément. Vous pouvez également l’intégrer au volet Office et aux commandes de fonction.

À la fin de cette procédure pas à pas, vous disposerez d’un complément qui s’exécute chaque fois qu’un nouvel élément est créé et définit l’objet.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans [l’ensemble de conditions requises 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), avec des événements supplémentaires désormais disponibles dans les ensembles de conditions requises suivants. Pour plus d’informations sur l’ensemble minimal de conditions requises d’un événement et sur les clients et plateformes qui le prennent en charge, consultez [événements pris](#supported-events) en [charge et ensembles de conditions requises pris en charge par les serveurs Exchange et les clients Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).
>
> L’activation basée sur les événements n’est pas prise en charge dans Outlook sur iOS ou Android.

## <a name="supported-events"></a>Événements pris en charge

Le tableau suivant répertorie les événements actuellement disponibles et les clients pris en charge pour chaque événement. Lorsqu’un événement est déclenché, le gestionnaire reçoit un `event` objet qui peut inclure des détails spécifiques au type d’événement. La colonne **Description** inclut un lien vers l’objet associé, le cas échéant.

|Événement|Description|Ensemble minimal de conditions requises et clients pris en charge|
|---|---|---|
|`OnNewMessageCompose`|Lors de la rédaction d’un nouveau message (y compris la réponse, répondre à tous et transférer), mais pas lors de la modification, par exemple, d’un brouillon.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac |
|`OnNewAppointmentOrganizer`|Lors de la création d’un rendez-vous, mais pas lors de la modification d’un rendez-vous existant.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac |
|`OnMessageAttachmentsChanged`|Lors de l’ajout ou de la suppression de pièces jointes lors de la composition d’un message.<br><br>Objet de données spécifique à l’événement : [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnAppointmentAttachmentsChanged`|Lors de l’ajout ou de la suppression de pièces jointes lors de la composition d’un rendez-vous.<br><br>Objet de données spécifique à l’événement : [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnMessageRecipientsChanged`|Lors de l’ajout ou de la suppression de destinataires lors de la composition d’un message.<br><br>Objet de données spécifique à l’événement : [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnAppointmentAttendeesChanged`|Lors de l’ajout ou de la suppression de participants lors de la composition d’un rendez-vous.<br><br>Objet de données spécifique à l’événement : [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnAppointmentTimeChanged`|Lors de la modification de la date/heure lors de la composition d’un rendez-vous.<br><br>Objet de données spécifique à l’événement : [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnAppointmentRecurrenceChanged`|Lors de l’ajout, de la modification ou de la suppression des détails de périodicité lors de la rédaction d’un rendez-vous. Si la date/heure est modifiée, l’événement `OnAppointmentTimeChanged` est également déclenché.<br><br>Objet de données spécifique à l’événement : [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnInfoBarDismissClicked`|Lors du rejet d’une notification lors de la composition d’un message ou d’un élément de rendez-vous. Seul le complément qui a ajouté la notification est averti.<br><br>Objet de données spécifique à l’événement : [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnMessageSend`|Lors de l’envoi d’un élément de message. Pour en savoir plus, consultez la [procédure pas à pas des alertes intelligentes](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Navigateur web|
|`OnAppointmentSend`|Lors de l’envoi d’un élément de rendez-vous. Pour en savoir plus, consultez la [procédure pas à pas des alertes intelligentes](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Navigateur web|
|`OnMessageCompose`|Lors de la rédaction d’un nouveau message (y compris la réponse, la réponse à tous et l’envoi) ou la modification d’un brouillon.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Navigateur web|
|`OnAppointmentOrganizer`|Lors de la création d’un rendez-vous ou de la modification d’un rendez-vous existant.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Navigateur web|

> [!NOTE]
> <sup>1 Les</sup> compléments basés sur les événements dans Outlook sur Windows nécessitent un minimum de Windows 10 version 1903 (build 18362) ou Windows Server 2019 version 1903 pour s’exécuter.

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Terminez le [démarrage rapide d’Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md).

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour activer l’activation basée sur les événements de votre complément, vous devez configurer l’élément [Runtimes](/javascript/api/manifest/runtimes) et le point [d’extension LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) dans le `VersionOverridesV1_1` nœud du manifeste. Pour l’instant, `DesktopFormFactor` est le seul facteur de forme pris en charge.

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

1. Sélectionnez l’intégralité **\<VersionOverrides\>** du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant, puis enregistrez vos modifications.

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.10">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and Outlook on the new Mac UI. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook on Windows. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="GroupLabel" />
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
                <Control xsi:type="Button" id="ActionButton">
                  <Label resid="ActionButton.Label"/>
                  <Supertip>
                    <Title resid="ActionButton.Label"/>
                    <Description resid="ActionButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>action</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>

          <!-- Can configure other command surface extension points for add-in command support. -->

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onNewAppointmentComposeHandler"/>
              
              <!-- Other available events -->
              <!--
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnMessageCompose" FunctionName="onMessageComposeHandler" />
              <LaunchEvent Type="OnAppointmentOrganizer" FunctionName="onAppointmentOrganizerHandler" />
              -->
            </LaunchEvents>
            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
            <SourceLocation resid="WebViewRuntime.Url"/>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
        <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
        <!-- Entry needed for Outlook on Windows. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/launchevent.js" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Contoso Add-in"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
        <bt:String id="ActionButton.Label" DefaultValue="Perform an action"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane displaying all available properties."/>
        <bt:String id="ActionButton.Tooltip" DefaultValue="Perform an action when clicked."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web et sur la nouvelle interface utilisateur Mac utilisent un fichier HTML qui peut référencer le même fichier JavaScript. Vous devez fournir des références à ces deux fichiers dans le `Resources` nœud du manifeste, car la plateforme Outlook détermine en fin de compte s’il faut utiliser html ou JavaScript en fonction du client Outlook. Par conséquent, pour configurer la gestion des événements, indiquez l’emplacement du code HTML dans l’élément **\<Runtime\>** , puis, dans son `Override` élément enfant, indiquez l’emplacement du fichier JavaScript incorporé ou référencé par le code HTML.

> [!TIP]
>
> - Pour en savoir plus sur les runtimes dans les compléments, consultez [Runtimes dans les compléments Office](../testing/runtimes.md).
> - Pour en savoir plus sur les manifestes pour les compléments Outlook, consultez [les manifestes de complément Outlook](manifests.md).

## <a name="implement-event-handling"></a>Implémenter la gestion des événements

Vous devez implémenter la gestion de vos événements sélectionnés.

Dans ce scénario, vous allez ajouter la gestion de la composition de nouveaux éléments.

1. À partir du même projet de démarrage rapide, créez un dossier nommé **launchevent** sous le répertoire **./src** .

1. Dans le dossier **./src/launchevent** , créez un fichier nommé **launchevent.js**.

1. Ouvrez le fichier **./src/launchevent/launchevent.js** dans votre éditeur de code et ajoutez le code JavaScript suivant.

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onNewMessageComposeHandler(event) {
      setSubject(event);
    }
    function onNewAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext": event
        },
        function (asyncResult) {
          // Handle success or error.
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
          }

          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
        });
    }

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
    ```

1. Enregistrez vos modifications.

> [!IMPORTANT]
> Windows : À l’heure actuelle, les importations ne sont pas prises en charge dans le fichier JavaScript où vous implémentez la gestion de l’activation basée sur les événements.

## <a name="update-the-commands-html-file"></a>Mettre à jour le fichier HTML des commandes

1. Dans le dossier **./src/commands** , ouvrez **commands.html**.

1. Immédiatement avant la balise **d’en-tête** fermante (`</head>`), ajoutez une entrée de script pour inclure le code JavaScript de gestion des événements.

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. Enregistrez vos modifications.

## <a name="update-webpack-config-settings"></a>Mettre à jour les paramètres de configuration webapck

1. Ouvrez le fichier **webpack.config.js** trouvé dans le répertoire racine du projet et effectuez les étapes suivantes.

1. Recherchez le `plugins` tableau dans l’objet `config` et ajoutez ce nouvel objet au début du tableau.

    ```js
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "./src/launchevent/launchevent.js",
          to: "launchevent.js",
        },
      ],
    }),
    ```

1. Enregistrez vos modifications.

## <a name="try-it-out"></a>Essayez

1. Exécutez les commandes suivantes dans le répertoire racine de votre projet. Lorsque vous exécutez `npm start`, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution) et votre complément est chargé de manière indépendante.

    ```command&nbsp;line
    npm run build
    ```

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Si votre complément n’a pas été chargé automatiquement, suivez les instructions [fournies dans Chargement indépendant des compléments Outlook à des fins de test pour](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) charger manuellement le complément dans Outlook.

1. Dans Outlook sur le web, créez un message.

    ![Fenêtre de message dans Outlook sur le web avec l’objet défini sur composer.](../images/outlook-web-autolaunch-1.png)

1. Dans Outlook sur la nouvelle interface utilisateur Mac, créez un message.

    ![Fenêtre de message dans Outlook sur la nouvelle interface utilisateur Mac avec l’objet défini sur composer.](../images/outlook-mac-autolaunch.png)

1. Dans Outlook sur Windows, créez un message.

    ![Fenêtre de message dans Outlook sur Windows avec l’objet défini sur composer.](../images/outlook-win-autolaunch.png)

## <a name="debug"></a>Débogage

Lorsque vous apportez des modifications à la gestion des événements de lancement dans votre complément, vous devez savoir que :

- Si vous avez mis à jour le manifeste, [supprimez le complément, puis chargez-le](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) à nouveau. Si vous utilisez Outlook sur Windows, fermez et rouvrez-le.
- Si vous avez apporté des modifications à des fichiers autres que le manifeste, fermez et rouvrez Outlook sur Windows, ou actualisez l’onglet du navigateur en cours d’exécution Outlook sur le web.

Lors de l’implémentation de vos propres fonctionnalités, vous devrez peut-être déboguer votre code. Pour obtenir des conseils sur le débogage de l’activation de compléments basés sur des événements, consultez [Déboguer votre complément Outlook basé sur les événements](debug-autolaunch.md).

La journalisation d’exécution est également disponible pour cette fonctionnalité sur Windows. Pour plus d’informations, consultez [Déboguer votre complément avec la journalisation du runtime](../testing/runtime-logging.md#runtime-logging-on-windows).

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="deploy-to-users"></a>Déployer sur les utilisateurs

Vous pouvez déployer des compléments basés sur des événements en chargeant le manifeste via le Centre d'administration Microsoft 365. Dans le portail d’administration, développez la section **Paramètres** dans le volet de navigation, puis sélectionnez **Applications intégrées**. Dans la page **Applications intégrées** , choisissez l’action **Charger des applications personnalisées** .

![Page Applications intégrées sur le Centre d'administration Microsoft 365 avec l’action Charger des applications personnalisées mise en évidence.](../images/outlook-deploy-event-based-add-ins.png)

> [!IMPORTANT]
> Les compléments basés sur des événements sont limités aux déploiements gérés par l’administrateur uniquement. Les utilisateurs ne peuvent pas activer des compléments basés sur des événements à partir d’AppSource ou d’Office Store dans l’application. Pour plus d’informations, consultez les [options de liste AppSource pour votre complément Outlook basé sur les événements](autolaunch-store-options.md).

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

## <a name="event-based-activation-behavior-and-limitations"></a>Comportement et limitations de l’activation basée sur les événements

Les gestionnaires d’événements de lancement de complément sont censés être de courte durée, légers et aussi peuvasifs que possible. Après l’activation, votre complément expire dans un délai d’environ 300 secondes, soit la durée maximale autorisée pour l’exécution de compléments basés sur des événements. Pour signaler que votre complément a terminé le traitement d’un événement de lancement, votre gestionnaire d’événements associé doit appeler la `event.completed` méthode. (Notez que le code inclus après l’exécution de l’instruction `event.completed` n’est pas garanti.) Chaque fois qu’un événement que vos handles de complément déclenchent, le complément est réactivé et exécute le gestionnaire d’événements associé, et la fenêtre de délai d’expiration est réinitialisée. Le complément se termine après son expiration, ou l’utilisateur ferme la fenêtre de composition ou envoie l’élément.

Si l’utilisateur a plusieurs compléments qui se sont abonnés au même événement, la plateforme Outlook lance les compléments dans un ordre particulier. Actuellement, seuls cinq compléments basés sur des événements peuvent être en cours d’exécution active.

Dans tous les clients Outlook pris en charge, l’utilisateur doit rester sur l’élément de messagerie actuel où le complément a été activé pour qu’il s’exécute. La navigation loin de l’élément actif (par exemple, le passage à une autre fenêtre de composition ou un autre onglet) met fin à l’opération de complément. Le complément cesse également de fonctionner lorsque l’utilisateur envoie le message ou le rendez-vous qu’il compose.

Les importations ne sont pas prises en charge dans le fichier JavaScript où vous implémentez la gestion de l’activation basée sur les événements dans le client Windows.

Certaines API Office.js qui modifient ou modifient l’interface utilisateur ne sont pas autorisées à partir de compléments basés sur des événements. Voici les API bloquées.

- Sous `Office.context.auth`:
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > [OfficeRuntime.auth](/javascript/api/office-runtime/officeruntime.auth) est pris en charge dans toutes les versions d’Outlook qui prennent en charge l’activation basée sur les événements et l’authentification unique (SSO), tandis [qu’Office.auth](/javascript/api/office/office.auth) est uniquement pris en charge dans certaines builds Outlook. Pour plus d’informations, consultez Activer l’authentification [unique (SSO) dans les compléments Outlook qui utilisent l’activation basée sur les événements](use-sso-in-event-based-activation.md).
- Sous `Office.context.mailbox`:
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Sous `Office.context.mailbox.item`:
  - `close`
- Sous `Office.context.ui`:
  - `displayDialogAsync`
  - `messageParent`

### <a name="requesting-external-data"></a>Demande de données externes

Vous pouvez demander des données externes à l’aide d’une API telle que [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) ou de [XMLHttpRequest (XHR),](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) une API web standard qui émet des requêtes HTTP pour interagir avec les serveurs.

N’oubliez pas que vous devez utiliser des mesures de sécurité supplémentaires lors de l’utilisation d’objets XMLHttpRequest, nécessitant une [même stratégie d’origine](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) et un [CORS simple (partage de ressources cross-origin)](https://developer.mozilla.org/docs/Web/HTTP/CORS).

Implémentation [CORS simple](https://developer.mozilla.org/docs/Web/HTTP/CORS#simple_requests) :

- Impossible d’utiliser des cookies.
- Prend uniquement en charge les méthodes simples, telles que `GET`, `HEAD`et `POST`.
- Accepte les en-têtes simples avec des noms de `Accept`champs, `Accept-Language`ou `Content-Language`.
- Peut utiliser le `Content-Type`, à condition que le type de contenu soit `application/x-www-form-urlencoded`, `text/plain`ou `multipart/form-data`.
- Impossible d’inscrire des écouteurs d’événements sur l’objet retourné par `XMLHttpRequest.upload`.
- Impossible d’utiliser des `ReadableStream` objets dans les requêtes.

> [!NOTE]
> La prise en charge complète de CORS est disponible dans Outlook sur le web, Mac et Windows (à compter de la version 2201, build 16.0.14813.10000).

## <a name="see-also"></a>Voir aussi

- [Manifestes de complément Outlook](manifests.md)
- [Comment déboguer des compléments basés sur des événements](debug-autolaunch.md)
- [Options de liste AppSource pour votre complément Outlook basé sur les événements](autolaunch-store-options.md)
- [Procédure pas à pas sur les alertes intelligentes et OnMessageSend](smart-alerts-onmessagesend-walkthrough.md)
- Exemples de code de compléments Office :
  - [Utiliser l’activation basée sur les événements Outlook pour chiffrer les pièces jointes, traiter les participants aux demandes de réunion et réagir aux modifications apportées à la date/l’heure du rendez-vous](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Utiliser l’activation Outlook basée sur un événement pour définir la signature](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [Utiliser l’activation basée sur les événements Outlook pour marquer les destinataires externes](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
  - [Utiliser Alertes intelligentes d’Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
