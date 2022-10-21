---
title: Configurer votre complément Outlook pour l’activation basée sur les événements
description: Découvrez comment configurer votre complément Outlook pour l’activation basée sur les événements.
ms.topic: article
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: ce2821ed5d226ff2c6a2b3c718d5711689523ac6
ms.sourcegitcommit: d402c37fc3388bd38761fedf203a7d10fce4e899
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/21/2022
ms.locfileid: "68664678"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>Configurer votre complément Outlook pour l’activation basée sur les événements

Sans la fonctionnalité d’activation basée sur les événements, un utilisateur doit lancer explicitement un complément pour effectuer ses tâches. Cette fonctionnalité permet à votre complément d’exécuter des tâches basées sur certains événements, en particulier pour les opérations qui s’appliquent à chaque élément. Vous pouvez également intégrer le volet Office et les commandes de fonction.

À la fin de cette procédure pas à pas, vous disposez d’un complément qui s’exécute chaque fois qu’un nouvel élément est créé et définit l’objet.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans [l’ensemble de conditions requises 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), avec des événements supplémentaires désormais disponibles dans les ensembles de conditions requises suivants. Pour plus d’informations sur l’ensemble de conditions requises minimales d’un événement et les clients et plateformes qui le prennent en charge, voir [Événements pris en charge](#supported-events) et [Ensembles de conditions requises pris en charge par les serveurs Exchange et les clients Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).
>
> L’activation basée sur les événements n’est pas prise en charge dans Outlook sur iOS ou Android.

## <a name="supported-events"></a>Événements pris en charge

Le tableau suivant répertorie les événements actuellement disponibles et les clients pris en charge pour chaque événement. Lorsqu’un événement est déclenché, le gestionnaire reçoit un `event` objet qui peut inclure des détails spécifiques au type d’événement. La colonne **Description** inclut un lien vers l’objet associé, le cas échéant.

|Nom canonique de l’événement</br>et nom du manifeste XML|Nom du manifeste Teams|Description|Ensemble de conditions requises minimales et clients pris en charge|
|---|---|---|---|
|`OnNewMessageCompose`| newMessageComposeCreated |Lors de la composition d’un nouveau message (y compris la réponse, la réponse et le transfert), mais pas lors de la modification, par exemple, un brouillon.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac |
|`OnNewAppointmentOrganizer`|newAppointmentOrganizerCreated|Lors de la création d’un rendez-vous, mais pas lors de la modification d’un rendez-vous existant.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac |
|`OnMessageAttachmentsChanged`|messageAttachmentsChanged|Lors de l’ajout ou de la suppression de pièces jointes lors de la composition d’un message.<br><br>Objet de données spécifique à l’événement : [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnAppointmentAttachmentsChanged`|appointmentAttachmentsChanged|Lors de l’ajout ou de la suppression de pièces jointes lors de la composition d’un rendez-vous.<br><br>Objet de données spécifique à l’événement : [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnMessageRecipientsChanged`|messageRecipientsChanged|Lors de l’ajout ou de la suppression de destinataires lors de la composition d’un message.<br><br>Objet de données spécifique à l’événement : [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnAppointmentAttendeesChanged`|appointmentAttendeesChanged|Lors de l’ajout ou de la suppression de participants lors de la composition d’un rendez-vous.<br><br>Objet de données spécifique à l’événement : [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnAppointmentTimeChanged`|appointmentTimeChanged|Lors de la modification de date/heure lors de la composition d’un rendez-vous.<br><br>Objet de données spécifique à l’événement : [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnAppointmentRecurrenceChanged`|appointmentRecurrenceChanged|Lors de l’ajout, de la modification ou de la suppression des détails de périodicité lors de la composition d’un rendez-vous. Si la date/heure est modifiée, l’événement `OnAppointmentTimeChanged` est également déclenché.<br><br>Objet de données spécifique à l’événement : [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnInfoBarDismissClicked`|infoBarDismissClicked|Lors de la suppression d’une notification lors de la composition d’un message ou d’un élément de rendez-vous. Seul le complément qui a ajouté la notification sera notifié.<br><br>Objet de données spécifique à l’événement : [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Navigateur web<br>- Nouvelle interface utilisateur Mac|
|`OnMessageSend`|messageSending|Lors de l’envoi d’un élément de message. Pour plus d’informations, consultez la [procédure pas à pas des alertes intelligentes](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Navigateur web|
|`OnAppointmentSend`|appointmentSending|Lors de l’envoi d’un élément de rendez-vous. Pour plus d’informations, consultez la [procédure pas à pas des alertes intelligentes](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Navigateur web|
|`OnMessageCompose`|messageComposeOpened|Lors de la composition d’un nouveau message (y compris la réponse, la réponse et le transfert) ou la modification d’un brouillon.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Navigateur web|
|`OnAppointmentOrganizer`|appointmentOrganizerOpened|Lors de la création d’un rendez-vous ou de la modification d’un rendez-vous existant.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Navigateur web|

> [!NOTE]
> <sup>1</sup> Les compléments basés sur les événements dans Outlook sur Windows nécessitent au moins Windows 10 version 1903 (build 18362) ou Windows Server 2019 version 1903 pour s’exécuter.

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Suivez le [guide de démarrage rapide Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md).

> [!NOTE]
> Si vous souhaitez utiliser le [manifeste Teams pour les compléments Office (préversion),](../develop/json-manifest-overview.md) effectuez l’autre démarrage rapide dans [Outlook avec un manifeste Teams (préversion),](../quickstarts/outlook-quickstart-json-manifest.md) mais ignorez toutes les sections après la section **Essayer** .

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour configurer le manifeste, sélectionnez l’onglet correspondant au type de manifeste que vous utilisez.

# <a name="xml-manifest"></a>[Manifeste XML](#tab/xmlmanifest)

Pour activer l’activation basée sur les événements de votre complément, vous devez configurer l’élément [Runtimes](/javascript/api/manifest/runtimes) et le point [d’extension LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) dans le `VersionOverridesV1_1` nœud du manifeste. Pour l’instant, `DesktopFormFactor` est le seul facteur de forme pris en charge.

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

1. Sélectionnez le nœud entier **\<VersionOverrides\>** (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant, puis enregistrez vos modifications.

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

Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web et sur la nouvelle interface utilisateur Mac utilisent un fichier HTML qui peut référencer le même fichier JavaScript. Vous devez fournir des références à ces deux fichiers dans le `Resources` nœud du manifeste, car la plateforme Outlook détermine en fin de compte s’il faut utiliser HTML ou JavaScript en fonction du client Outlook. Par conséquent, pour configurer la gestion des événements, indiquez l’emplacement du code HTML dans l’élément **\<Runtime\>** , puis, dans son `Override` élément enfant, indiquez l’emplacement du fichier JavaScript inclus ou référencé par le code HTML.

# <a name="teams-manifest-developer-preview"></a>[Manifeste Teams (préversion pour les développeurs)](#tab/jsonmanifest)

1. Ouvrez le fichier **manifest.json** .

1. Ajoutez l’objet suivant au tableau « extensions.runtimes ». Notez les points suivants concernant ce balisage :

   - La valeur « minVersion » de l’ensemble de conditions requises pour la boîte aux lettres est définie sur « 1.10 », car le tableau plus haut dans cet article spécifie qu’il s’agit de la version la plus basse de l’ensemble de conditions requises qui prend en charge les `OnNewMessageCompose` événements et `OnNewAppointmentCompose` .
   - Le « id » du runtime est défini sur le nom descriptif « autorun_runtime ».
   - La propriété « code » a une propriété « page » enfant qui est définie sur un fichier HTML et une propriété « script » enfant qui est définie sur un fichier JavaScript. Vous allez créer ou modifier ces fichiers dans les étapes ultérieures. Office utilise l’une de ces valeurs en fonction de la plateforme.
       - Office sur Windows exécute les gestionnaires d’événements dans un runtime JavaScript uniquement, qui charge directement un fichier JavaScript.
       - Office sur Mac et le web exécutent les gestionnaires dans un runtime de navigateur, qui charge un fichier HTML. Ce fichier, à son tour, contient une `<script>` balise qui charge le fichier JavaScript.
     Pour plus d’informations, voir [Runtimes dans les compléments Office](../testing/runtimes.md).
   - La propriété « lifetime » est définie sur « short », ce qui signifie que le runtime démarre quand l’un des événements est déclenché et s’arrête lorsque le gestionnaire se termine. (Dans certains cas rares, le runtime s’arrête avant la fin du gestionnaire. Voir [Runtimes dans les compléments Office](../testing/runtimes.md).)
   - Il existe deux types d'« actions » qui peuvent s’exécuter dans le runtime. Vous allez créer des fonctions pour correspondre à ces actions dans une étape ultérieure.

    ```json
     {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.10"
                }
            ]
        },
        "id": "autorun_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html",
            "script": "https://localhost:3000/launchevent.js"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "onNewMessageComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewMessageComposeHandler"
            },
            {
                "id": "onNewAppointmentComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewAppointmentComposeHandler"
            }
        ]
    }
    ```

1. Ajoutez le tableau « autoRunEvents » suivant en tant que propriété de l’objet dans le tableau « extensions ».

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. Ajoutez l’objet suivant au tableau « autoRunEvents ». La propriété « events » mappe les gestionnaires aux événements, comme décrit dans le tableau ci-dessus dans cet article. Les noms de gestionnaires doivent correspondre à ceux utilisés dans les propriétés « id » des objets du tableau « actions » dans une étape précédente.

    ```json
      {
          "requirements": {
              "capabilities": [
                  {
                      "name": "Mailbox",
                      "minVersion": "1.10"
                  }
              ],
              "scopes": [
                  "mail"
              ]
          },
          "events": [
              {
                  "type": "newMessageComposeCreated",
                  "actionId": "onNewMessageComposeHandler"
              },
              {
                  "type": "newAppointmentOrganizerCreated",
                  "actionId": "onNewAppointmentComposeHandler"
              }
          ]
      }
    ```

---

> [!TIP]
>
> - Pour en savoir plus sur les runtimes dans les compléments, voir [Runtimes dans les compléments Office](../testing/runtimes.md).
> - Pour en savoir plus sur les manifestes pour les compléments Outlook, voir [Manifestes de complément Outlook](manifests.md).

## <a name="implement-event-handling"></a>Implémenter la gestion des événements

Vous devez implémenter la gestion des événements sélectionnés.

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
> Windows : À l’heure actuelle, les importations ne sont pas prises en charge dans le fichier JavaScript dans lequel vous implémentez la gestion de l’activation basée sur les événements.

## <a name="update-the-commands-html-file"></a>Mettre à jour le fichier HTML des commandes

1. Dans le dossier **./src/commands** , ouvrez **commands.html**.

1. Juste avant la balise **head** fermante (`</head>`), ajoutez une entrée de script pour inclure le code JavaScript de gestion des événements.

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. Enregistrez vos modifications.

## <a name="update-webpack-config-settings"></a>Mettre à jour les paramètres de configuration webapck

1. Ouvrez le fichier **webpack.config.js** qui se trouve dans le répertoire racine du projet et effectuez les étapes suivantes.

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
    > Si votre complément n’a pas été automatiquement chargé de manière indépendante, suivez les instructions fournies dans Charger une version test des [compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) pour charger manuellement une version test du complément dans Outlook.

1. Dans Outlook sur le web, créez un message.

    ![Une fenêtre de message dans Outlook sur le web avec l’objet défini sur compose.](../images/outlook-web-autolaunch-1.png)

1. Dans Outlook sur la nouvelle interface utilisateur Mac, créez un message.

    ![Fenêtre de message dans Outlook sur la nouvelle interface utilisateur Mac avec l’objet défini sur composer.](../images/outlook-mac-autolaunch.png)

1. Dans Outlook sur Windows, créez un message.

    ![Fenêtre de message dans Outlook sur Windows avec l’objet défini sur composer.](../images/outlook-win-autolaunch.png)

## <a name="debug"></a>Débogage

Lorsque vous apportez des modifications à la gestion des événements de lancement dans votre complément, vous devez savoir que :

- Si vous avez mis à jour le manifeste, [supprimez le complément, puis chargez-le](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) à nouveau. Si vous utilisez Outlook sur Windows, fermez et rouvrez Outlook.
- Si vous avez apporté des modifications à des fichiers autres que le manifeste, fermez et rouvrez Outlook sur Windows, ou actualisez l’onglet du navigateur en exécutant Outlook sur le web.

Lors de l’implémentation de vos propres fonctionnalités, vous devrez peut-être déboguer votre code. Pour obtenir des conseils sur la façon de déboguer l’activation d’un complément basé sur les événements, consultez [Déboguer votre complément Outlook basé sur les événements](debug-autolaunch.md).

La journalisation du runtime est également disponible pour cette fonctionnalité sur Windows. Pour plus d’informations, consultez [Déboguer votre complément avec la journalisation du runtime](../testing/runtime-logging.md#runtime-logging-on-windows).

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="deploy-to-users"></a>Déployer sur les utilisateurs

Vous pouvez déployer des compléments basés sur des événements en chargeant le manifeste via le Centre d'administration Microsoft 365. Dans le portail d’administration, développez la section **Paramètres** dans le volet de navigation, puis sélectionnez **Applications intégrées**. Dans la page **Applications intégrées** , choisissez l’action **Charger des applications personnalisées** .

![La page Applications intégrées sur le Centre d'administration Microsoft 365 avec l’action Charger des applications personnalisées mise en surbrillance.](../images/outlook-deploy-event-based-add-ins.png)

> [!IMPORTANT]
> Les compléments basés sur les événements sont limités aux déploiements gérés par l’administrateur uniquement. Les utilisateurs ne peuvent pas activer les compléments basés sur les événements à partir d’AppSource ou de l’Office Store dans l’application. Pour plus d’informations, consultez [Options de référencement AppSource pour votre complément Outlook basé sur les événements](autolaunch-store-options.md).

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

## <a name="event-based-activation-behavior-and-limitations"></a>Comportement et limitations de l’activation basée sur les événements

Les gestionnaires d’événements de lancement de compléments sont censés être courts, légers et aussi non invasifs que possible. Après l’activation, votre complément expire dans un délai d’environ 300 secondes, durée maximale autorisée pour l’exécution des compléments basés sur les événements. Pour signaler que votre complément a terminé le traitement d’un événement de lancement, le gestionnaire d’événements associé doit appeler la `event.completed` méthode . (Notez que l’exécution du code inclus après l’instruction `event.completed` n’est pas garantie.) Chaque fois qu’un événement géré par votre complément est déclenché, le complément est réactivé et exécute le gestionnaire d’événements associé, et la fenêtre de délai d’expiration est réinitialisée. Le complément se termine après son expiration, ou l’utilisateur ferme la fenêtre de composition ou envoie l’élément.

Si l’utilisateur a plusieurs compléments qui se sont abonnés au même événement, la plateforme Outlook lance les compléments dans aucun ordre particulier. Actuellement, seuls cinq compléments basés sur des événements peuvent être en cours d’exécution.

Dans tous les clients Outlook pris en charge, l’utilisateur doit rester sur l’élément de courrier actuel dans lequel le complément a été activé pour qu’il s’exécute. Si vous quittez l’élément actif (par exemple, en basculant vers une autre fenêtre de composition ou un autre onglet), l’opération de complément est terminée. Le complément cesse également de fonctionner lorsque l’utilisateur envoie le message ou le rendez-vous qu’il compose.

Les importations ne sont pas prises en charge dans le fichier JavaScript dans lequel vous implémentez la gestion de l’activation basée sur les événements dans le client Windows.

Certaines API Office.js qui modifient ou modifient l’interface utilisateur ne sont pas autorisées à partir de compléments basés sur des événements. Voici les API bloquées.

- Sous `Office.context.auth`:
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > [OfficeRuntime.auth](/javascript/api/office-runtime/officeruntime.auth) est pris en charge dans toutes les versions d’Outlook qui prennent en charge l’activation basée sur les événements et l’authentification unique (SSO), tandis [qu’Office.auth](/javascript/api/office/office.auth) n’est pris en charge que dans certaines builds Outlook. Pour plus d’informations, consultez [Activer l’authentification unique (SSO) dans les compléments Outlook qui utilisent l’activation basée sur les événements](use-sso-in-event-based-activation.md).
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

Vous pouvez demander des données externes à l’aide d’une API telle que [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) ou à l’aide de [XMLHttpRequest (XHR),](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) une API web standard qui émet des requêtes HTTP pour interagir avec les serveurs.

N’oubliez pas que vous devez utiliser des mesures de sécurité supplémentaires lors de l’utilisation d’objets XMLHttpRequest, nécessitant la [même stratégie d’origine](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) et [un simple CORS (Cross-Origin Resource Sharing).](https://developer.mozilla.org/docs/Web/HTTP/CORS)

Une implémentation [CORS simple](https://developer.mozilla.org/docs/Web/HTTP/CORS#simple_requests) :

- Impossible d’utiliser des cookies.
- Prend uniquement en charge les méthodes simples, telles que `GET`, `HEAD`et `POST`.
- Accepte des en-têtes simples avec des noms `Accept`de champ , `Accept-Language`ou `Content-Language`.
- Peut utiliser le `Content-Type`, à condition que le type de contenu soit `application/x-www-form-urlencoded`, `text/plain`ou `multipart/form-data`.
- Les écouteurs d’événements ne peuvent pas être inscrits sur l’objet retourné par `XMLHttpRequest.upload`.
- Impossible d’utiliser des `ReadableStream` objets dans les requêtes.

> [!NOTE]
> La prise en charge complète de CORS est disponible dans Outlook sur le web, Mac et Windows (à partir de la version 2201, build 16.0.14813.10000).

## <a name="see-also"></a>Voir aussi

- [Manifestes de complément Outlook](manifests.md)
- [Comment déboguer des compléments basés sur des événements](debug-autolaunch.md)
- [Options de liste AppSource pour votre complément Outlook basé sur les événements](autolaunch-store-options.md)
- [Procédure pas à pas des alertes intelligentes et OnMessageSend](smart-alerts-onmessagesend-walkthrough.md)
- Exemples de code des compléments Office :
  - [Utiliser l’activation basée sur les événements Outlook pour chiffrer les pièces jointes, traiter les participants aux demandes de réunion et réagir aux modifications de date/heure de rendez-vous](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Utiliser l’activation Outlook basée sur un événement pour définir la signature](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [Utiliser l’activation basée sur les événements Outlook pour marquer les destinataires externes](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
  - [Utiliser Alertes intelligentes d’Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
