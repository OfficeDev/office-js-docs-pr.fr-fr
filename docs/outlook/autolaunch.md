---
title: Configurer votre complément Outlook pour l’activation basée sur des événements
description: Découvrez comment configurer votre complément Outlook pour l’activation basée sur des événements.
ms.topic: article
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: ff1dc8da523d752d616981a570b4c83d9f1a423d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349013"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>Configurer votre complément Outlook pour l’activation basée sur des événements

Sans la fonctionnalité d’activation basée sur des événements, un utilisateur doit lancer explicitement un complément pour effectuer ses tâches. Cette fonctionnalité permet à votre application d’exécuter des tâches basées sur certains événements, en particulier pour les opérations qui s’appliquent à chaque élément. Vous pouvez également intégrer le volet Des tâches et la fonctionnalité sans interface utilisateur.

À la fin de cette walkthrough, vous aurez un add-in qui s’exécute chaque fois qu’un nouvel élément est créé et définit l’objet.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans [l’ensemble de conditions requises 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md). Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="supported-events"></a>Événements pris en charge

Pour l’instant, les événements suivants sont pris en charge sur le web et sur Windows.

|Événement|Description|Minimum<br>ensemble de conditions requises|
|---|---|---|
|`OnNewMessageCompose`|Lors de la composition d’un nouveau message (y compris répondre, répondre à tous et transmettre), mais pas lors de la modification, par exemple, d’un brouillon.|1.10|
|`OnNewAppointmentOrganizer`|Lors de la création d’un rendez-vous, mais pas de la modification d’un rendez-vous existant.|1.10|
|`OnMessageAttachmentsChanged`|Lors de l’ajout ou de la suppression de pièces jointes lors de la composition d’un message.|Aperçu|
|`OnAppointmentAttachmentsChanged`|Lors de l’ajout ou de la suppression de pièces jointes lors de la composition d’un rendez-vous.|Aperçu|
|`OnMessageRecipientsChanged`|Lors de l’ajout ou de la suppression de destinataires lors de la composition d’un message.|Aperçu|
|`OnAppointmentAttendeesChanged`|Lors de l’ajout ou de la suppression de participants lors de la composition d’un rendez-vous.|Aperçu|
|`OnAppointmentTimeChanged`|Lors de la modification de la date et de l’heure lors de la composition d’un rendez-vous.|Aperçu|
|`OnAppointmentRecurrenceChanged`|Lors de l’ajout, de la modification ou de la suppression des détails de la récurrence lors de la composition d’un rendez-vous. Si la date/l’heure est modifiée, `OnAppointmentTimeChanged` l’événement est également déclenché.|Aperçu|
|`OnInfoBarDismissClicked`|Lors du rejet d’une notification lors de la composition d’un élément de message ou de rendez-vous. Seul le add-in qui a ajouté la notification sera averti.|Aperçu|

> [!IMPORTANT]
> Les événements toujours en prévisualisation sont disponibles uniquement avec un abonnement Microsoft 365 dans Outlook sur le web et Windows. Pour plus d’informations, voir [La prévisualisation](#how-to-preview) dans cet article. Les événements d’aperçu ne doivent pas être utilisés dans les modules de production.

### <a name="how-to-preview"></a>Comment prévisualiser

Nous vous invitons à tester les événements maintenant en prévisualisation ! Faites-nous part de vos scénarios et de la façon dont nous pouvons les améliorer en nous faisant part de vos commentaires GitHub (voir la **section** Commentaires à la fin de cette page).

Pour afficher un aperçu de ces événements :

- Pour Outlook sur le web :
  - [Configurez la version ciblée sur votre Microsoft 365 client.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)
  - Référencez **la bibliothèque** bêta sur le CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . Le [fichier de définition de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) type pour la compilation et la IntelliSense TypeScript se trouve aux CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` .
- Pour Outlook sur Windows :
  - La build minimale requise est 16.0.14026.20000. Rejoignez le [Office Insider pour](https://insider.office.com) accéder à Office versions bêta.
  - Configurez le Registre. Outlook inclut une copie locale des versions de production et bêta de Office.js au lieu de charger à partir du CDN. Par défaut, la copie de production locale de l’API est référencé. Pour basculer vers la copie bêta locale des API JavaScript Outlook, vous devez ajouter cette entrée de Registre, sinon les API bêta risquent de ne pas être trouvées.
    1. Créez la clé de `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` Registre.
    1. Ajoutez une entrée nommée `EnableBetaAPIsInJavaScript` et définissez la valeur sur `1` . L’image suivante indique à quoi doit ressembler le registre.

        ![Capture d’écran de l’éditeur du Registre avec une valeur de clé de Registre EnableBetaAPIsInJavaScript.](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a>Configuration de votre environnement

[Complétez Outlook démarrage](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rapide qui crée un projet de compl?ment avec le générateur Yeoman pour Office compl?ments.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour activer l’activation basée sur des événements de votre complément, vous devez configurer l’élément [Runtimes](../reference/manifest/runtimes.md) et le point d’extension [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) dans le nœud `VersionOverridesV1_1` du manifeste. Pour l’instant, `DesktopFormFactor` est le seul facteur de forme pris en charge.

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez **lemanifest.xml** situé à la racine de votre projet.

1. Sélectionnez l’intégralité du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le par le `<VersionOverrides>` code XML suivant, puis enregistrez vos modifications.

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
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
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
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
        <!-- Entry needed for Outlook Desktop. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
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

Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web utilise un fichier HTML qui peut référencer le même fichier JavaScript. Vous devez fournir des références à ces deux fichiers dans le nœud du manifeste, car la plateforme Outlook détermine en fin de compte s’il faut utiliser du code HTML ou JavaScript en fonction du `Resources` client Outlook. En tant que tel, pour configurer la gestion des événements, fournissez l’emplacement du code HTML dans l’élément, puis, dans son élément enfant, fournissez l’emplacement du fichier JavaScript indiqué ou référencé par le `Runtime` `Override` code HTML.

> [!TIP]
> Pour en savoir plus sur les manifestes de Outlook des Outlook, consultez la Outlook [des manifestes de ces derniers.](manifests.md)

## <a name="implement-event-handling"></a>Implémenter la gestion des événements

Vous devez implémenter la gestion de vos événements sélectionnés.

Dans ce scénario, vous allez ajouter la gestion de la composition de nouveaux éléments.

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.

1. Après la `action` fonction, insérez les fonctions JavaScript suivantes.

    ```js
    function onMessageComposeHandler(event) {
      setSubject(event);
    }
    function onAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext" : event
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
    ```

1. Ajoutez le code JavaScript suivant à la fin du fichier.

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. Enregistrez vos modifications.

> [!IMPORTANT]
> Windows : actuellement, les importations ne sont pas pris en charge dans le fichier JavaScript où vous implémentez la gestion de l’activation basée sur des événements.

## <a name="try-it-out"></a>Essayez

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution) et votre complément est chargé.

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Si votre application [n’a](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) pas été automatiquement chargé de manière test, suivez les instructions du chargement de version test des Outlook pour tester le chargement de version test du Outlook.

1. Dans Outlook sur le web, créez un message.

    ![Capture d’écran d’une fenêtre de message Outlook sur le web avec l’objet de la composition.](../images/outlook-web-autolaunch-1.png)

1. Dans Outlook sur Windows, créez un message.

    ![Capture d’écran d’une fenêtre de message Outlook sur Windows avec l’objet de la composition.](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > Si vous exécutez votre add-in à partir de l’host local et que vous voyez l’erreur « Désolé, nous n’avons pas pu accéder à *{votre-add-in-name-here}*». Assurez-vous que vous avez une connexion réseau. Si le problème persiste, veuillez essayer à nouveau plus tard. », vous devrez peut-être activer une exemption de bouclisation.
    >
    > 1. Fermez Outlook.
    > 1. Ouvrez **le Gestionnaire des tâches** et assurez-vous que le processus **msoadfsb.exe** n’est pas en cours d’exécution.
    > 1. Exécutez la commande suivante :
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. Redémarrez Outlook.

## <a name="debug"></a>Debug

Lorsque vous modifiez la gestion des événements de lancement dans votre add-in, vous devez savoir que :

- Si vous avez mis à jour le manifeste, [supprimez le add-in,](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) puis chargez-le de nouveau.
- Si vous avez apporté des modifications à des fichiers autres que le manifeste, fermez et rouvrez Outlook sur Windows ou actualisez l’onglet du navigateur en cours d’exécution Outlook sur le web.

Lors de l’implémentation de vos propres fonctionnalités, vous devrez peut-être déboguer votre code. Pour obtenir des instructions sur le débogage de l’activation de complément basée sur des événements, voir [Déboguer](debug-autolaunch.md)votre complément basé sur Outlook événement.

La journalisation runtime est également disponible pour cette fonctionnalité sur Windows. Pour plus d’informations, voir [Déboguer votre add-in avec la journalisation runtime.](../testing/runtime-logging.md#runtime-logging-on-windows)

## <a name="deploy-to-users"></a>Déployer pour les utilisateurs

Vous pouvez déployer des add-ins basés sur des événements en chargeant le manifeste via le Centre d’administration Microsoft 365. Dans le portail d’administration, **développez la** section Paramètres dans le volet de navigation, puis sélectionnez **Applications intégrées.** Dans la page **Applications intégrées,** sélectionnez l Télécharger **d’applications personnalisées.**

![Capture d’écran de la page Applications intégrées sur le Centre d’administration Microsoft 365, y compris l’action Télécharger d’applications personnalisées.](../images/outlook-deploy-event-based-add-ins.png)

Magasins AppSource et inclients : la possibilité de déployer des compléments basés sur des événements ou de mettre à jour des compléments existants pour inclure la fonctionnalité d’activation basée sur des événements devrait être disponible prochainement.

> [!IMPORTANT]
> Les add-ins basés sur des événements sont limités aux déploiements gérés par l’administrateur uniquement. Pour l’instant, les utilisateurs ne peuvent pas obtenir de add-ins basés sur des événements à partir d’AppSource ou de magasins inclients.

## <a name="event-based-activation-behavior-and-limitations"></a>Comportement et limitations de l’activation basée sur des événements

Les handlers d’événements de lancement de modules sont censés être de courte durée, légers et aussi peu invasifs que possible. Après l’activation, votre complément prendra un délai d’environ 300 secondes, durée maximale autorisée pour l’exécution de compléments basés sur des événements. Pour signaler que votre add-in a terminé le traitement d’un événement de lancement, nous vous recommandons d’avoir le handler associé qui appelle la `event.completed` méthode. (Notez que le code inclus après `event.completed` l’instruction n’est pas garanti pour s’exécuter.) Chaque fois qu’un événement géré par votre add-in est déclenché, celui-ci est réactivé et exécute le handler d’événements associé, et la fenêtre d’délai est réinitialisée. Le add-in se termine une fois qu’il n’est plus à son terme, ou l’utilisateur ferme la fenêtre de composition ou envoie l’élément.

Si l’utilisateur a plusieurs add-ins abonnés au même événement, la plateforme Outlook lance les modules dans un ordre particulier. Actuellement, seuls cinq add-ins basés sur des événements peuvent être activement en cours d’exécution.

L’utilisateur peut basculer ou naviguer à partir de l’élément de messagerie actuel où le module a commencé à s’exécute. Le module qui a été lancé terminera son opération en arrière-plan.

Les importations ne sont pas pris en charge dans le fichier JavaScript où vous implémentez la gestion de l’activation basée sur des événements dans Windows client.

Certaines Office.js API qui modifient ou modifient l’interface utilisateur ne sont pas autorisées à partir des add-ins basés sur des événements. Les API bloquées sont les suivantes.

- Sous `OfficeRuntime.auth` :
  - `getAccessToken`(Windows uniquement)
- Sous `Office.context.auth` :
  - `getAccessToken`
  - `getAccessTokenAsync`
- Sous `Office.context.mailbox` :
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Sous `Office.context.mailbox.item` :
  - `close`
- Sous `Office.context.ui` :
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a>Voir aussi

- [Manifestes de complément Outlook](manifests.md)
- [Comment déboguer des add-ins basés sur des événements](debug-autolaunch.md)
- Exemples PnP :
  - [Utiliser Outlook’activation basée sur un événement pour définir la signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)
  - [Utiliser Outlook’activation basée sur un événement pour baliser des destinataires externes](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-tag-external)
