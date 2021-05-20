---
title: Configurez votre Outlook add-in pour l’activation basée sur l’événement (aperçu)
description: Découvrez comment configurer vos Outlook pour l’activation basée sur l’événement.
ms.topic: article
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 721f05e1c835e066744598ecb2bd416c6a6b0526
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555244"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>Configurez votre Outlook add-in pour l’activation basée sur l’événement (aperçu)

Sans la fonction d’activation basée sur l’événement, un utilisateur doit lancer explicitement un module supplémentaire pour accomplir ses tâches. Cette fonctionnalité permet à votre module d’exécuter des tâches en fonction de certains événements, en particulier pour les opérations qui s’appliquent à chaque élément. Vous pouvez également intégrer avec le volet de tâche et les fonctionnalités sans interface utilisateur.

À la fin de cette procédure pas à pas, vous aurez un add-in qui s’exécute chaque fois qu’un nouvel élément est créé et définit le sujet.

> [!IMPORTANT]
> Cette fonctionnalité n’est prise en [charge que](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) pour un aperçu Outlook sur le web et sur Windows avec un abonnement Microsoft 365 spécial. Pour plus de détails, voir [Comment prévisualiser la fonction d’activation basée sur l’événement](#how-to-preview-the-event-based-activation-feature) dans cet article.
>
> Étant donné que les fonctionnalités d’aperçu sont sujettes à changement sans préavis, elles ne doivent pas être utilisées dans les modules de production.

## <a name="supported-events"></a>Événements pris en charge

À l’heure actuelle, les événements suivants sont pris en charge.

|Événement|Description|Clients|
|---|---|---|
|`OnNewMessageCompose`|Sur la composition d’un nouveau message (inclut la réponse, répondre tous, et en avant) mais pas sur l’édition, par exemple, un projet.|Windows, web|
|`OnNewAppointmentOrganizer`|Sur la création d’un nouveau rendez-vous, mais pas sur l’édition d’un existant.|Windows, web|
|`OnMessageAttachmentsChanged`|Lors de l’ajout ou de la suppression des pièces jointes lors de la composition d’un message.|Windows|
|`OnAppointmentAttachmentsChanged`|Lors de l’ajout ou de la suppression des pièces jointes lors de la composition d’un rendez-vous.|Windows|
|`OnMessageRecipientsChanged`|Lors de l’ajout ou de la suppression de destinataires lors de la composition d’un message.|Windows|
|`OnAppointmentAttendeesChanged`|Sur l’ajout ou la suppression des participants lors de la composition d’un rendez-vous.|Windows|
|`OnAppointmentTimeChanged`|À la date/heure changeante tout en composant un rendez-vous.|Windows|
|`OnAppointmentRecurrenceChanged`|Lors de l’ajout, de la modification ou de la suppression des détails de récurrence lors de la composition d’un rendez-vous. Si la date/heure est modifiée, `OnAppointmentTimeChanged` l’événement sera également déclenché.|Windows|
|`OnInfoBarDismissClicked`|Lors du rejet d’une notification lors de la composition d’un message ou d’un élément de rendez-vous. Seul l’add-in qui a ajouté la notification sera notifié.|Windows|

## <a name="how-to-preview-the-event-based-activation-feature"></a>Comment prévisualiser la fonction d’activation basée sur l’événement

Nous vous invitons à essayer la fonction d’activation basée sur l’événement! Faites-nous part de vos scénarios et de la façon dont nous pouvons nous améliorer en nous donnant des commentaires par GitHub **(voir la** section Commentaires à la fin de cette page).

Pour prévisualiser cette fonctionnalité :

- Pour Outlook sur le web :
  - [Configurez la version ciblée sur votre Microsoft 365 locataire](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).
  - Référencez **la bibliothèque** bêta sur le CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . Le [fichier de définition de type](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) pour la compilation typescript IntelliSense est trouvé à la CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` .
- Pour Outlook sur Windows :
  - La construction minimale requise est de 16.0.14026.20000. Rejoignez [le Office Insider pour](https://insider.office.com) accéder aux versions Office bêta.
  - Configurez le registre. Outlook comprend une copie locale des versions bêta et de production des Office.js au lieu de charger à partir du CDN. Par défaut, la copie de production locale de l’API est référencée. Pour passer à la copie bêta locale des API JavaScript Outlook, vous devez ajouter cette entrée de registre, sinon les API bêta peuvent ne pas être trouvées.
    1. Créez la clé du registre `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` .
    1. Ajouter une entrée nommée `EnableBetaAPIsInJavaScript` et définir la valeur à `1` . L’image suivante indique à quoi doit ressembler le registre.

        ![Capture d’écran de l’éditeur du registre avec une valeur clé du registre EnableBetaAPIsInJavaScript](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Complétez [Outlook démarrage rapide](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet d’ajout avec le générateur Yeoman pour Office add-ins.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour activer l’activation basée sur l’événement de votre module, vous devez configurer [l’élément Runtimes](../reference/manifest/runtimes.md) et le point [d’extension LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` dans le nœud du manifeste. Pour l’instant, `DesktopFormFactor` est le seul facteur de forme pris en charge.

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez **manifest.xml** fichier situé à l’origine de votre projet.

1. Sélectionnez `<VersionOverrides>` l’ensemble du nœud (y compris les balises ouvertes et proches) et remplacez-le par le XML suivant, puis enregistrez vos modifications.

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
              <!-- Events supported on the web and on Windows. -->
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
              <!-- Events supported only on Windows. -->
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

Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web utilise un fichier HTML qui peut référencer le même fichier JavaScript. Vous devez fournir des références à ces deux fichiers dans `Resources` le nœud du manifeste que la plate-forme Outlook détermine en fin de compte s’il faut utiliser HTML ou JavaScript en fonction de la Outlook client. En tant que tel, pour configurer la gestion d’événements, fournir l’emplacement du HTML dans `Runtime` l’élément, puis dans `Override` son élément enfant fournir l’emplacement du fichier JavaScript inlined ou référencé par le HTML.

> [!TIP]
> Pour en savoir plus sur les manifestes Outlook les add-ins, [consultez Outlook manifestes add-in](manifests.md).

## <a name="implement-event-handling"></a>Implémenter la gestion des événements

Vous devez implémenter la manipulation de vos événements sélectionnés.

Dans ce scénario, vous ajouterez la manipulation pour composer de nouveaux éléments.

1. À partir du même projet de démarrage rapide, ouvrez le **fichier ./src/commandes/commands.jsdans** votre éditeur de code.

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

## <a name="try-it-out"></a>Try it out

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution) et votre complément est chargé.

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Si votre module d’ajout n’a pas été automatiquement sideloaded, puis suivez les instructions [dans sideload Outlook add-ins](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) pour les tests pour sideload manuellement l’add-in dans Outlook.

1. Dans Outlook sur le web, créez un message.

    ![Capture d’écran d’une fenêtre de message Outlook sur le web avec le sujet mis sur composer](../images/outlook-web-autolaunch-1.png)

1. En Outlook sur Windows, créez un nouveau message.

    ![Capture d’écran d’une fenêtre de message Outlook sur Windows avec le sujet mis sur composer](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > Si vous lancez votre add-in depuis localhost et que vous voyez l’erreur « Nous sommes désolés, nous n’avons pas *pu accéder à {your-add-in-name-here}*. Assurez-vous d’avoir une connexion réseau. Si le problème persiste, s’il vous plaît réessayer plus tard. », vous devrez peut-être activer une exemption loopback.
    >
    > 1. Fermez Outlook.
    > 1. Ouvrez le **gestionnaire de tâches et** assurez-vous que le processus **msoadfsb.exe'est** pas en cours d’exécution.
    > 1. Exécutez la commande suivante.
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. Redémarrez Outlook.

## <a name="debug"></a>Debug

Lorsque vous modifiez la gestion des événements de lancement dans votre module d’ajout, vous devez savoir que :

- Si vous avez mis à jour le manifeste, [retirez l’add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) puis chargez-le de nouveau.
- Si vous avez apporté des modifications à des fichiers autres que le manifeste, fermez et rouvrez les Outlook sur Windows, ou actualisez l’onglet navigateur en cours d’exécution Outlook sur le Web.

Lors de la mise en œuvre de vos propres fonctionnalités, vous devrez peut-être débogdier votre code. Pour obtenir des conseils sur la façon de déboger l’activation add-in basée sur les événements, [consultez Debug votre module basé sur Outlook’add-in](debug-autolaunch.md).

L’enregistrement de temps d’exécution est également disponible pour cette fonctionnalité Windows. Pour plus d’informations, [consultez Votre add-in avec l’enregistrement de temps d’exécution](../testing/runtime-logging.md#runtime-logging-on-windows).

## <a name="deploy-to-users"></a>Déployer aux utilisateurs

Vous pouvez déployer des modules d’add-in basés sur des événements en téléchargeant le manifeste via le Microsoft 365'administration. Dans le portail admin, élargissez la section **Paramètres** dans le volet navigation puis sélectionnez **applications intégrées**. Sur la page **Applications intégrées,** choisissez l’action **Télécharger’applications personnalisées.**

![Capture d’écran de la page Applications intégrées sur le Microsoft 365 d’administration, y compris l’action Télécharger’applications personnalisées](../images/outlook-deploy-event-based-add-ins.png)

AppSource et magasins inclients : La possibilité de déployer des modules d’ajout basés sur des événements ou de mettre à jour les modules d’activation existants pour inclure la fonction d’activation basée sur l’événement devrait être disponible prochainement.

> [!IMPORTANT]
> Les modules d’accès basés sur des événements sont limités aux déploiements gérés par admin uniquement. Pour l’instant, les utilisateurs ne peuvent pas obtenir d’add-ins basés sur des événements à partir d’AppSource ou de magasins inclients.

## <a name="event-based-activation-behavior-and-limitations"></a>Comportement et limitations d’activation basés sur l’événement

On s’attend à ce que les gestionnaires d’événements de lancement add-in soient de courte durée, légers et non invasifs que possible. Après activation, votre module s’exécutera dans un délai d’environ 300 secondes, soit la durée maximale autorisée pour l’exécution d’add-ins basés sur l’événement. Pour signaler que votre module a terminé le traitement d’un événement de lancement, nous vous recommandons d’appeler la méthode par le gestionnaire `event.completed` associé. (Notez que le code inclus après l’instruction `event.completed` n’est pas garanti pour s’exécuter.) Chaque fois qu’un événement déclenché par vos poignées d’ajout est déclenché, l’add-in est réactivé et exécute le gestionnaire d’événements associé, et la fenêtre de délai d’attente est réinitialisée. L’add-in se termine après qu’il s’arrête, ou l’utilisateur ferme la fenêtre de composition ou envoie l’élément.

Si l’utilisateur dispose de plusieurs modules d’ajout qui se sont abonnés au même événement, la plate-forme Outlook lance les modules sans ordre particulier. Actuellement, seuls cinq modules d’ajout basés sur des événements peuvent être activement en cours d’exécution.

L’utilisateur peut passer ou naviguer loin de l’élément de messagerie actuel où l’add-in a commencé à s’exécuter. L’add-in qui a été lancé terminera son opération en arrière-plan.

Certaines Office.js qui modifient ou modifient l’interface utilisateur ne sont pas autorisées à partir d’add-ins basés sur des événements. Voici les API bloquées :

- Sous `OfficeRuntime.auth` :
  - `getAccessToken`(Windows seulement)
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
- [Comment débobug les modules d’add-in basés sur les événements](debug-autolaunch.md)
