---
title: Utiliser les alertes intelligentes et l’événement OnMessageSend dans votre Outlook de gestion (aperçu)
description: Découvrez comment gérer l’événement d’envoi de message dans Outlook complément à l’aide de l’activation basée sur un événement.
ms.topic: article
ms.date: 12/13/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2412e1a713c2f15a6b04c77eaba6f368d3607dfb
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/15/2021
ms.locfileid: "61514074"
---
# <a name="use-smart-alerts-and-the-onmessagesend-event-in-your-outlook-add-in-preview"></a>Utiliser les alertes intelligentes et l’événement OnMessageSend dans votre Outlook de gestion (aperçu)

`OnMessageSend`L’événement tire parti des alertes intelligentes qui vous  permettent d’exécuter la logique après qu’un utilisateur a sélectionné Envoyer Outlook message. Votre handler d’événements vous permet de donner à vos utilisateurs la possibilité d’améliorer leurs e-mails avant qu’ils ne soit envoyés. `OnAppointmentSend`L’événement est similaire mais s’applique à un rendez-vous.

À la fin de cette walkthrough, vous aurez un module qui s’exécute chaque fois qu’un message est envoyé et vérifie si l’utilisateur a oublié d’ajouter un document ou une image qu’il a mentionnés dans son e-mail.

> [!IMPORTANT]
> Les événements et les événements sont disponibles uniquement en `OnMessageSend` `OnAppointmentSend` prévisualisation avec un abonnement Microsoft 365 dans Outlook sur Windows. Pour plus d’informations, voir [Comment prévisualiser](autolaunch.md#how-to-preview). Les événements d’aperçu ne doivent pas être utilisés dans les modules de production.

## <a name="prerequisites"></a>Conditions préalables

`OnMessageSend`L’événement est disponible via la fonctionnalité d’activation basée sur des événements. Pour comprendre comment configurer votre complément pour utiliser cette fonctionnalité, les événements disponibles, comment afficher un aperçu de cet événement, le débogage, les limitations de fonctionnalités, etc., reportez-vous à Configurer votre complément Outlook pour [l’activation](autolaunch.md)basée sur des événements.

## <a name="set-up-your-environment"></a>Configuration de votre environnement

[Complétez Outlook démarrage](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rapide qui crée un projet de compl?ment avec le générateur Yeoman pour Office compl?ments.

## <a name="configure-the-manifest"></a>Configurer le manifeste

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
               This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
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

          <!-- Enable launching the add-in on the included event. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
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

> [!TIP]
>
> - Pour les options **SendMode** disponibles avec l’événement, `OnMessageSend` reportez-vous aux [options SendMode disponibles.](../reference/manifest/launchevent.md#available-sendmode-options-preview)
> - Pour en savoir plus sur les manifestes de Outlook des Outlook, consultez la Outlook [des manifestes de ces derniers.](manifests.md)

## <a name="implement-event-handling"></a>Implémenter la gestion des événements

Vous devez implémenter la gestion de l’événement sélectionné.

Dans ce scénario, vous allez ajouter la gestion de l’envoi d’un message. Votre add-in recherche certains mots clés dans le message. Si l’un de ces mots clés est trouvé, il vérifie s’il existe des pièces jointes. S’il n’existe aucune pièce jointe, votre add-in recommande à l’utilisateur d’ajouter la pièce jointe éventuellement manquante.

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.

1. Après la `action` fonction, insérez les fonctions JavaScript suivantes.

    ```js
    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { "asyncContext": event },
        function (asyncResult) {
          var event = asyncResult.asyncContext;
          var body = "";
          if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
            body = asyncResult.value;
          }
        
          var arrayOfTerms = ["send", "picture", "document", "attachment"];
          for (var index = 0; index < arrayOfTerms.length; index++) {
            var term = arrayOfTerms[index].trim();
            const regex = RegExp(term, 'i');
            if (regex.test(body)) {
              matches.push(term);
            }
          }
        
          if (matches.length > 0) {
            // Let's verify if there's an attachment!
            Office.context.mailbox.item.getAttachmentsAsync(
              { "asyncContext": event },
              function(result){
                var event = asyncResult.asyncContext;
                if (result.value.length <= 0) {
                  var message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                } else {
                  for (var i=0;i<result.value.length;i++) {
                    if(result.value[i].isInline == false) {
                      event.completed({ allowEvent: true });
                      return;
                    }
                  }
                    
                  var message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                }
              });
            } else {
              event.completed({ allowEvent: true });
            }
          }
        );
    }
    ```

1. Ajoutez le code JavaScript suivant à la fin du fichier.

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
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
    > Si votre application [n’a](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) pas été automatiquement rechargée de manière test, suivez les instructions du chargement de version test des Outlook pour tester le chargement de version test du Outlook.

1. Dans Outlook sur Windows, créez un message et définissez l’objet. Dans le corps, ajoutez du texte tel que « Hey, regardez cette image de mon chien ! ».
1. Envoyez le message. Une boîte de dialogue doit s’ouvrir avec une recommandation pour ajouter une pièce jointe.
1. Ajoutez une pièce jointe, puis renvoyez le message. Il ne doit pas y avoir d’alerte cette fois.

> [!NOTE]
> Si vous exécutez votre add-in à partir de localhost et que vous voyez l’erreur « Nous sommes désolés, nous n’avons pas pu accéder à *{votre-add-in-name-here}*». Assurez-vous que vous avez une connexion réseau. Si le problème persiste, veuillez essayer à nouveau plus tard. », vous devrez peut-être activer une exemption de bouclisation.
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

## <a name="see-also"></a>Voir aussi

- [Manifestes de complément Outlook](manifests.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](autolaunch.md)
- [Comment déboguer des add-ins basés sur des événements](debug-autolaunch.md)
- [Options de liste AppSource pour votre Outlook d’événements](autolaunch-store-options.md)
