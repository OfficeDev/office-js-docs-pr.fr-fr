---
title: Utiliser des alertes intelligentes et les événements OnMessageSend et OnAppointmentSend dans votre complément Outlook (préversion)
description: Découvrez comment gérer les événements en envoi dans votre complément Outlook à l’aide de l’activation basée sur les événements.
ms.topic: article
ms.date: 05/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0174d766423a9b70c67b0c2cf559f5b1ea24c9fe
ms.sourcegitcommit: 35e7646c5ad0d728b1b158c24654423d999e0775
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/02/2022
ms.locfileid: "65833926"
---
# <a name="use-smart-alerts-and-the-onmessagesend-and-onappointmentsend-events-in-your-outlook-add-in-preview"></a>Utiliser des alertes intelligentes et les événements OnMessageSend et OnAppointmentSend dans votre complément Outlook (préversion)

Les `OnMessageSend` événements et `OnAppointmentSend` les alertes actives tirent parti des alertes intelligentes, qui vous permettent d’exécuter la logique après qu’un utilisateur a sélectionné **Envoyer** dans son message ou rendez-vous Outlook. Votre gestionnaire d’événements vous permet de donner à vos utilisateurs la possibilité d’améliorer leurs e-mails et invitations aux réunions avant qu’ils ne soient envoyés.

La procédure pas à pas suivante utilise l’événement `OnMessageSend` . À la fin de cette procédure pas à pas, vous disposerez d’un complément qui s’exécute chaque fois qu’un message est envoyé et vérifie si l’utilisateur a oublié d’ajouter un document ou une image qu’il a mentionné dans son e-mail.

> [!IMPORTANT]
> Les `OnMessageSend` événements et `OnAppointmentSend` les événements sont disponibles uniquement en préversion avec un abonnement Microsoft 365 dans Outlook sur Windows. Pour plus d’informations, consultez [La préversion](autolaunch.md#how-to-preview). Les événements d’aperçu ne doivent pas être utilisés dans les compléments de production.

## <a name="prerequisites"></a>Conditions préalables

L’événement `OnMessageSend` est disponible via la fonctionnalité d’activation basée sur les événements. Pour comprendre comment configurer votre complément pour utiliser cette fonctionnalité, utilisez d’autres événements disponibles, configurez la préversion pour cet événement, déboguez votre complément, etc., [reportez-vous à Configurer votre complément Outlook pour l’activation basée sur les événements](autolaunch.md).

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Terminez le [Outlook démarrage rapide](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator), qui crée un projet de complément avec le générateur Yeoman pour Office compléments.

## <a name="configure-the-manifest"></a>Configurer le manifeste

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

1. Sélectionnez l’intégralité du nœud **VersionOverrides** (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant, puis enregistrez vos modifications.

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

> [!TIP]
>
> - Pour **les options SendMode** disponibles avec les événements et `OnAppointmentSend` les `OnMessageSend` événements, reportez-vous aux [options SendMode disponibles](/javascript/api/manifest/launchevent#available-sendmode-options-preview).
> - Pour en savoir plus sur les manifestes pour Outlook compléments, consultez [Outlook manifestes de complément](manifests.md).

## <a name="implement-event-handling"></a>Implémenter la gestion des événements

Vous devez implémenter la gestion de votre événement sélectionné.

Dans ce scénario, vous allez ajouter la gestion de l’envoi d’un message. Votre complément recherche certains mots clés dans le message. Si l’un de ces mots clés est trouvé, il vérifie s’il existe des pièces jointes. S’il n’existe aucune pièce jointe, votre complément recommande à l’utilisateur d’ajouter la pièce jointe éventuellement manquante.

1. À partir du même projet de démarrage rapide, créez un dossier nommé **launchevent** sous le répertoire **./src** .

1. Dans le dossier **./src/launchevent** , créez un fichier nommé **launchevent.js**.

1. Ouvrez le fichier **./src/launchevent/launchevent.js** dans votre éditeur de code et ajoutez le code JavaScript suivant.

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { asyncContext: event },
        getBodyCallback
      );
    }

    function getBodyCallback(asyncResult){
      let event = asyncResult.asyncContext;
      let body = "";
      if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
        body = asyncResult.value;
      } else {
        let message = "Failed to get body text";
        console.error(message);
        event.completed({ allowEvent: false, errorMessage: message });
        return;
      }

      let matches = hasMatches(body);
      if (matches) {
        Office.context.mailbox.item.getAttachmentsAsync(
          { asyncContext: event },
          getAttachmentsCallback);
      } else {
        event.completed({ allowEvent: true });
      }
    }

    function hasMatches(body) {
      if (body == null || body == "") {
        return false;
      }

      const arrayOfTerms = ["send", "picture", "document", "attachment"];
      for (let index = 0; index < arrayOfTerms.length; index++) {
        const term = arrayOfTerms[index].trim();
        const regex = RegExp(term, 'i');
        if (regex.test(body)) {
          return true;
        }
      }

      return false;
    }

    function getAttachmentsCallback(asyncResult) {
      let event = asyncResult.asyncContext;
      if (asyncResult.value.length > 0) {
        for (let i = 0; i < asyncResult.value.length; i++) {
          if (asyncResult.value[i].isInline == false) {
            event.completed({ allowEvent: true });
            return;
          }
        }

        event.completed({ allowEvent: false, errorMessage: "Looks like you forgot to include an attachment?" });
      } else {
        event.completed({ allowEvent: false, errorMessage: "Looks like you're forgetting to include an attachment?" });
      }
    }

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
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
    > Si votre complément n’a pas été chargé automatiquement, suivez les instructions de chargement indépendant [Outlook compléments à tester pour](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) charger manuellement le complément dans Outlook.

1. Dans Outlook sur Windows, créez un message et définissez l’objet. Dans le corps, ajoutez un texte tel que « Hey, regardez cette photo de mon chien ! ».
1. Envoyez le message. Une boîte de dialogue doit s’afficher avec une recommandation vous permettant d’ajouter une pièce jointe.

    ![Boîte de dialogue recommandant que l’utilisateur inclue une pièce jointe.](../images/outlook-win-smart-alert.png)

1. Ajoutez une pièce jointe, puis renvoyez le message. Il ne doit pas y avoir d’alerte cette fois.

## <a name="smart-alerts-feature-behavior-and-scenarios"></a>Comportement et scénarios des fonctionnalités d’alertes intelligentes

Les descriptions des options **SendMode** et les recommandations relatives au moment de leur utilisation sont détaillées dans les [options SendMode disponibles](/javascript/api/manifest/launchevent). L’article suivant décrit le comportement de la fonctionnalité pour certains scénarios.

### <a name="add-in-is-unavailable"></a>Le complément n’est pas disponible

Si le complément n’est pas disponible lors de l’envoi d’un message ou d’un rendez-vous (par exemple, une erreur se produit qui empêche le chargement du complément), l’utilisateur est alerté. Les options disponibles pour l’utilisateur varient en fonction de l’option **SendMode** appliquée au complément.

Si l’option ou `SoftBlock` l’option `PromptUser` est utilisée, l’utilisateur peut choisir **Envoyer quand même** pour envoyer l’élément sans que le complément ne l’vérifie, ou **essayer plus tard** de laisser l’élément être vérifié par le complément lorsqu’il redevient disponible.

![Boîte de dialogue qui avertit l’utilisateur que le complément n’est pas disponible et donne à l’utilisateur la possibilité d’envoyer l’élément maintenant ou ultérieurement.](../images/outlook-soft-block-promptUser-unavailable.png)

Si l’option `Block` est utilisée, l’utilisateur ne peut pas envoyer l’élément tant que le complément n’est pas disponible.

![Boîte de dialogue qui avertit l’utilisateur que le complément n’est pas disponible. L’utilisateur ne peut envoyer l’élément que lorsque le complément est à nouveau disponible.](../images/outlook-hard-block-unavailable.png)

### <a name="long-running-add-in-operations"></a>Opérations de complément de longue durée

Si le complément s’exécute pendant plus de cinq secondes, mais moins de cinq minutes, l’utilisateur est averti que le complément prend plus de temps que prévu pour traiter le message ou le rendez-vous.

Si l’option `PromptUser` est utilisée, l’utilisateur peut choisir **Envoyer quand même** pour envoyer l’élément sans que le complément termine sa vérification. L’utilisateur peut également sélectionner **Ne pas envoyer** pour arrêter le traitement du complément.

![Boîte de dialogue qui avertit l’utilisateur que le complément prend plus de temps que prévu pour traiter l’élément. L’utilisateur peut choisir d’envoyer l’élément sans que le complément termine sa vérification ou d’empêcher le complément de traiter l’élément.](../images/outlook-promptUser-long-running.png)

Toutefois, si l’option ou `Block` l’option `SoftBlock` est utilisée, l’utilisateur ne peut pas envoyer l’élément tant que le complément n’a pas terminé son traitement.

![Boîte de dialogue qui avertit l’utilisateur que le complément prend plus de temps que prévu pour traiter l’élément. L’utilisateur doit attendre que le complément termine le traitement de l’élément avant de pouvoir l’envoyer.](../images/outlook-soft-hard-block-long-running.png)

`OnMessageSend` et `OnAppointmentSend` les compléments doivent être de courte durée et légers. Pour éviter la boîte de dialogue d’opération de longue durée, utilisez d’autres événements pour traiter les vérifications conditionnelles avant l’activation de l’événement`OnMessageSend`.`OnAppointmentSend` Par exemple, si l’utilisateur est tenu de chiffrer les pièces jointes pour chaque message ou rendez-vous, envisagez d’utiliser le ou `OnAppointmentAttachmentsChanged` l’événement `OnMessageAttachmentsChanged` pour effectuer la vérification.

### <a name="add-in-timed-out"></a>Délai d’expiration du complément

Si le complément s’exécute pendant cinq minutes ou plus, il expire. Si l’option `PromptUser` est utilisée, l’utilisateur peut choisir **Envoyer quand même** pour envoyer l’élément sans que le complément termine sa vérification. L’utilisateur peut également choisir **Ne pas envoyer**.

![Boîte de dialogue qui avertit l’utilisateur que le processus de complément a expiré. L’utilisateur peut choisir d’envoyer l’élément sans que le complément termine sa vérification, ou de ne pas envoyer l’élément.](../images/outlook-promptUser-timeout.png)

Si l’option ou `Block` l’option `SoftBlock` est utilisée, l’utilisateur ne peut pas envoyer l’élément tant que le complément n’a pas terminé sa vérification. L’utilisateur doit réessayer d’envoyer l’élément pour réactiver le complément.

![Boîte de dialogue qui avertit l’utilisateur que le processus de complément a expiré. L’utilisateur doit réessayer d’envoyer l’élément pour activer le complément avant de pouvoir envoyer le message ou le rendez-vous.](../images/outlook-soft-hard-block-timeout.png)

## <a name="limitations"></a>Limites

Étant donné que les événements et `OnAppointmentSend` les `OnMessageSend` événements sont pris en charge par le biais de la fonctionnalité d’activation basée sur les événements, les mêmes limitations de fonctionnalité s’appliquent aux compléments qui s’activent à la suite de ces événements. Pour obtenir une description de ces limitations, [reportez-vous au comportement et aux limitations de l’activation basée sur les événements](autolaunch.md#event-based-activation-behavior-and-limitations).

En plus de ces contraintes, une seule instance de l’événement `OnMessageSend` peut `OnAppointmentSend` être déclarée dans le manifeste. Si vous avez besoin de plusieurs `OnMessageSend` événements, `OnAppointmentSend` vous devez les déclarer dans un manifeste ou un complément distinct.

Bien qu’un message de boîte de dialogue Alertes intelligentes puisse être modifié en fonction de votre scénario de complément à l’aide de la [propriété errorMessage](/javascript/api/office/office.addincommands.eventcompletedoptions) de la méthode event.completed, les éléments suivants ne peuvent pas être personnalisés.

- Barre de titre de la boîte de dialogue. Le nom de votre complément s’y affiche toujours.
- Format du message. Par exemple, vous ne pouvez pas modifier la taille et la couleur de police du texte ou insérer une liste à puces.
- Options de la boîte de dialogue. Par exemple, les options **Envoyer quand même** et **Ne pas envoyer** sont fixes et dépendent de [l’option SendMode](/javascript/api/manifest/launchevent) que vous sélectionnez.
- Boîtes de dialogue d’informations sur le traitement et la progression de l’activation basée sur les événements. Par exemple, le texte et les options qui apparaissent dans les dialogues d’expiration et d’opération de longue durée ne peuvent pas être modifiés.

## <a name="see-also"></a>Voir aussi

- [Manifestes de complément Outlook](manifests.md)
- [Configurer votre complément Outlook pour l’activation basée sur les événements](autolaunch.md)
- [Comment déboguer des compléments basés sur des événements](debug-autolaunch.md)
- [Options de liste AppSource pour votre complément Outlook basé sur les événements](autolaunch-store-options.md)
