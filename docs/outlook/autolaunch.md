---
title: Configurer votre complément Outlook pour l’activation basée sur les événements (aperçu)
description: Découvrez comment configurer votre complément Outlook pour l’activation basée sur les événements.
ms.topic: article
ms.date: 05/22/2020
localization_priority: Normal
ms.openlocfilehash: 73cdd4949b870d9bc5a5ad2006ce2081575558df
ms.sourcegitcommit: 77617f6ad06e07f5ff8078b26301748f73e2ee01
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/29/2020
ms.locfileid: "44413195"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>Configurer votre complément Outlook pour l’activation basée sur les événements (aperçu)

Sans la fonctionnalité d’activation basée sur un événement, un utilisateur doit lancer explicitement un complément pour effectuer ses tâches. Cette fonctionnalité permet à votre complément d’exécuter des tâches en fonction de certains événements, en particulier pour les opérations qui s’appliquent à chaque élément. Vous pouvez également intégrer le volet des tâches et les fonctionnalités sans interface utilisateur. Pour le moment, les événements pris en charge sont les suivants.

- `OnNewMessageCompose`: Lors de la composition d’un nouveau message (inclut répondre, répondre à tous et transférer)
- `OnNewAppointmentOrganizer`: Lors de la création d’un rendez-vous

À la fin de cette procédure pas à pas, vous disposez d’un complément qui s’exécute chaque fois qu’un nouveau message est créé.

> [!IMPORTANT]
> Cette fonctionnalité est uniquement prise en charge pour l' [Aperçu](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) dans Outlook sur le Web avec un abonnement Office 365. Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](#how-to-preview-the-event-based-activation-feature) dans cet article.
>
> Les fonctionnalités d’aperçu étant susceptibles d’être modifiées sans préavis, elles ne doivent pas être utilisées dans les compléments de production.

## <a name="how-to-preview-the-event-based-activation-feature"></a>Comment afficher un aperçu de la fonctionnalité d’activation basée sur un événement

Nous vous invitons à tester la fonctionnalité d’activation basée sur les événements. Faites-nous part de vos scénarios et de vos possibilités d’amélioration en nous donnant des commentaires via GitHub (voir la section **Commentaires** à la fin de cette page).

Pour afficher un aperçu de cette fonctionnalité :

- Faites référence à la bibliothèque **beta** sur le CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . Le [fichier de définition de type](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) pour la compilation de la machine à écrire et IntelliSense se trouve dans le CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` .
- Demander l’accès à des bits d’aperçu pour Outlook sur le Web à l’aide de votre compte Microsoft 365 en remplissant et envoyant [ce formulaire de demande](https://aka.ms/OWAPreview). Nous allons vous indiquer quand votre client est prêt.

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Terminez le [démarrage rapide Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le générateur Yeoman pour les compléments Office.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour activer l’activation basée sur les événements de votre complément, vous devez configurer l’élément [runtimes](../reference/manifest/runtimes.md) et le point d’extension [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) dans le manifeste. Pour le moment, `DesktopFormFactor` est le seul facteur de forme pris en charge.

1. Dans votre éditeur de code, ouvrez le projet Quick Start.

1. Ouvrez le fichier **Manifest. xml** situé à la racine de votre projet.

1. Sélectionnez le `<VersionOverrides>` nœud entier (y compris les balises ouvrantes et fermantes) et remplacez-le par le code XML suivant.

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
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/runtime.js" />
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

Outlook sur Windows utilise un fichier JavaScript, tandis qu’Outlook sur le Web utilise un fichier HTML qui fait référence au même fichier JavaScript. Vous devez fournir des références à ces deux fichiers dans le manifeste, car la plateforme Outlook détermine finalement s’il faut utiliser du code HTML ou JavaScript basé sur le client Outlook. Par exemple, pour configurer la gestion des événements, indiquez l’emplacement du code HTML dans l' `Runtime` élément, puis dans son `Override` élément enfant, indiquez l’emplacement du fichier JavaScript inline ou référencé par le code html.

> [!TIP]
> Pour en savoir plus sur les manifestes pour les compléments Outlook, consultez la rubrique [manifestes des compléments Outlook](manifests.md).

## <a name="implement-event-handling"></a>Implémenter la gestion des événements

Vous devez implémenter la gestion de vos événements sélectionnés.

Dans ce scénario, vous allez ajouter la gestion de la composition de nouveaux éléments.

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./SRC/Commands/Commands.js** dans votre éditeur de code.

1. Après la `action` fonction, insérez les fonctions JavaScript suivantes.

    ```js
    function onMessageComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function onAppointmentComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function setSubject() {
      Office.context.mailbox.item.subject.setAsync("Set by an event-based add-in!");
    }
    ```

1. À la fin du fichier, ajoutez les instructions suivantes.

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a>Essayez

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).

    ```command&nbsp;line
    npm run dev-server
    ```

1. Suivez les instructions indiquées dans l’article [Chargement de version test des compléments Outlook](sideload-outlook-add-ins-for-testing.md) pour charger le complément dans Outlook.

1. Dans Outlook sur le web, créez un message.

    ![Capture d’écran d’une fenêtre de message dans Outlook sur le Web avec l’objet défini sur la composition.](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a>Comportement et limitations de l’activation basée sur les événements

Les compléments qui s’activent sur la base des événements sont conçus pour une exécution à court terme, jusqu’à 330 secondes seulement. Nous vous recommandons de faire en sorte que votre complément appelle la `event.completed` méthode pour signaler qu’il a terminé le traitement de l’événement Launch. Le complément se termine également lorsque l’utilisateur ferme la fenêtre de composition.

Si l’utilisateur a plusieurs compléments qui s’abonnent au même événement, la plateforme Outlook lance les compléments sans ordre particulier. Actuellement, seuls cinq compléments basés sur les événements peuvent être exécutés activement. Tout complément supplémentaire est placé dans une file d’attente, puis exécuté comme les compléments précédemment actifs sont terminés ou désactivés.

L’utilisateur peut basculer ou naviguer hors de l’élément de courrier actuel dans lequel le complément a commencé. Le complément qui a été lancé terminera son opération en arrière-plan.

Certaines API Office. js qui modifient ou modifient l’interface utilisateur ne sont pas autorisées dans les compléments basés sur des événements. Les API bloquées sont les suivantes.

- Sous `Office.context.mailbox` :
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Sous `Office.context.ui` :
  - `displayDialogAsync`
  - `messageParent`
- Sous `Office.context.auth` :
  - `getAccessTokenAsync`

## <a name="see-also"></a>Voir aussi

[Manifestes de complément Outlook](manifests.md)
