---
title: Implémenter l’ajout à l’envoi dans votre Outlook de messagerie
description: Découvrez comment implémenter la fonctionnalité d’ajout à l’envoi dans votre Outlook de messagerie.
ms.topic: article
ms.date: 02/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1a7ec463c3998d9f85f1a085b6b28c43637651bd
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496893"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>Implémenter l’ajout à l’envoi dans votre Outlook de messagerie

À la fin de cette walkthrough, vous aurez un Outlook qui peut insérer une clause d’exclusion de responsabilité lorsqu’un message est envoyé.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.9. Voir [les clients et les plateformes](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="set-up-your-environment"></a>Configuration de votre environnement

[Complétez Outlook démarrage](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rapide qui crée un projet de compl?ment avec le générateur Yeoman pour Office compl?ments.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour activer la fonctionnalité d’ajout à l’envoi dans votre application, `AppendOnSend` vous devez inclure l’autorisation dans la collection [extendedPermissions](/javascript/api/manifest/extendedpermissions).

Pour ce scénario, au lieu `action` d’exécuter la fonction sur le bouton Effectuer une **action** , vous exécuterez la `appendOnSend` fonction.

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez **lemanifest.xml** situé à la racine de votre projet.

1. Sélectionnez l’intégralité `<VersionOverrides>` du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant.

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
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgComposeGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgComposeOpenPaneButton">
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
                        <FunctionName>appendDisclaimerOnSend</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>

              <!-- Configure AppointmentOrganizerCommandSurface extension point to support
              append on sending a new appointment. -->

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
        <ExtendedPermissions>
          <ExtendedPermission>AppendOnSend</ExtendedPermission>
        </ExtendedPermissions>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Pour en savoir plus sur les manifestes de Outlook des modules, voir Outlook [manifestes de ces derniers](manifests.md).

## <a name="implement-append-on-send-handling"></a>Implémenter la gestion de l’envoi

Ensuite, implémentez l’application sur l’événement d’envoi.

> [!IMPORTANT]
> Si votre application implémente également la gestion des [événements `ItemSend`](outlook-on-send-addins.md)d’envoi à l’aide de , `AppendOnSendAsync` l’appel dans le handler d’envoi renvoie une erreur, car ce scénario n’est pas pris en charge.

Pour ce scénario, vous allez implémenter l’application d’une clause d’exclusion de responsabilité à l’élément lorsque l’utilisateur l’envoie.

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.

1. Après la `action` fonction, insérez la fonction JavaScript suivante.

    ```js
    function appendDisclaimerOnSend(event) {
      var appendText =
        '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  
      /**
        *************************************************************
         Ideal Usage - Call the getBodyType API. Use the coercionType
         it returns as the parameter value below.
        *************************************************************
      */
      Office.context.mailbox.item.body.appendOnSendAsync(
        appendText,
        {
          coercionType: Office.CoercionType.Html
        },
        function(asyncResult) {
          console.log(asyncResult);
        }
      );

      event.completed();
    }
    ```

1. À la fin du fichier, ajoutez l’instruction suivante.

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a>Essayez

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre s’il n’est pas déjà en cours d’exécution et que votre application est rechargée de nouveau. 

    ```command&nbsp;line
    npm start
    ```

1. Créez un message et ajoutez-vous à la **ligne À** .

1. Dans le ruban ou le menu de dépassement, **sélectionnez Effectuer une action**.

1. Envoyez le message, puis ouvrez-le à  partir de votre  boîte de réception ou dossier Éléments envoyés pour afficher la clause d’exclusion de responsabilité.

    ![Capture d’écran d’un exemple de message avec la clause d’exclusion de responsabilité à l’envoi Outlook sur le web.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>Voir aussi

[Manifestes de complément Outlook](manifests.md)
