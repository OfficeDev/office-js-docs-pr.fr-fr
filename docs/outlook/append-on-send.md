---
title: Implémenter Append-on-Send dans votre complément Outlook
description: Découvrez comment implémenter la fonctionnalité Ajout d’envoi dans votre complément Outlook.
ms.topic: article
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 62234f580f6ff6be418f1c252510f234e297b0c6
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626455"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>Implémenter Append-on-Send dans votre complément Outlook

À la fin de cette procédure pas à pas, vous disposez d’un complément Outlook qui peut insérer une clause d’exclusion de responsabilité lors de l’envoi d’un message.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1,9. Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Terminez le [démarrage rapide Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le générateur Yeoman pour les compléments Office.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour activer la fonctionnalité Ajout à l’envoi dans votre complément, vous devez inclure l' `AppendOnSend` autorisation dans la collection de [ExtendedPermissions](../reference/manifest/extendedpermissions.md).

Pour ce scénario, au lieu d’exécuter la `action` fonction en cliquant sur le bouton **effectuer une action** , vous exécuterez `appendOnSend` la fonction.

1. Dans votre éditeur de code, ouvrez le projet Quick Start.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

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
> Pour en savoir plus sur les manifestes pour les compléments Outlook, consultez la rubrique [manifestes des compléments Outlook](manifests.md).

## <a name="implement-append-on-send-handling"></a>Implémenter la gestion des ajouts à l’envoi

Ensuite, implémentez l’ajout sur l’événement Send.

> [!IMPORTANT]
> Si votre complément implémente également la [gestion des événements d’envoi à l' `ItemSend` aide ](outlook-on-send-addins.md)de, l’appel `AppendOnSendAsync` dans le gestionnaire d’envoi renvoie une erreur dans la mesure où ce scénario n’est pas pris en charge.

Pour ce scénario, vous allez implémenter l’ajout d’une clause d’exclusion de responsabilité à l’élément lorsque l’utilisateur envoie.

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

## <a name="try-it-out"></a>Try it out

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur Web local démarre s’il n’est pas déjà en cours d’exécution.

    ```command&nbsp;line
    npm run dev-server
    ```

1. Suivez les instructions de [chargement compléments Outlook à des fins de test](sideload-outlook-add-ins-for-testing.md).

1. Créez un message et ajoutez-vous à la ligne **à** .

1. Dans le menu du ruban ou du buffer overflow, sélectionnez **effectuer une action**.

1. Envoyez le message, puis ouvrez-le à partir de votre dossier **boîte de réception** ou **éléments envoyés** pour afficher la clause d’exclusion de responsabilité ajoutée.

    ![Capture d’écran d’un exemple de message avec la clause d’exclusion de responsabilité ajoutée lors de l’envoi dans Outlook sur le Web.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>Voir aussi

[Manifestes de complément Outlook](manifests.md)
