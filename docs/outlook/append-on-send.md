---
title: Implémenter Append-on-Send dans votre complément Outlook (aperçu)
description: Découvrez comment implémenter la fonctionnalité Ajout d’envoi dans votre complément Outlook.
ms.topic: article
ms.date: 09/09/2020
localization_priority: Normal
ms.openlocfilehash: 2199f837351c1030e6f6d0d23db7bf81e498d433
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430932"
---
# <a name="implement-append-on-send-in-your-outlook-add-in-preview"></a>Implémenter Append-on-Send dans votre complément Outlook (aperçu)

À la fin de cette procédure pas à pas, vous disposez d’un complément Outlook qui peut insérer une clause d’exclusion de responsabilité lors de l’envoi d’un message.

> [!IMPORTANT]
> Cette fonctionnalité est actuellement [prise en charge pour la](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion dans Outlook sur le Web et Windows avec un abonnement Microsoft 365. Pour plus d’informations, reportez-vous [à la rubrique relative à l’aperçu de la fonctionnalité Ajout à l’envoi](#how-to-preview-the-append-on-send-feature) de cet article.
>
> Les fonctionnalités d’aperçu étant susceptibles d’être modifiées sans préavis, elles ne doivent pas être utilisées dans les compléments de production.

## <a name="how-to-preview-the-append-on-send-feature"></a>Comment afficher un aperçu de la fonctionnalité Ajouter-on-Send

Nous vous invitons à tester la fonctionnalité Ajout à l’envoi ! Faites-nous part de vos scénarios et de vos possibilités d’amélioration en nous donnant des commentaires via GitHub (voir la section **Commentaires** à la fin de cette page).

Pour afficher un aperçu de cette fonctionnalité :

- Faites référence à la bibliothèque **beta** sur le CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . Le [fichier de définition de type](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) pour la compilation de la machine à écrire et IntelliSense se trouve dans le CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` .
- Pour Windows, vous devrez peut-être rejoindre le [programme Office Insider](https://insider.office.com) pour accéder à des builds Office plus récentes.
- Pour Outlook sur le Web, [configurez la version ciblée sur votre client Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).

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

## <a name="try-it-out"></a>Essayez

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
