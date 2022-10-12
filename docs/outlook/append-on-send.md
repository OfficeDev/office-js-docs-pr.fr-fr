---
title: Implémenter append-on-send dans votre complément Outlook
description: Découvrez comment implémenter la fonctionnalité d’ajout à l’envoi dans votre complément Outlook.
ms.topic: article
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 18d3e8300a53d08cf484f14cd4fd05adf6382fe3
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541142"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>Implémenter append-on-send dans votre complément Outlook

À la fin de cette procédure pas à pas, vous disposerez d’un complément Outlook qui peut insérer une clause d’exclusion de responsabilité lors de l’envoi d’un message.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.9. Voir [les clients et les plateformes](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Terminez le [démarrage rapide d’Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le générateur Yeoman pour les compléments Office.

> [!NOTE]
> Si vous souhaitez utiliser le [manifeste Teams pour les compléments Office (préversion),](../develop/json-manifest-overview.md) **suivez** le guide de démarrage rapide [d’Outlook avec un manifeste Teams (préversion),](../quickstarts/outlook-quickstart-json-manifest.md) mais ignorez toutes les sections après la section Essayer.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour configurer le manifeste, ouvrez l’onglet correspondant au type de manifeste que vous utilisez.

# <a name="xml-manifest"></a>[Manifeste XML](#tab/xmlmanifest)

Pour activer la fonctionnalité d’ajout à l’envoi dans votre complément, vous devez inclure l’autorisation `AppendOnSend` dans la collection de [ExtendedPermissions](/javascript/api/manifest/extendedpermissions).

Pour ce scénario, au lieu d’exécuter la `action` fonction lors du choix du bouton **Exécuter une action** , vous allez exécuter la `appendOnSend` fonction.

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

1. Sélectionnez l’intégralité **\<VersionOverrides\>** du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant.

    ```XML
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Requirements>
          <bt:Sets DefaultMinVersion="1.9">
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

# <a name="teams-manifest-developer-preview"></a>[Manifeste Teams (préversion du développeur)](#tab/jsonmanifest)

1. Ouvrez le fichier manifest.json.

1. Ajoutez l’objet suivant au tableau « extensions.runtimes ». Notez ce qui suit à propos de ce code.

   - La valeur « minVersion » de l’ensemble de conditions requises de boîte aux lettres est définie sur « 1.9 » afin que le complément ne puisse pas être installé sur les plateformes et les versions d’Office où cette fonctionnalité n’est pas prise en charge. 
   - L'« ID » du runtime est défini sur le nom descriptif « function_command_runtime ».
   - La propriété « code.page » est définie sur l’URL du fichier HTML sans interface utilisateur qui chargera la commande de fonction.
   - La propriété « lifetime » est définie sur « short », ce qui signifie que le runtime démarre lorsque le bouton de commande de fonction est sélectionné et s’arrête une fois la fonction terminée. (Dans certains cas rares, le runtime s’arrête avant la fin du gestionnaire. Voir [Runtimes in Office Add-ins](../testing/runtimes.md).)
   - Il existe une action pour exécuter une fonction nommée « appendDisclaimerOnSend ». Vous allez créer cette fonction dans une étape ultérieure.

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.9"
                }
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "id": "function_command_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "appendDisclaimerOnSend",
                "type": "executeFunction",
                "displayName": "appendDisclaimerOnSend"
            }
        ]
    }
    ```

1. Dans le tableau « authorization.permissions.resourceSpecific », ajoutez l’objet suivant. Assurez-vous qu’il est séparé des autres objets du tableau par une virgule.

    ```json
    {
      "name": "Mailbox.AppendOnSend.User",
      "type": "Delegated"
    }
    ```

---

> [!TIP]
> Pour en savoir plus sur les manifestes pour les compléments Outlook, consultez [les manifestes de complément Outlook](manifests.md).

## <a name="implement-append-on-send-handling"></a>Implémenter la gestion des ajouts sur l’envoi

Ensuite, implémentez l’ajout sur l’événement d’envoi.

> [!IMPORTANT]
> Si votre complément implémente également la [gestion des événements lors de l’envoi à l’aide `ItemSend`](outlook-on-send-addins.md), l’appel `AppendOnSendAsync` dans le gestionnaire d’envoi retourne une erreur, car ce scénario n’est pas pris en charge.

Pour ce scénario, vous allez implémenter l’ajout d’une clause d’exclusion de responsabilité à l’élément lorsque l’utilisateur l’envoie.

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.

1. Après la `action` fonction, insérez la fonction JavaScript suivante.

    ```js
    function appendDisclaimerOnSend(event) {
      const appendText =
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

1. Juste en dessous de la fonction, ajoutez la ligne suivante pour inscrire la fonction.

    ```js
    Office.actions.associate("appendDisclaimerOnSend", appendDisclaimerOnSend);
    ```

## <a name="try-it-out"></a>Essayez

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre s’il n’est pas déjà en cours d’exécution et que votre complément est chargé de manière indépendante.

    ```command&nbsp;line
    npm start
    ```

1. Créez un message et ajoutez-vous à la ligne **À** .

1. Dans le menu du ruban ou du dépassement de capacité, choisissez **Effectuer une action**.

1. Envoyez le message, puis ouvrez-le à partir de votre dossier **Boîte de réception** ou **Éléments envoyés** pour afficher la clause d’exclusion de responsabilité ajoutée.

    ![Exemple de message avec l’exclusion de responsabilité ajoutée lors de l’envoi de Outlook sur le web.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>Voir aussi

[Manifestes de complément Outlook](manifests.md)
