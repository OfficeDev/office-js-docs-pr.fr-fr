---
title: Activer votre complément Outlook sur plusieurs messages (préversion)
description: Découvrez comment activer votre complément Outlook lorsque plusieurs messages sont sélectionnés.
ms.topic: article
ms.date: 10/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2b77772aa2fc661e4be84c48555e3ddceda224c4
ms.sourcegitcommit: 787fbe4d4a5462ff6679ad7fd00748bf07391610
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68546437"
---
# <a name="activate-your-outlook-add-in-on-multiple-messages-preview"></a>Activer votre complément Outlook sur plusieurs messages (préversion)

Avec la fonctionnalité multi-sélection d’élément, votre complément Outlook peut désormais activer et effectuer des opérations sur plusieurs messages sélectionnés en une seule fois. Certaines opérations, telles que le chargement de messages dans votre système CRM (Customer Relationship Management) ou la catégorisation de nombreux éléments, peuvent désormais être facilement effectuées en un seul clic.

Les sections suivantes expliquent comment configurer votre complément pour récupérer la ligne d’objet de plusieurs messages en mode lecture.

> [!IMPORTANT]
> La fonctionnalité de sélection multiple de l’élément est disponible uniquement en préversion avec un abonnement Microsoft 365 dans Outlook sur Windows. Les fonctionnalités en préversion ne doivent pas être utilisées dans les compléments de production. Nous vous invitons à tester cette fonctionnalité dans des environnements de test ou de développement et à recevoir des commentaires sur votre expérience via GitHub (voir la section **Commentaires** à la fin de cette page).

> [!NOTE]
> La fonctionnalité multi-sélection d’élément n’est actuellement pas prise en charge dans le [manifeste Teams (préversion),](../develop/json-manifest-overview.md) mais l’équipe de fonctionnalités s’efforce de rendre cette fonctionnalité disponible.

## <a name="prerequisites-to-preview-item-multi-select"></a>Prérequis pour afficher un aperçu de l’élément à sélection multiple

Pour afficher un aperçu de la fonctionnalité à sélection multiple, installez Outlook sur Windows, à compter de la version 2209 (build 15629.20110). Une fois installé, rejoignez le [programme Office Insider](https://insider.office.com/join/windows) et sélectionnez l’option **Canal bêta** pour accéder aux versions bêta d’Office.

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Terminez le [démarrage rapide d’Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) pour créer un projet de complément avec le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md).

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour permettre à votre complément de s’activer sur plusieurs messages sélectionnés, vous devez ajouter l’élément enfant [SupportsMultiSelect](/javascript/api/manifest/action?view=outlook-js-preview&preserve-view=true#supportsmultiselect-preview) à l’élément **\<Action\>** et définir sa valeur `true`sur . Étant donné que l’élément multi-sélection prend uniquement en charge les messages pour le moment, la valeur d’attribut de l’élément **\<ExtensionPoint\>** doit être définie `MessageReadCommandSurface` sur ou `MessageComposeCommandSurface`.`xsi:type`

1. Dans votre éditeur de code préféré, ouvrez le projet de démarrage rapide Outlook que vous avez créé.

1. Ouvrez le fichier **manifest.xml** situé à la racine du projet.

1. Affectez la **\<Permissions\>** valeur à l’élément `ReadWriteMailbox` .

    ```xml
    <Permissions>ReadWriteMailbox</Permissions>
    ```

1. Sélectionnez l’intégralité **\<VersionOverrides\>** du nœud et remplacez-le par le code XML suivant.

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
        <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
            <Requirements>
                <bt:Sets DefaultMinVersion="1.12">
                  <bt:Set Name="Mailbox"/>
                </bt:Sets>
            </Requirements>
            <Hosts>
                <Host xsi:type="MailHost">
                    <DesktopFormFactor>
                        <!-- Message Read mode-->
                        <ExtensionPoint xsi:type="MessageReadCommandSurface">
                            <OfficeTab id="TabDefault">
                                <Group id="msgReadGroup">
                                    <Label resid="GroupLabel"/>
                                    <Control xsi:type="Button" id="msgReadOpenPaneButton">
                                        <Label resid="TaskpaneButton.Label"/>
                                        <Supertip>
                                            <Title resid="TaskpaneButton.Label"/>
                                            <Description resid="TaskpaneButton.Tooltip"/>
                                        </Supertip>
                                        <Icon>
                                            <bt:Image size="16" resid="Icon.16x16"/>
                                            <bt:Image size="32" resid="Icon.32x32"/>
                                            <bt:Image size="80" resid="Icon.80x80"/>
                                        </Icon>
                                        <Action xsi:type="ShowTaskpane">
                                            <SourceLocation resid="Taskpane.Url"/>
                                            <!-- Enables your add-in to activate on multiple selected messages. -->
                                            <SupportsMultiSelect>true</SupportsMultiSelect>
                                        </Action>
                                    </Control>
                                </Group>
                            </OfficeTab>
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
                  <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
                </bt:Urls>
                <bt:ShortStrings>
                  <bt:String id="GroupLabel" DefaultValue="Item Multi-select"/>
                  <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
                </bt:ShortStrings>
                <bt:LongStrings>
                  <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane which displays an option to retrieve the subject line of selected messages."/>
                </bt:LongStrings>
            </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

1. Enregistrez vos modifications.

## <a name="configure-the-task-pane"></a>Configurer le volet Office

L’élément multi-sélection s’appuie sur l’événement [SelectedItemsChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true) pour déterminer quand les messages sont sélectionnés ou désélectionnés. Cet événement nécessite une implémentation du volet Office.

1. Dans le dossier **./src/taskpane** , ouvrez **taskpane.html**.

1. Dans l’élément **\<script\>** , définissez l’attribut sur `src` `"https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"`. Cela fait référence à la bibliothèque bêta sur le réseau de distribution de contenu (CDN).

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"></script>
    ```

1. Dans l’élément **\<body\>** , remplacez l’élément entier **\<main\>** par le balisage suivant.

    ```html
    <main id="app-body" class="ms-welcome__main">
        <h2 class="ms-font-xl">Retrieve the subject line of multiple messages with one click!</h2>
        <ul id="selected-items"></ul>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. Enregistrez vos modifications.

## <a name="implement-a-handler-for-the-selecteditemschanged-event"></a>Implémenter un gestionnaire pour l’événement SelectedItemsChanged

Pour alerter votre complément lorsque l’événement `SelectedItemsChanged` se produit, vous devez inscrire un gestionnaire d’événements à l’aide de la `addHandlerAsync` méthode.

1. Dans le dossier **./src/taskpane** , ouvrez **taskpane.js**.

1. Dans la `Office.onReady()` fonction de rappel, remplacez le code existant par ce qui suit :

    ```javascript
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    
        // Register an event handler to identify when messages are selected.
        Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, asyncResult => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          }
    
          console.log("Event handler added.");
        });
    }
    ```

## <a name="retrieve-the-subject-line-of-selected-messages"></a>Récupérer la ligne d’objet des messages sélectionnés

Maintenant que vous avez inscrit un gestionnaire d’événements, vous appelez la méthode [getSelectedItemsAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getselecteditemsasync-member(1)) pour récupérer la ligne d’objet des messages sélectionnés et les consigner dans le volet Office. La `getSelectedItemsAsync` méthode peut également être utilisée pour obtenir d’autres propriétés de message, telles que l’ID d’élément, le type d’élément (`Message` est le seul type pris en charge pour l’instant) et le mode élément (`Read` ou `Compose`).

1. Dans **taskpane.js**, accédez à la `run` fonction et insérez le code suivant.

    ```javascript
    // Clear list of previously selected messages, if any.
    const list = document.getElementById("selected-items");
    while (list.firstChild) {
        list.removeChild(list.firstChild);
    }

    // Retrieve the subject line of the selected messages and log it to a list in the task pane.
    Office.context.mailbox.getSelectedItemsAsync(asyncResult => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;      
        }

        asyncResult.value.forEach(item => {
            const listItem = document.createElement("li");
            listItem.textContent = item.subject;
            list.appendChild(listItem);
        });
    });
    ```

1. Enregistrez vos modifications.

## <a name="try-it-out"></a>Essayez

1. À partir d’un terminal, exécutez le code suivant dans le répertoire racine de votre projet. Cela démarre le serveur web local et charge de manière indépendante votre complément.

    ```command line
    npm start
    ```

    > [!TIP]
    > Si votre complément ne se charge pas automatiquement, suivez les instructions [fournies dans Le chargement indépendant des compléments Outlook à des fins de test pour](sideload-outlook-add-ins-for-testing.md?tabs=windows#outlook-on-the-desktop) le charger manuellement dans Outlook.

1. Dans Outlook, vérifiez que le volet de lecture est activé. Pour activer le volet de lecture, consultez [Utiliser et configurer le volet de lecture pour afficher un aperçu des messages](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0).

1. Accédez à votre boîte de réception et choisissez plusieurs messages en maintenant **la touche Ctrl** enfoncée lors de la sélection des messages.

1. Sélectionnez **Afficher le volet Office** dans le ruban.

1. Dans le volet Office, sélectionnez **Exécuter** pour afficher la liste des lignes d’objet des messages sélectionnés.

    :::image type="content" source="../images/outlook-multi-select.png" alt-text="Exemple de liste de lignes d’objet récupérées à partir de plusieurs messages sélectionnés.":::

## <a name="item-multi-select-behavior-and-limitations"></a>Comportement et limitations des sélections multiples d’éléments

L’élément multi-sélection prend uniquement en charge les messages d’une boîte aux lettres Exchange en mode lecture et composition. Un complément Outlook s’active uniquement sur plusieurs messages si les conditions suivantes sont remplies.

- Les messages doivent être sélectionnés à partir d’une boîte aux lettres Exchange à la fois. Les boîtes aux lettres non Exchange ne sont pas prises en charge.
- Les messages doivent être sélectionnés dans un dossier de boîte aux lettres à la fois. Un complément ne s’active pas sur plusieurs messages s’ils se trouvent dans des dossiers différents, sauf si la vue Conversations est activée. Pour plus d’informations, consultez [Sélection multiple dans les conversations](#multi-select-in-conversations).
- Un complément doit implémenter un volet Office pour détecter l’événement `SelectedItemsChanged` .
- Le [volet de lecture](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0) dans Outlook doit être activé.
- Un maximum de 100 messages peuvent être sélectionnés à la fois.

> [!NOTE]
> Les invitations et réponses aux réunions sont considérées comme des messages, et non des rendez-vous, et peuvent donc être incluses dans une sélection.

### <a name="multi-select-in-conversations"></a>Sélection multiple dans les conversations

L’élément multi-sélection prend en charge [l’affichage Conversations](https://support.microsoft.com/office/0eeec76c-f59b-4834-98e6-05cfdfa9fb07) , qu’il soit activé sur votre boîte aux lettres ou sur des dossiers spécifiques. Le tableau suivant décrit les comportements attendus lorsque les conversations sont développées ou réduites, lorsque l’en-tête de conversation est sélectionné et lorsque les messages de conversation se trouvent dans un dossier différent de celui actuellement affiché.

|Sélection|Vue de conversation développée|Vue de conversation réduite|
|------|------|------|
|**L’en-tête de conversation est sélectionné**|Si l’en-tête de conversation est le seul élément sélectionné, un complément prenant en charge la sélection multiple ne s’active pas. Toutefois, si d’autres messages non-en-tête sont également sélectionnés, le complément s’active uniquement sur ceux-ci et non sur l’en-tête sélectionné.|Le message le plus récent (autrement dit, le premier message dans la pile des conversations) est inclus dans la sélection du message.<br><br>Si le message le plus récent de la conversation se trouve dans un autre dossier de celui actuellement affiché, le message suivant dans la pile située dans le dossier actif est inclus dans la sélection.|
|**Les messages de conversation sélectionnés se trouvent dans le même dossier que celui actuellement affiché**|Tous les messages de conversation choisis sont inclus dans la sélection.|Non applicable Seul l’en-tête de conversation est disponible pour la sélection en mode de conversation réduit.|
|**Les messages de conversation sélectionnés se trouvent dans différents dossiers de celui actuellement affiché** |Tous les messages de conversation choisis sont inclus dans la sélection.|Non applicable Seul l’en-tête de conversation est disponible pour la sélection en mode de conversation réduit.|

## <a name="next-steps"></a>Prochaines étapes

Maintenant que vous avez activé votre complément pour qu’il fonctionne sur plusieurs messages sélectionnés, vous pouvez étendre les fonctionnalités de votre complément et améliorer davantage l’expérience utilisateur. Explorez l’exécution d’opérations plus complexes à l’aide des ID d’élément des messages sélectionnés avec des services tels que [Exchange Web Services (EWS)](web-services.md) et [Microsoft Graph](/graph/overview).

## <a name="see-also"></a>Voir aussi

- [Manifestes de complément Outlook](manifests.md)
- [Appeler des services web à partir d’un complément Outlook](web-services.md)
- [Présentation de Microsoft Graph](/graph/overview)
