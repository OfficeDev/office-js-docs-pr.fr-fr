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
# <a name="implement-append-on-send-in-your-outlook-add-in"></a><span data-ttu-id="e4e8b-103">Implémenter Append-on-Send dans votre complément Outlook</span><span class="sxs-lookup"><span data-stu-id="e4e8b-103">Implement append-on-send in your Outlook add-in</span></span>

<span data-ttu-id="e4e8b-104">À la fin de cette procédure pas à pas, vous disposez d’un complément Outlook qui peut insérer une clause d’exclusion de responsabilité lors de l’envoi d’un message.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!NOTE]
> <span data-ttu-id="e4e8b-105">La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1,9.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-105">Support for this feature was introduced in requirement set 1.9.</span></span> <span data-ttu-id="e4e8b-106">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-106">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="e4e8b-107">Configuration de votre environnement</span><span class="sxs-lookup"><span data-stu-id="e4e8b-107">Set up your environment</span></span>

<span data-ttu-id="e4e8b-108">Terminez le [démarrage rapide Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le générateur Yeoman pour les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-108">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="e4e8b-109">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="e4e8b-109">Configure the manifest</span></span>

<span data-ttu-id="e4e8b-110">Pour activer la fonctionnalité Ajout à l’envoi dans votre complément, vous devez inclure l' `AppendOnSend` autorisation dans la collection de [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="e4e8b-110">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="e4e8b-111">Pour ce scénario, au lieu d’exécuter la `action` fonction en cliquant sur le bouton **effectuer une action** , vous exécuterez `appendOnSend` la fonction.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-111">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="e4e8b-112">Dans votre éditeur de code, ouvrez le projet Quick Start.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-112">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="e4e8b-113">Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-113">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="e4e8b-114">Sélectionnez le `<VersionOverrides>` nœud entier (y compris les balises ouvrantes et fermantes) et remplacez-le par le code XML suivant.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-114">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="e4e8b-115">Pour en savoir plus sur les manifestes pour les compléments Outlook, consultez la rubrique [manifestes des compléments Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="e4e8b-115">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="e4e8b-116">Implémenter la gestion des ajouts à l’envoi</span><span class="sxs-lookup"><span data-stu-id="e4e8b-116">Implement append-on-send handling</span></span>

<span data-ttu-id="e4e8b-117">Ensuite, implémentez l’ajout sur l’événement Send.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-117">Next, implement appending on the send event.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e4e8b-118">Si votre complément implémente également la [gestion des événements d’envoi à l' `ItemSend` aide ](outlook-on-send-addins.md)de, l’appel `AppendOnSendAsync` dans le gestionnaire d’envoi renvoie une erreur dans la mesure où ce scénario n’est pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-118">If your add-in also implements [on-send event handling using `ItemSend`](outlook-on-send-addins.md), calling `AppendOnSendAsync` in the on-send handler returns an error as this scenario isn't supported.</span></span>

<span data-ttu-id="e4e8b-119">Pour ce scénario, vous allez implémenter l’ajout d’une clause d’exclusion de responsabilité à l’élément lorsque l’utilisateur envoie.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-119">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="e4e8b-120">À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-120">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="e4e8b-121">Après la `action` fonction, insérez la fonction JavaScript suivante.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-121">After the `action` function, insert the following JavaScript function.</span></span>

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

1. <span data-ttu-id="e4e8b-122">À la fin du fichier, ajoutez l’instruction suivante.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-122">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="e4e8b-123">Try it out</span><span class="sxs-lookup"><span data-stu-id="e4e8b-123">Try it out</span></span>

1. <span data-ttu-id="e4e8b-124">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-124">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="e4e8b-125">Lorsque vous exécutez cette commande, le serveur Web local démarre s’il n’est pas déjà en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-125">When you run this command, the local web server will start if it's not already running.</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="e4e8b-126">Suivez les instructions de [chargement compléments Outlook à des fins de test](sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="e4e8b-126">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="e4e8b-127">Créez un message et ajoutez-vous à la ligne **à** .</span><span class="sxs-lookup"><span data-stu-id="e4e8b-127">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="e4e8b-128">Dans le menu du ruban ou du buffer overflow, sélectionnez **effectuer une action**.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-128">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="e4e8b-129">Envoyez le message, puis ouvrez-le à partir de votre dossier **boîte de réception** ou **éléments envoyés** pour afficher la clause d’exclusion de responsabilité ajoutée.</span><span class="sxs-lookup"><span data-stu-id="e4e8b-129">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![Capture d’écran d’un exemple de message avec la clause d’exclusion de responsabilité ajoutée lors de l’envoi dans Outlook sur le Web.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="e4e8b-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e4e8b-131">See also</span></span>

[<span data-ttu-id="e4e8b-132">Manifestes de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="e4e8b-132">Outlook add-in manifests</span></span>](manifests.md)
