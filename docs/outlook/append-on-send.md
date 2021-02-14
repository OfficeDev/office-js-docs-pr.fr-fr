---
title: Implémenter l’ajout à l’envoi dans votre application Outlook
description: Découvrez comment implémenter la fonctionnalité d’ajout à l’envoi dans votre application Outlook.
ms.topic: article
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 8b69fbbaef1d0f060f0675fe5c4948a70d935b7a
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234288"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a><span data-ttu-id="6d3b2-103">Implémenter l’ajout à l’envoi dans votre application Outlook</span><span class="sxs-lookup"><span data-stu-id="6d3b2-103">Implement append-on-send in your Outlook add-in</span></span>

<span data-ttu-id="6d3b2-104">À la fin de cette walkthrough, vous aurez un add-in Outlook qui peut insérer une clause d’exclusion de responsabilité lorsqu’un message est envoyé.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!NOTE]
> <span data-ttu-id="6d3b2-105">La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.9.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-105">Support for this feature was introduced in requirement set 1.9.</span></span> <span data-ttu-id="6d3b2-106">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-106">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="6d3b2-107">Configuration de votre environnement</span><span class="sxs-lookup"><span data-stu-id="6d3b2-107">Set up your environment</span></span>

<span data-ttu-id="6d3b2-108">Terminez [le démarrage rapide d’Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de compl?ment avec le générateur Yeoman pour les compl?ments Office.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-108">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="6d3b2-109">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="6d3b2-109">Configure the manifest</span></span>

<span data-ttu-id="6d3b2-110">Pour activer la fonctionnalité d’ajout à l’envoi dans votre add-in, vous devez inclure l’autorisation dans la `AppendOnSend` collection [de ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="6d3b2-110">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="6d3b2-111">Pour ce scénario, au lieu d’exécuter la fonction sur le bouton Effectuer une action, vous exécuterez `action` la  `appendOnSend` fonction.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-111">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="6d3b2-112">Dans votre éditeur de code, ouvrez le projet de démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-112">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="6d3b2-113">Ouvrez **lemanifest.xml** situé à la racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-113">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="6d3b2-114">Sélectionnez l’intégralité du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le `<VersionOverrides>` par le code XML suivant.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-114">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="6d3b2-115">Pour en savoir plus sur les manifestes pour les add-ins Outlook, consultez les [manifestes de ces derniers.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="6d3b2-115">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="6d3b2-116">Implémenter la gestion de l’envoi</span><span class="sxs-lookup"><span data-stu-id="6d3b2-116">Implement append-on-send handling</span></span>

<span data-ttu-id="6d3b2-117">Ensuite, implémentez l’application sur l’événement d’envoi.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-117">Next, implement appending on the send event.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6d3b2-118">Si votre application implémente également la gestion des [événements `ItemSend` ](outlook-on-send-addins.md)d’envoi à l’aide de , l’appel dans le handler d’envoi renvoie une erreur, car ce `AppendOnSendAsync` scénario n’est pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-118">If your add-in also implements [on-send event handling using `ItemSend`](outlook-on-send-addins.md), calling `AppendOnSendAsync` in the on-send handler returns an error as this scenario isn't supported.</span></span>

<span data-ttu-id="6d3b2-119">Pour ce scénario, vous allez implémenter l’application d’une clause d’exclusion de responsabilité à l’élément lorsque l’utilisateur l’envoie.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-119">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="6d3b2-120">À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-120">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="6d3b2-121">Après la `action` fonction, insérez la fonction JavaScript suivante.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-121">After the `action` function, insert the following JavaScript function.</span></span>

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

1. <span data-ttu-id="6d3b2-122">À la fin du fichier, ajoutez l’instruction suivante.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-122">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="6d3b2-123">Try it out</span><span class="sxs-lookup"><span data-stu-id="6d3b2-123">Try it out</span></span>

1. <span data-ttu-id="6d3b2-124">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-124">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="6d3b2-125">Lorsque vous exécutez cette commande, le serveur web local démarre s’il n’est pas déjà en cours d’exécution et que votre application est rechargée de nouveau.</span><span class="sxs-lookup"><span data-stu-id="6d3b2-125">When you run this command, the local web server will start if it's not already running and your add-in will be sideloaded.</span></span> 

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="6d3b2-126">Créez un message et ajoutez-vous à la **ligne À.**</span><span class="sxs-lookup"><span data-stu-id="6d3b2-126">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="6d3b2-127">Dans le ruban ou le menu de dépassement, choisissez **Effectuer une action.**</span><span class="sxs-lookup"><span data-stu-id="6d3b2-127">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="6d3b2-128">Envoyez le message, puis  ouvrez-le à partir de votre boîte de réception ou dossier Éléments envoyés pour afficher la clause d’exclusion de responsabilité. </span><span class="sxs-lookup"><span data-stu-id="6d3b2-128">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![Capture d’écran d’un exemple de message avec la clause d’exclusion de responsabilité à l’envoi dans Outlook sur le web.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="6d3b2-130">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6d3b2-130">See also</span></span>

[<span data-ttu-id="6d3b2-131">Manifestes de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="6d3b2-131">Outlook add-in manifests</span></span>](manifests.md)
