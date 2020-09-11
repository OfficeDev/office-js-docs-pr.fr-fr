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
# <a name="implement-append-on-send-in-your-outlook-add-in-preview"></a><span data-ttu-id="3207e-103">Implémenter Append-on-Send dans votre complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="3207e-103">Implement append-on-send in your Outlook add-in (preview)</span></span>

<span data-ttu-id="3207e-104">À la fin de cette procédure pas à pas, vous disposez d’un complément Outlook qui peut insérer une clause d’exclusion de responsabilité lors de l’envoi d’un message.</span><span class="sxs-lookup"><span data-stu-id="3207e-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3207e-105">Cette fonctionnalité est actuellement [prise en charge pour la](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion dans Outlook sur le Web et Windows avec un abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="3207e-105">This feature is currently supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="3207e-106">Pour plus d’informations, reportez-vous [à la rubrique relative à l’aperçu de la fonctionnalité Ajout à l’envoi](#how-to-preview-the-append-on-send-feature) de cet article.</span><span class="sxs-lookup"><span data-stu-id="3207e-106">See [How to preview the append-on-send feature](#how-to-preview-the-append-on-send-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="3207e-107">Les fonctionnalités d’aperçu étant susceptibles d’être modifiées sans préavis, elles ne doivent pas être utilisées dans les compléments de production.</span><span class="sxs-lookup"><span data-stu-id="3207e-107">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-append-on-send-feature"></a><span data-ttu-id="3207e-108">Comment afficher un aperçu de la fonctionnalité Ajouter-on-Send</span><span class="sxs-lookup"><span data-stu-id="3207e-108">How to preview the append-on-send feature</span></span>

<span data-ttu-id="3207e-109">Nous vous invitons à tester la fonctionnalité Ajout à l’envoi !</span><span class="sxs-lookup"><span data-stu-id="3207e-109">We invite you to try out the append-on-send feature!</span></span> <span data-ttu-id="3207e-110">Faites-nous part de vos scénarios et de vos possibilités d’amélioration en nous donnant des commentaires via GitHub (voir la section **Commentaires** à la fin de cette page).</span><span class="sxs-lookup"><span data-stu-id="3207e-110">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="3207e-111">Pour afficher un aperçu de cette fonctionnalité :</span><span class="sxs-lookup"><span data-stu-id="3207e-111">To preview this feature:</span></span>

- <span data-ttu-id="3207e-112">Faites référence à la bibliothèque **beta** sur le CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="3207e-112">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="3207e-113">Le [fichier de définition de type](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) pour la compilation de la machine à écrire et IntelliSense se trouve dans le CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="3207e-113">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="3207e-114">Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="3207e-114">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="3207e-115">Pour Windows, vous devrez peut-être rejoindre le [programme Office Insider](https://insider.office.com) pour accéder à des builds Office plus récentes.</span><span class="sxs-lookup"><span data-stu-id="3207e-115">For Windows, you may need to join the [Office Insider program](https://insider.office.com) to access more recent Office builds.</span></span>
- <span data-ttu-id="3207e-116">Pour Outlook sur le Web, [configurez la version ciblée sur votre client Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="3207e-116">For Outlook on the web, [configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="3207e-117">Configuration de votre environnement</span><span class="sxs-lookup"><span data-stu-id="3207e-117">Set up your environment</span></span>

<span data-ttu-id="3207e-118">Terminez le [démarrage rapide Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le générateur Yeoman pour les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="3207e-118">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="3207e-119">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="3207e-119">Configure the manifest</span></span>

<span data-ttu-id="3207e-120">Pour activer la fonctionnalité Ajout à l’envoi dans votre complément, vous devez inclure l' `AppendOnSend` autorisation dans la collection de [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="3207e-120">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="3207e-121">Pour ce scénario, au lieu d’exécuter la `action` fonction en cliquant sur le bouton **effectuer une action** , vous exécuterez `appendOnSend` la fonction.</span><span class="sxs-lookup"><span data-stu-id="3207e-121">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="3207e-122">Dans votre éditeur de code, ouvrez le projet Quick Start.</span><span class="sxs-lookup"><span data-stu-id="3207e-122">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="3207e-123">Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="3207e-123">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="3207e-124">Sélectionnez le `<VersionOverrides>` nœud entier (y compris les balises ouvrantes et fermantes) et remplacez-le par le code XML suivant.</span><span class="sxs-lookup"><span data-stu-id="3207e-124">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="3207e-125">Pour en savoir plus sur les manifestes pour les compléments Outlook, consultez la rubrique [manifestes des compléments Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="3207e-125">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="3207e-126">Implémenter la gestion des ajouts à l’envoi</span><span class="sxs-lookup"><span data-stu-id="3207e-126">Implement append-on-send handling</span></span>

<span data-ttu-id="3207e-127">Ensuite, implémentez l’ajout sur l’événement Send.</span><span class="sxs-lookup"><span data-stu-id="3207e-127">Next, implement appending on the send event.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3207e-128">Si votre complément implémente également la [gestion des événements d’envoi à l' `ItemSend` aide ](outlook-on-send-addins.md)de, l’appel `AppendOnSendAsync` dans le gestionnaire d’envoi renvoie une erreur dans la mesure où ce scénario n’est pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="3207e-128">If your add-in also implements [on-send event handling using `ItemSend`](outlook-on-send-addins.md), calling `AppendOnSendAsync` in the on-send handler returns an error as this scenario isn't supported.</span></span>

<span data-ttu-id="3207e-129">Pour ce scénario, vous allez implémenter l’ajout d’une clause d’exclusion de responsabilité à l’élément lorsque l’utilisateur envoie.</span><span class="sxs-lookup"><span data-stu-id="3207e-129">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="3207e-130">À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="3207e-130">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="3207e-131">Après la `action` fonction, insérez la fonction JavaScript suivante.</span><span class="sxs-lookup"><span data-stu-id="3207e-131">After the `action` function, insert the following JavaScript function.</span></span>

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

1. <span data-ttu-id="3207e-132">À la fin du fichier, ajoutez l’instruction suivante.</span><span class="sxs-lookup"><span data-stu-id="3207e-132">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="3207e-133">Essayez</span><span class="sxs-lookup"><span data-stu-id="3207e-133">Try it out</span></span>

1. <span data-ttu-id="3207e-134">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="3207e-134">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="3207e-135">Lorsque vous exécutez cette commande, le serveur Web local démarre s’il n’est pas déjà en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="3207e-135">When you run this command, the local web server will start if it's not already running.</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="3207e-136">Suivez les instructions de [chargement compléments Outlook à des fins de test](sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="3207e-136">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="3207e-137">Créez un message et ajoutez-vous à la ligne **à** .</span><span class="sxs-lookup"><span data-stu-id="3207e-137">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="3207e-138">Dans le menu du ruban ou du buffer overflow, sélectionnez **effectuer une action**.</span><span class="sxs-lookup"><span data-stu-id="3207e-138">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="3207e-139">Envoyez le message, puis ouvrez-le à partir de votre dossier **boîte de réception** ou **éléments envoyés** pour afficher la clause d’exclusion de responsabilité ajoutée.</span><span class="sxs-lookup"><span data-stu-id="3207e-139">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![Capture d’écran d’un exemple de message avec la clause d’exclusion de responsabilité ajoutée lors de l’envoi dans Outlook sur le Web.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="3207e-141">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3207e-141">See also</span></span>

[<span data-ttu-id="3207e-142">Manifestes de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="3207e-142">Outlook add-in manifests</span></span>](manifests.md)
