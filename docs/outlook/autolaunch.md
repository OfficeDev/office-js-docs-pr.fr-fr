---
title: Configurer votre complément Outlook pour l’activation basée sur les événements (aperçu)
description: Découvrez comment configurer votre complément Outlook pour l’activation basée sur les événements.
ms.topic: article
ms.date: 05/22/2020
localization_priority: Normal
ms.openlocfilehash: 7b14dfb425223dbf486173e7df103d1a2cd09641
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093426"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="794c1-103">Configurer votre complément Outlook pour l’activation basée sur les événements (aperçu)</span><span class="sxs-lookup"><span data-stu-id="794c1-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="794c1-104">Sans la fonctionnalité d’activation basée sur un événement, un utilisateur doit lancer explicitement un complément pour effectuer ses tâches.</span><span class="sxs-lookup"><span data-stu-id="794c1-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="794c1-105">Cette fonctionnalité permet à votre complément d’exécuter des tâches en fonction de certains événements, en particulier pour les opérations qui s’appliquent à chaque élément.</span><span class="sxs-lookup"><span data-stu-id="794c1-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="794c1-106">Vous pouvez également intégrer le volet des tâches et les fonctionnalités sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="794c1-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="794c1-107">Pour le moment, les événements pris en charge sont les suivants.</span><span class="sxs-lookup"><span data-stu-id="794c1-107">At present, the supported events are as follows.</span></span>

- <span data-ttu-id="794c1-108">`OnNewMessageCompose`: Lors de la composition d’un nouveau message (inclut répondre, répondre à tous et transférer)</span><span class="sxs-lookup"><span data-stu-id="794c1-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="794c1-109">`OnNewAppointmentOrganizer`: Lors de la création d’un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="794c1-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

<span data-ttu-id="794c1-110">À la fin de cette procédure pas à pas, vous disposez d’un complément qui s’exécute chaque fois qu’un nouveau message est créé.</span><span class="sxs-lookup"><span data-stu-id="794c1-110">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="794c1-111">Cette fonctionnalité est uniquement prise en charge pour l' [Aperçu](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) dans Outlook sur le Web avec un abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="794c1-111">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span> <span data-ttu-id="794c1-112">Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](#how-to-preview-the-event-based-activation-feature) dans cet article.</span><span class="sxs-lookup"><span data-stu-id="794c1-112">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="794c1-113">Les fonctionnalités d’aperçu étant susceptibles d’être modifiées sans préavis, elles ne doivent pas être utilisées dans les compléments de production.</span><span class="sxs-lookup"><span data-stu-id="794c1-113">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="794c1-114">Comment afficher un aperçu de la fonctionnalité d’activation basée sur un événement</span><span class="sxs-lookup"><span data-stu-id="794c1-114">How to preview the event-based activation feature</span></span>

<span data-ttu-id="794c1-115">Nous vous invitons à tester la fonctionnalité d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="794c1-115">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="794c1-116">Faites-nous part de vos scénarios et de vos possibilités d’amélioration en nous donnant des commentaires via GitHub (voir la section **Commentaires** à la fin de cette page).</span><span class="sxs-lookup"><span data-stu-id="794c1-116">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="794c1-117">Pour afficher un aperçu de cette fonctionnalité :</span><span class="sxs-lookup"><span data-stu-id="794c1-117">To preview this feature:</span></span>

- <span data-ttu-id="794c1-118">Faites référence à la bibliothèque **beta** sur le CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="794c1-118">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="794c1-119">Le [fichier de définition de type](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) pour la compilation de la machine à écrire et IntelliSense se trouve dans le CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="794c1-119">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="794c1-120">Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="794c1-120">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="794c1-121">Demander l’accès à des bits d’aperçu pour Outlook sur le Web à l’aide de votre compte Microsoft 365 en remplissant et envoyant [ce formulaire de demande](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="794c1-121">Request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this request form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="794c1-122">Nous allons vous indiquer quand votre client est prêt.</span><span class="sxs-lookup"><span data-stu-id="794c1-122">We'll let you know when your tenant is ready.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="794c1-123">Configuration de votre environnement</span><span class="sxs-lookup"><span data-stu-id="794c1-123">Set up your environment</span></span>

<span data-ttu-id="794c1-124">Terminez le [démarrage rapide Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet de complément avec le générateur Yeoman pour les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="794c1-124">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="794c1-125">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="794c1-125">Configure the manifest</span></span>

<span data-ttu-id="794c1-126">Pour activer l’activation basée sur les événements de votre complément, vous devez configurer l’élément [runtimes](../reference/manifest/runtimes.md) et le point d’extension [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="794c1-126">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest.</span></span> <span data-ttu-id="794c1-127">Pour le moment, `DesktopFormFactor` est le seul facteur de forme pris en charge.</span><span class="sxs-lookup"><span data-stu-id="794c1-127">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="794c1-128">Dans votre éditeur de code, ouvrez le projet Quick Start.</span><span class="sxs-lookup"><span data-stu-id="794c1-128">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="794c1-129">Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="794c1-129">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="794c1-130">Sélectionnez le `<VersionOverrides>` nœud entier (y compris les balises ouvrantes et fermantes) et remplacez-le par le code XML suivant.</span><span class="sxs-lookup"><span data-stu-id="794c1-130">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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

<span data-ttu-id="794c1-131">Outlook sur Windows utilise un fichier JavaScript, tandis qu’Outlook sur le Web utilise un fichier HTML qui fait référence au même fichier JavaScript.</span><span class="sxs-lookup"><span data-stu-id="794c1-131">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file.</span></span> <span data-ttu-id="794c1-132">Vous devez fournir des références à ces deux fichiers dans le manifeste, car la plateforme Outlook détermine finalement s’il faut utiliser du code HTML ou JavaScript basé sur le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="794c1-132">You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="794c1-133">Par exemple, pour configurer la gestion des événements, indiquez l’emplacement du code HTML dans l' `Runtime` élément, puis dans son `Override` élément enfant, indiquez l’emplacement du fichier JavaScript inline ou référencé par le code html.</span><span class="sxs-lookup"><span data-stu-id="794c1-133">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="794c1-134">Pour en savoir plus sur les manifestes pour les compléments Outlook, consultez la rubrique [manifestes des compléments Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="794c1-134">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="794c1-135">Implémenter la gestion des événements</span><span class="sxs-lookup"><span data-stu-id="794c1-135">Implement event handling</span></span>

<span data-ttu-id="794c1-136">Vous devez implémenter la gestion de vos événements sélectionnés.</span><span class="sxs-lookup"><span data-stu-id="794c1-136">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="794c1-137">Dans ce scénario, vous allez ajouter la gestion de la composition de nouveaux éléments.</span><span class="sxs-lookup"><span data-stu-id="794c1-137">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="794c1-138">À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="794c1-138">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="794c1-139">Après la `action` fonction, insérez les fonctions JavaScript suivantes.</span><span class="sxs-lookup"><span data-stu-id="794c1-139">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="794c1-140">À la fin du fichier, ajoutez les instructions suivantes.</span><span class="sxs-lookup"><span data-stu-id="794c1-140">At the end of the file, add the following statements.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a><span data-ttu-id="794c1-141">Try it out</span><span class="sxs-lookup"><span data-stu-id="794c1-141">Try it out</span></span>

1. <span data-ttu-id="794c1-142">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="794c1-142">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="794c1-143">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="794c1-143">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="794c1-144">Suivez les instructions indiquées dans l’article [Chargement de version test des compléments Outlook](sideload-outlook-add-ins-for-testing.md) pour charger le complément dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="794c1-144">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="794c1-145">Dans Outlook sur le web, créez un message.</span><span class="sxs-lookup"><span data-stu-id="794c1-145">In Outlook on the web, create a new message.</span></span>

    ![Capture d’écran d’une fenêtre de message dans Outlook sur le Web avec l’objet défini sur la composition.](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="794c1-147">Comportement et limitations de l’activation basée sur les événements</span><span class="sxs-lookup"><span data-stu-id="794c1-147">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="794c1-148">Les compléments qui s’activent sur la base des événements sont conçus pour une exécution à court terme, jusqu’à 330 secondes seulement.</span><span class="sxs-lookup"><span data-stu-id="794c1-148">Add-ins that activate based on events are designed to be short-running, up to 330 seconds only.</span></span> <span data-ttu-id="794c1-149">Nous vous recommandons de faire en sorte que votre complément appelle la `event.completed` méthode pour signaler qu’il a terminé le traitement de l’événement Launch.</span><span class="sxs-lookup"><span data-stu-id="794c1-149">We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event.</span></span> <span data-ttu-id="794c1-150">Le complément se termine également lorsque l’utilisateur ferme la fenêtre de composition.</span><span class="sxs-lookup"><span data-stu-id="794c1-150">The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="794c1-151">Si l’utilisateur a plusieurs compléments qui s’abonnent au même événement, la plateforme Outlook lance les compléments sans ordre particulier.</span><span class="sxs-lookup"><span data-stu-id="794c1-151">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="794c1-152">Actuellement, seuls cinq compléments basés sur les événements peuvent être exécutés activement.</span><span class="sxs-lookup"><span data-stu-id="794c1-152">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="794c1-153">Tout complément supplémentaire est placé dans une file d’attente, puis exécuté comme les compléments précédemment actifs sont terminés ou désactivés.</span><span class="sxs-lookup"><span data-stu-id="794c1-153">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="794c1-154">L’utilisateur peut basculer ou naviguer hors de l’élément de courrier actuel dans lequel le complément a commencé.</span><span class="sxs-lookup"><span data-stu-id="794c1-154">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="794c1-155">Le complément qui a été lancé terminera son opération en arrière-plan.</span><span class="sxs-lookup"><span data-stu-id="794c1-155">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="794c1-156">Certaines API de Office.js qui modifient ou modifient l’interface utilisateur ne sont pas autorisées dans les compléments basés sur des événements. Les API bloquées sont les suivantes.</span><span class="sxs-lookup"><span data-stu-id="794c1-156">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="794c1-157">Sous `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="794c1-157">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="794c1-158">Sous `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="794c1-158">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="794c1-159">Sous `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="794c1-159">Under `Office.context.auth`:</span></span>
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="794c1-160">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="794c1-160">See also</span></span>

[<span data-ttu-id="794c1-161">Manifestes de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="794c1-161">Outlook add-in manifests</span></span>](manifests.md)
