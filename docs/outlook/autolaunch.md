---
title: Configurer votre complément Outlook pour l'activation basée sur des événements (prévisualisation)
description: Découvrez comment configurer votre complément Outlook pour l'activation basée sur des événements.
ms.topic: article
ms.date: 04/29/2021
localization_priority: Normal
ms.openlocfilehash: 45f9ff16b3aca0a1fb8f3a8ee3d9ffa8e0f33ea2
ms.sourcegitcommit: 6057afc1776e1667b231d2e9809d261d372151f6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/30/2021
ms.locfileid: "52100298"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="0f346-103">Configurer votre complément Outlook pour l'activation basée sur des événements (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="0f346-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="0f346-104">Sans la fonctionnalité d'activation basée sur des événements, un utilisateur doit lancer explicitement un complément pour effectuer ses tâches.</span><span class="sxs-lookup"><span data-stu-id="0f346-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="0f346-105">Cette fonctionnalité permet à votre application d'exécuter des tâches basées sur certains événements, en particulier pour les opérations qui s'appliquent à chaque élément.</span><span class="sxs-lookup"><span data-stu-id="0f346-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="0f346-106">Vous pouvez également intégrer le volet Des tâches et la fonctionnalité sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0f346-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="0f346-107">Pour l'instant, les événements suivants sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="0f346-107">At present, the following events are supported.</span></span>

|<span data-ttu-id="0f346-108">Événement</span><span class="sxs-lookup"><span data-stu-id="0f346-108">Event</span></span>|<span data-ttu-id="0f346-109">Description</span><span class="sxs-lookup"><span data-stu-id="0f346-109">Description</span></span>|
|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="0f346-110">Lors de la composition d'un nouveau message (y compris répondre, répondre à tous et transmettre), mais pas lors de la modification, par exemple, d'un brouillon.</span><span class="sxs-lookup"><span data-stu-id="0f346-110">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="0f346-111">Lors de la création d'un rendez-vous, mais pas de la modification d'un rendez-vous existant.</span><span class="sxs-lookup"><span data-stu-id="0f346-111">On creating a new appointment but not on editing an existing one.</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="0f346-112">Lors de l'ajout ou de la suppression de pièces jointes lors de la composition d'un message.</span><span class="sxs-lookup"><span data-stu-id="0f346-112">On adding or removing attachments while composing a message.</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="0f346-113">Lors de l'ajout ou de la suppression de pièces jointes lors de la composition d'un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0f346-113">On adding or removing attachments while composing an appointment.</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="0f346-114">Lors de l'ajout ou de la suppression de destinataires lors de la composition d'un message.</span><span class="sxs-lookup"><span data-stu-id="0f346-114">On adding or removing recipients while composing a message.</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="0f346-115">Lors de l'ajout ou de la suppression de participants lors de la composition d'un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0f346-115">On adding or removing attendees while composing an appointment.</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="0f346-116">Lors de la modification de la date et de l'heure lors de la composition d'un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0f346-116">On changing date/time while composing an appointment.</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="0f346-117">Lors de l'ajout, de la modification ou de la suppression des détails de la récurrence lors de la composition d'un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0f346-117">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="0f346-118">Si la date/l'heure est modifiée, `OnAppointmentTimeChanged` l'événement est également déclenché.</span><span class="sxs-lookup"><span data-stu-id="0f346-118">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="0f346-119">Lors du rejet d'une notification lors de la composition d'un élément de message ou de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0f346-119">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="0f346-120">Seul le add-in qui a ajouté la notification sera averti.</span><span class="sxs-lookup"><span data-stu-id="0f346-120">Only the add-in that added the notification will be notified.</span></span>|

<span data-ttu-id="0f346-121">À la fin de cette walkthrough, vous aurez un add-in qui s'exécute chaque fois qu'un nouvel élément est créé et définit l'objet.</span><span class="sxs-lookup"><span data-stu-id="0f346-121">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0f346-122">Cette fonctionnalité est uniquement prise en charge pour la [prévisualisation](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) dans Outlook sur le web et sur Windows avec un Microsoft 365 abonnement.</span><span class="sxs-lookup"><span data-stu-id="0f346-122">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="0f346-123">Pour [plus d'informations,](#how-to-preview-the-event-based-activation-feature) voir comment afficher un aperçu de la fonctionnalité d'activation basée sur des événements dans cet article.</span><span class="sxs-lookup"><span data-stu-id="0f346-123">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="0f346-124">Étant donné que les fonctionnalités d'aperçu sont sujettes à modification sans préavis, elles ne doivent pas être utilisées dans les modules de production.</span><span class="sxs-lookup"><span data-stu-id="0f346-124">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="0f346-125">Comment afficher un aperçu de la fonctionnalité d'activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="0f346-125">How to preview the event-based activation feature</span></span>

<span data-ttu-id="0f346-126">Nous vous invitons à tester la fonctionnalité d'activation basée sur des événements !</span><span class="sxs-lookup"><span data-stu-id="0f346-126">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="0f346-127">Faites-nous part de vos scénarios et de la façon dont nous pouvons les améliorer en nous faisant part de vos commentaires GitHub (voir la **section** Commentaires à la fin de cette page).</span><span class="sxs-lookup"><span data-stu-id="0f346-127">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="0f346-128">Pour afficher un aperçu de cette fonctionnalité :</span><span class="sxs-lookup"><span data-stu-id="0f346-128">To preview this feature:</span></span>

- <span data-ttu-id="0f346-129">Pour Outlook sur le web :</span><span class="sxs-lookup"><span data-stu-id="0f346-129">For Outlook on the web:</span></span>
  - <span data-ttu-id="0f346-130">[Configurez la version ciblée sur votre Microsoft 365 client.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="0f346-130">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="0f346-131">Référencez **la bibliothèque** bêta sur le CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="0f346-131">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="0f346-132">Le [fichier de définition de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) type pour la compilation et la IntelliSense TypeScript se trouve aux CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="0f346-132">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="0f346-133">Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="0f346-133">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="0f346-134">Pour Outlook sur Windows : la build minimale requise est 16.0.13729.20000.</span><span class="sxs-lookup"><span data-stu-id="0f346-134">For Outlook on Windows: The minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="0f346-135">Rejoignez le [Office Insider pour](https://insider.office.com) accéder à Office versions bêta.</span><span class="sxs-lookup"><span data-stu-id="0f346-135">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="0f346-136">Configuration de votre environnement</span><span class="sxs-lookup"><span data-stu-id="0f346-136">Set up your environment</span></span>

<span data-ttu-id="0f346-137">[Complétez Outlook démarrage](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rapide qui crée un projet de compl?ment avec le générateur Yeoman pour Office compl?ments.</span><span class="sxs-lookup"><span data-stu-id="0f346-137">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="0f346-138">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="0f346-138">Configure the manifest</span></span>

<span data-ttu-id="0f346-139">Pour activer l'activation basée sur des événements de votre complément, vous devez configurer l'élément [Runtimes](../reference/manifest/runtimes.md) et le point d'extension [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) dans le nœud `VersionOverridesV1_1` du manifeste.</span><span class="sxs-lookup"><span data-stu-id="0f346-139">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="0f346-140">Pour l'instant, `DesktopFormFactor` est le seul facteur de forme pris en charge.</span><span class="sxs-lookup"><span data-stu-id="0f346-140">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="0f346-141">Dans votre éditeur de code, ouvrez le projet de démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="0f346-141">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="0f346-142">Ouvrez **lemanifest.xml** situé à la racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="0f346-142">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="0f346-143">Sélectionnez l'intégralité du nœud (y compris les balises d'ouverture et de fermeture) et remplacez-le par le `<VersionOverrides>` code XML suivant, puis enregistrez vos modifications.</span><span class="sxs-lookup"><span data-stu-id="0f346-143">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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
        <!-- Entry needed for Outlook Desktop. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
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

<span data-ttu-id="0f346-144">Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web utilise un fichier HTML qui peut référencer le même fichier JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0f346-144">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="0f346-145">Vous devez fournir des références à ces deux fichiers dans le nœud du manifeste, car la plateforme Outlook détermine en fin de compte s'il faut utiliser du code HTML ou JavaScript en fonction du `Resources` client Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f346-145">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="0f346-146">En tant que tel, pour configurer la gestion des événements, fournissez l'emplacement du code HTML dans l'élément, puis, dans son élément enfant, fournissez l'emplacement du fichier JavaScript indiqué ou référencé par le `Runtime` `Override` code HTML.</span><span class="sxs-lookup"><span data-stu-id="0f346-146">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="0f346-147">Pour en savoir plus sur les manifestes de Outlook des Outlook, consultez la Outlook [des manifestes de ces derniers.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="0f346-147">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="0f346-148">Implémenter la gestion des événements</span><span class="sxs-lookup"><span data-stu-id="0f346-148">Implement event handling</span></span>

<span data-ttu-id="0f346-149">Vous devez implémenter la gestion de vos événements sélectionnés.</span><span class="sxs-lookup"><span data-stu-id="0f346-149">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="0f346-150">Dans ce scénario, vous allez ajouter la gestion de la composition de nouveaux éléments.</span><span class="sxs-lookup"><span data-stu-id="0f346-150">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="0f346-151">À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="0f346-151">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="0f346-152">Après la `action` fonction, insérez les fonctions JavaScript suivantes.</span><span class="sxs-lookup"><span data-stu-id="0f346-152">After the `action` function, insert the following JavaScript functions.</span></span>

    ```js
    function onMessageComposeHandler(event) {
      setSubject(event);
    }
    function onAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext" : event
        },
        function (asyncResult) {
          // Handle success or error.
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
          }
    
          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
        });
    }
    ```

1. <span data-ttu-id="0f346-153">Pour que les fonctions fonctionnent dans Outlook sur le **web** avec ce projet généré par le générateur Yeoman pour les applications Office, ajoutez les instructions suivantes à la fin du fichier.</span><span class="sxs-lookup"><span data-stu-id="0f346-153">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="0f346-154">Pour que les fonctions fonctionnent dans **Outlook sur Windows**, ajoutez le code JavaScript suivant à la fin du fichier.</span><span class="sxs-lookup"><span data-stu-id="0f346-154">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="0f346-155">**Remarque**: la vérification `Office.actions` permet de s'assurer Outlook sur le web ignore ces instructions.</span><span class="sxs-lookup"><span data-stu-id="0f346-155">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

1. <span data-ttu-id="0f346-156">Enregistrez vos modifications.</span><span class="sxs-lookup"><span data-stu-id="0f346-156">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="0f346-157">Try it out</span><span class="sxs-lookup"><span data-stu-id="0f346-157">Try it out</span></span>

1. <span data-ttu-id="0f346-158">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="0f346-158">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="0f346-159">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution) et votre complément est chargé.</span><span class="sxs-lookup"><span data-stu-id="0f346-159">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="0f346-160">Dans Outlook sur le web, créez un message.</span><span class="sxs-lookup"><span data-stu-id="0f346-160">In Outlook on the web, create a new message.</span></span>

    ![Capture d'écran d'une fenêtre de message Outlook sur le web avec l'objet de la composition](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="0f346-162">Dans Outlook sur Windows, créez un message.</span><span class="sxs-lookup"><span data-stu-id="0f346-162">In Outlook on Windows, create a new message.</span></span>

    ![Capture d'écran d'une fenêtre de message Outlook sur Windows avec l'objet de la composition](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="0f346-164">Si l'erreur « Nous ne pouvons pas ouvrir ce module à partir de localhost » s'est produite, vous devez activer une exemption de bouclisation.</span><span class="sxs-lookup"><span data-stu-id="0f346-164">If you see the error "We can't open this add-in from localhost," you'll need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="0f346-165">Fermez Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f346-165">Close Outlook.</span></span>
    > 2. <span data-ttu-id="0f346-166">Ouvrez **le Gestionnaire des tâches** et assurez-vous que le processus **msoadfs.exe** n'est pas en cours d'exécution.</span><span class="sxs-lookup"><span data-stu-id="0f346-166">Open the **Task Manager** and ensure that the **msoadfs.exe** process is not running.</span></span>
    > 3. <span data-ttu-id="0f346-167">Exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="0f346-167">Run the following command.</span></span>
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. <span data-ttu-id="0f346-168">Redémarrez Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f346-168">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="0f346-169">Debug</span><span class="sxs-lookup"><span data-stu-id="0f346-169">Debug</span></span>

<span data-ttu-id="0f346-170">Lorsque vous implémentez vos propres fonctionnalités, vous devrez peut-être déboguer votre code.</span><span class="sxs-lookup"><span data-stu-id="0f346-170">As you implement your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="0f346-171">Pour obtenir des instructions sur le débogage de l'activation de complément basée sur des événements, voir [Déboguer](debug-autolaunch.md)votre complément basé sur Outlook événement.</span><span class="sxs-lookup"><span data-stu-id="0f346-171">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="0f346-172">Comportement et limitations de l'activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="0f346-172">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="0f346-173">Les add-ins qui s'activent en fonction des événements sont censés être de courte durée, légers et aussi légers que possible.</span><span class="sxs-lookup"><span data-stu-id="0f346-173">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="0f346-174">Pour signaler que votre add-in a terminé le traitement de l'événement de lancement, nous vous recommandons de demander à votre module d'appeler la `event.completed` méthode.</span><span class="sxs-lookup"><span data-stu-id="0f346-174">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="0f346-175">Si cet appel n'est pas effectué, le délai d'un délai d'environ 300 secondes s'élève à environ 300 secondes, la durée maximale autorisée pour l'exécution de ces derniers. Le add-in se termine également lorsque l'utilisateur ferme la fenêtre de composition.</span><span class="sxs-lookup"><span data-stu-id="0f346-175">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="0f346-176">Si l'utilisateur a plusieurs add-ins abonnés au même événement, la plateforme Outlook lance les modules dans un ordre particulier.</span><span class="sxs-lookup"><span data-stu-id="0f346-176">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="0f346-177">Actuellement, seuls cinq add-ins basés sur des événements peuvent être activement en cours d'exécution.</span><span class="sxs-lookup"><span data-stu-id="0f346-177">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="0f346-178">Tous les compléments supplémentaires sont dirigés vers une file d'attente, puis exécutés à mesure que les compléments précédemment actifs sont terminés ou désactivés.</span><span class="sxs-lookup"><span data-stu-id="0f346-178">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="0f346-179">L'utilisateur peut basculer ou naviguer à partir de l'élément de messagerie actuel où le module a commencé à s'exécute.</span><span class="sxs-lookup"><span data-stu-id="0f346-179">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="0f346-180">Le module qui a été lancé terminera son opération en arrière-plan.</span><span class="sxs-lookup"><span data-stu-id="0f346-180">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="0f346-181">Certaines Office.js API qui modifient ou modifient l'interface utilisateur ne sont pas autorisées à partir des add-ins basés sur des événements. Les API bloquées sont les suivantes :</span><span class="sxs-lookup"><span data-stu-id="0f346-181">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="0f346-182">Sous `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="0f346-182">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="0f346-183">Sous `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="0f346-183">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="0f346-184">Sous `Office.context.mailbox.item` :</span><span class="sxs-lookup"><span data-stu-id="0f346-184">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="0f346-185">Sous `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="0f346-185">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="0f346-186">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0f346-186">See also</span></span>

<span data-ttu-id="0f346-187">[Outlook manifestes de la Outlook de l'équipe](manifests.md) 
 [Comment déboguer des add-ins basés sur des événements](debug-autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="0f346-187">[Outlook add-in manifests](manifests.md)
[How to debug event-based add-ins](debug-autolaunch.md)</span></span>
