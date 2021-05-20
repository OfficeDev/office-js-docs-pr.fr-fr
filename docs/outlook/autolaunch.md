---
title: Configurez votre Outlook add-in pour l’activation basée sur l’événement (aperçu)
description: Découvrez comment configurer vos Outlook pour l’activation basée sur l’événement.
ms.topic: article
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 721f05e1c835e066744598ecb2bd416c6a6b0526
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555244"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="7ff6b-103">Configurez votre Outlook add-in pour l’activation basée sur l’événement (aperçu)</span><span class="sxs-lookup"><span data-stu-id="7ff6b-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="7ff6b-104">Sans la fonction d’activation basée sur l’événement, un utilisateur doit lancer explicitement un module supplémentaire pour accomplir ses tâches.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="7ff6b-105">Cette fonctionnalité permet à votre module d’exécuter des tâches en fonction de certains événements, en particulier pour les opérations qui s’appliquent à chaque élément.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="7ff6b-106">Vous pouvez également intégrer avec le volet de tâche et les fonctionnalités sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="7ff6b-107">À la fin de cette procédure pas à pas, vous aurez un add-in qui s’exécute chaque fois qu’un nouvel élément est créé et définit le sujet.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7ff6b-108">Cette fonctionnalité n’est prise en [charge que](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) pour un aperçu Outlook sur le web et sur Windows avec un abonnement Microsoft 365 spécial.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-108">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="7ff6b-109">Pour plus de détails, voir [Comment prévisualiser la fonction d’activation basée sur l’événement](#how-to-preview-the-event-based-activation-feature) dans cet article.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-109">For more details, see [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article.</span></span>
>
> <span data-ttu-id="7ff6b-110">Étant donné que les fonctionnalités d’aperçu sont sujettes à changement sans préavis, elles ne doivent pas être utilisées dans les modules de production.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-110">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="supported-events"></a><span data-ttu-id="7ff6b-111">Événements pris en charge</span><span class="sxs-lookup"><span data-stu-id="7ff6b-111">Supported events</span></span>

<span data-ttu-id="7ff6b-112">À l’heure actuelle, les événements suivants sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-112">At present, the following events are supported.</span></span>

|<span data-ttu-id="7ff6b-113">Événement</span><span class="sxs-lookup"><span data-stu-id="7ff6b-113">Event</span></span>|<span data-ttu-id="7ff6b-114">Description</span><span class="sxs-lookup"><span data-stu-id="7ff6b-114">Description</span></span>|<span data-ttu-id="7ff6b-115">Clients</span><span class="sxs-lookup"><span data-stu-id="7ff6b-115">Clients</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="7ff6b-116">Sur la composition d’un nouveau message (inclut la réponse, répondre tous, et en avant) mais pas sur l’édition, par exemple, un projet.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-116">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="7ff6b-117">Windows, web</span><span class="sxs-lookup"><span data-stu-id="7ff6b-117">Windows, web</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="7ff6b-118">Sur la création d’un nouveau rendez-vous, mais pas sur l’édition d’un existant.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-118">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="7ff6b-119">Windows, web</span><span class="sxs-lookup"><span data-stu-id="7ff6b-119">Windows, web</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="7ff6b-120">Lors de l’ajout ou de la suppression des pièces jointes lors de la composition d’un message.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-120">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="7ff6b-121">Windows</span><span class="sxs-lookup"><span data-stu-id="7ff6b-121">Windows</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="7ff6b-122">Lors de l’ajout ou de la suppression des pièces jointes lors de la composition d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-122">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="7ff6b-123">Windows</span><span class="sxs-lookup"><span data-stu-id="7ff6b-123">Windows</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="7ff6b-124">Lors de l’ajout ou de la suppression de destinataires lors de la composition d’un message.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-124">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="7ff6b-125">Windows</span><span class="sxs-lookup"><span data-stu-id="7ff6b-125">Windows</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="7ff6b-126">Sur l’ajout ou la suppression des participants lors de la composition d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-126">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="7ff6b-127">Windows</span><span class="sxs-lookup"><span data-stu-id="7ff6b-127">Windows</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="7ff6b-128">À la date/heure changeante tout en composant un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-128">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="7ff6b-129">Windows</span><span class="sxs-lookup"><span data-stu-id="7ff6b-129">Windows</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="7ff6b-130">Lors de l’ajout, de la modification ou de la suppression des détails de récurrence lors de la composition d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-130">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="7ff6b-131">Si la date/heure est modifiée, `OnAppointmentTimeChanged` l’événement sera également déclenché.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-131">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="7ff6b-132">Windows</span><span class="sxs-lookup"><span data-stu-id="7ff6b-132">Windows</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="7ff6b-133">Lors du rejet d’une notification lors de la composition d’un message ou d’un élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-133">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="7ff6b-134">Seul l’add-in qui a ajouté la notification sera notifié.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-134">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="7ff6b-135">Windows</span><span class="sxs-lookup"><span data-stu-id="7ff6b-135">Windows</span></span>|

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="7ff6b-136">Comment prévisualiser la fonction d’activation basée sur l’événement</span><span class="sxs-lookup"><span data-stu-id="7ff6b-136">How to preview the event-based activation feature</span></span>

<span data-ttu-id="7ff6b-137">Nous vous invitons à essayer la fonction d’activation basée sur l’événement!</span><span class="sxs-lookup"><span data-stu-id="7ff6b-137">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="7ff6b-138">Faites-nous part de vos scénarios et de la façon dont nous pouvons nous améliorer en nous donnant des commentaires par GitHub **(voir la** section Commentaires à la fin de cette page).</span><span class="sxs-lookup"><span data-stu-id="7ff6b-138">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="7ff6b-139">Pour prévisualiser cette fonctionnalité :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-139">To preview this feature:</span></span>

- <span data-ttu-id="7ff6b-140">Pour Outlook sur le web :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-140">For Outlook on the web:</span></span>
  - <span data-ttu-id="7ff6b-141">[Configurez la version ciblée sur votre Microsoft 365 locataire](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="7ff6b-141">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="7ff6b-142">Référencez **la bibliothèque** bêta sur le CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="7ff6b-142">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="7ff6b-143">Le [fichier de définition de type](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) pour la compilation typescript IntelliSense est trouvé à la CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="7ff6b-143">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="7ff6b-144">Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="7ff6b-144">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="7ff6b-145">Pour Outlook sur Windows :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-145">For Outlook on Windows:</span></span>
  - <span data-ttu-id="7ff6b-146">La construction minimale requise est de 16.0.14026.20000.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-146">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="7ff6b-147">Rejoignez [le Office Insider pour](https://insider.office.com) accéder aux versions Office bêta.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-147">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="7ff6b-148">Configurez le registre.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-148">Configure the registry.</span></span> <span data-ttu-id="7ff6b-149">Outlook comprend une copie locale des versions bêta et de production des Office.js au lieu de charger à partir du CDN.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-149">Outlook includes a local copy of the production and beta versions of Office.js instead of loading from the CDN.</span></span> <span data-ttu-id="7ff6b-150">Par défaut, la copie de production locale de l’API est référencée.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-150">By default, the local production copy of the API is referenced.</span></span> <span data-ttu-id="7ff6b-151">Pour passer à la copie bêta locale des API JavaScript Outlook, vous devez ajouter cette entrée de registre, sinon les API bêta peuvent ne pas être trouvées.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-151">To switch to the local beta copy of the Outlook JavaScript APIs, you need to add this registry entry, otherwise beta APIs may not be found.</span></span>
    1. <span data-ttu-id="7ff6b-152">Créez la clé du registre `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` .</span><span class="sxs-lookup"><span data-stu-id="7ff6b-152">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="7ff6b-153">Ajouter une entrée nommée `EnableBetaAPIsInJavaScript` et définir la valeur à `1` .</span><span class="sxs-lookup"><span data-stu-id="7ff6b-153">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span> <span data-ttu-id="7ff6b-154">L’image suivante indique à quoi doit ressembler le registre.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-154">The following image shows what the registry should look like.</span></span>

        ![Capture d’écran de l’éditeur du registre avec une valeur clé du registre EnableBetaAPIsInJavaScript](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a><span data-ttu-id="7ff6b-156">Configuration de votre environnement</span><span class="sxs-lookup"><span data-stu-id="7ff6b-156">Set up your environment</span></span>

<span data-ttu-id="7ff6b-157">Complétez [Outlook démarrage rapide](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) qui crée un projet d’ajout avec le générateur Yeoman pour Office add-ins.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-157">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="7ff6b-158">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="7ff6b-158">Configure the manifest</span></span>

<span data-ttu-id="7ff6b-159">Pour activer l’activation basée sur l’événement de votre module, vous devez configurer [l’élément Runtimes](../reference/manifest/runtimes.md) et le point [d’extension LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` dans le nœud du manifeste.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-159">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="7ff6b-160">Pour l’instant, `DesktopFormFactor` est le seul facteur de forme pris en charge.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-160">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="7ff6b-161">Dans votre éditeur de code, ouvrez le projet de démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-161">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="7ff6b-162">Ouvrez **manifest.xml** fichier situé à l’origine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-162">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="7ff6b-163">Sélectionnez `<VersionOverrides>` l’ensemble du nœud (y compris les balises ouvertes et proches) et remplacez-le par le XML suivant, puis enregistrez vos modifications.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-163">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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
              <!-- Events supported on the web and on Windows. -->
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
              <!-- Events supported only on Windows. -->
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
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

<span data-ttu-id="7ff6b-164">Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web utilise un fichier HTML qui peut référencer le même fichier JavaScript.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-164">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="7ff6b-165">Vous devez fournir des références à ces deux fichiers dans `Resources` le nœud du manifeste que la plate-forme Outlook détermine en fin de compte s’il faut utiliser HTML ou JavaScript en fonction de la Outlook client.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-165">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="7ff6b-166">En tant que tel, pour configurer la gestion d’événements, fournir l’emplacement du HTML dans `Runtime` l’élément, puis dans `Override` son élément enfant fournir l’emplacement du fichier JavaScript inlined ou référencé par le HTML.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-166">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="7ff6b-167">Pour en savoir plus sur les manifestes Outlook les add-ins, [consultez Outlook manifestes add-in](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="7ff6b-167">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="7ff6b-168">Implémenter la gestion des événements</span><span class="sxs-lookup"><span data-stu-id="7ff6b-168">Implement event handling</span></span>

<span data-ttu-id="7ff6b-169">Vous devez implémenter la manipulation de vos événements sélectionnés.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-169">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="7ff6b-170">Dans ce scénario, vous ajouterez la manipulation pour composer de nouveaux éléments.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-170">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="7ff6b-171">À partir du même projet de démarrage rapide, ouvrez le **fichier ./src/commandes/commands.jsdans** votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-171">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="7ff6b-172">Après la `action` fonction, insérez les fonctions JavaScript suivantes.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-172">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="7ff6b-173">Ajoutez le code JavaScript suivant à la fin du fichier.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-173">Add the following JavaScript code at the end of the file.</span></span>

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. <span data-ttu-id="7ff6b-174">Enregistrez vos modifications.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-174">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="7ff6b-175">Try it out</span><span class="sxs-lookup"><span data-stu-id="7ff6b-175">Try it out</span></span>

1. <span data-ttu-id="7ff6b-176">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-176">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="7ff6b-177">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution) et votre complément est chargé.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-177">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="7ff6b-178">Si votre module d’ajout n’a pas été automatiquement sideloaded, puis suivez les instructions [dans sideload Outlook add-ins](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) pour les tests pour sideload manuellement l’add-in dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-178">If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="7ff6b-179">Dans Outlook sur le web, créez un message.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-179">In Outlook on the web, create a new message.</span></span>

    ![Capture d’écran d’une fenêtre de message Outlook sur le web avec le sujet mis sur composer](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="7ff6b-181">En Outlook sur Windows, créez un nouveau message.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-181">In Outlook on Windows, create a new message.</span></span>

    ![Capture d’écran d’une fenêtre de message Outlook sur Windows avec le sujet mis sur composer](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="7ff6b-183">Si vous lancez votre add-in depuis localhost et que vous voyez l’erreur « Nous sommes désolés, nous n’avons pas *pu accéder à {your-add-in-name-here}*.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-183">If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*.</span></span> <span data-ttu-id="7ff6b-184">Assurez-vous d’avoir une connexion réseau.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-184">Make sure you have a network connection.</span></span> <span data-ttu-id="7ff6b-185">Si le problème persiste, s’il vous plaît réessayer plus tard. », vous devrez peut-être activer une exemption loopback.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-185">If the problem continues, please try again later.", you may need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="7ff6b-186">Fermez Outlook.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-186">Close Outlook.</span></span>
    > 1. <span data-ttu-id="7ff6b-187">Ouvrez le **gestionnaire de tâches et** assurez-vous que le processus **msoadfsb.exe'est** pas en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-187">Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.</span></span>
    > 1. <span data-ttu-id="7ff6b-188">Exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-188">Run the following command.</span></span>
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. <span data-ttu-id="7ff6b-189">Redémarrez Outlook.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-189">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="7ff6b-190">Debug</span><span class="sxs-lookup"><span data-stu-id="7ff6b-190">Debug</span></span>

<span data-ttu-id="7ff6b-191">Lorsque vous modifiez la gestion des événements de lancement dans votre module d’ajout, vous devez savoir que :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-191">As you make changes to launch-event handling in your add-in, you should be aware that:</span></span>

- <span data-ttu-id="7ff6b-192">Si vous avez mis à jour le manifeste, [retirez l’add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) puis chargez-le de nouveau.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-192">If you updated the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) then sideload it again.</span></span>
- <span data-ttu-id="7ff6b-193">Si vous avez apporté des modifications à des fichiers autres que le manifeste, fermez et rouvrez les Outlook sur Windows, ou actualisez l’onglet navigateur en cours d’exécution Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-193">If you made changes to files other than the manifest, close and reopen Outlook on Windows, or refresh the browser tab running Outlook on the web.</span></span>

<span data-ttu-id="7ff6b-194">Lors de la mise en œuvre de vos propres fonctionnalités, vous devrez peut-être débogdier votre code.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-194">While implementing your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="7ff6b-195">Pour obtenir des conseils sur la façon de déboger l’activation add-in basée sur les événements, [consultez Debug votre module basé sur Outlook’add-in](debug-autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="7ff6b-195">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

<span data-ttu-id="7ff6b-196">L’enregistrement de temps d’exécution est également disponible pour cette fonctionnalité Windows.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-196">Runtime logging is also available for this feature on Windows.</span></span> <span data-ttu-id="7ff6b-197">Pour plus d’informations, [consultez Votre add-in avec l’enregistrement de temps d’exécution](../testing/runtime-logging.md#runtime-logging-on-windows).</span><span class="sxs-lookup"><span data-stu-id="7ff6b-197">For more information, see [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows).</span></span>

## <a name="deploy-to-users"></a><span data-ttu-id="7ff6b-198">Déployer aux utilisateurs</span><span class="sxs-lookup"><span data-stu-id="7ff6b-198">Deploy to users</span></span>

<span data-ttu-id="7ff6b-199">Vous pouvez déployer des modules d’add-in basés sur des événements en téléchargeant le manifeste via le Microsoft 365'administration.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-199">You can deploy event-based add-ins by uploading the manifest through the Microsoft 365 admin center.</span></span> <span data-ttu-id="7ff6b-200">Dans le portail admin, élargissez la section **Paramètres** dans le volet navigation puis sélectionnez **applications intégrées**.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-200">In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.</span></span> <span data-ttu-id="7ff6b-201">Sur la page **Applications intégrées,** choisissez l’action **Télécharger’applications personnalisées.**</span><span class="sxs-lookup"><span data-stu-id="7ff6b-201">On the **Integrated apps** page, choose the **Upload custom apps** action.</span></span>

![Capture d’écran de la page Applications intégrées sur le Microsoft 365 d’administration, y compris l’action Télécharger’applications personnalisées](../images/outlook-deploy-event-based-add-ins.png)

<span data-ttu-id="7ff6b-203">AppSource et magasins inclients : La possibilité de déployer des modules d’ajout basés sur des événements ou de mettre à jour les modules d’activation existants pour inclure la fonction d’activation basée sur l’événement devrait être disponible prochainement.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-203">AppSource and inclient stores: The ability to deploy event-based add-ins or update existing add-ins to include the event-based activation feature should be available soon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7ff6b-204">Les modules d’accès basés sur des événements sont limités aux déploiements gérés par admin uniquement.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-204">Event-based add-ins are restricted to admin-managed deployments only.</span></span> <span data-ttu-id="7ff6b-205">Pour l’instant, les utilisateurs ne peuvent pas obtenir d’add-ins basés sur des événements à partir d’AppSource ou de magasins inclients.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-205">For now, users can't get event-based add-ins from AppSource or inclient stores.</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="7ff6b-206">Comportement et limitations d’activation basés sur l’événement</span><span class="sxs-lookup"><span data-stu-id="7ff6b-206">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="7ff6b-207">On s’attend à ce que les gestionnaires d’événements de lancement add-in soient de courte durée, légers et non invasifs que possible.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-207">Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible.</span></span> <span data-ttu-id="7ff6b-208">Après activation, votre module s’exécutera dans un délai d’environ 300 secondes, soit la durée maximale autorisée pour l’exécution d’add-ins basés sur l’événement. Pour signaler que votre module a terminé le traitement d’un événement de lancement, nous vous recommandons d’appeler la méthode par le gestionnaire `event.completed` associé.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-208">After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, we recommend you have the associated handler call the `event.completed` method.</span></span> <span data-ttu-id="7ff6b-209">(Notez que le code inclus après l’instruction `event.completed` n’est pas garanti pour s’exécuter.) Chaque fois qu’un événement déclenché par vos poignées d’ajout est déclenché, l’add-in est réactivé et exécute le gestionnaire d’événements associé, et la fenêtre de délai d’attente est réinitialisée.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-209">(Note that code included after the `event.completed` statement is not guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset.</span></span> <span data-ttu-id="7ff6b-210">L’add-in se termine après qu’il s’arrête, ou l’utilisateur ferme la fenêtre de composition ou envoie l’élément.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-210">The add-in ends after it times out, or the user closes the compose window or sends the item.</span></span>

<span data-ttu-id="7ff6b-211">Si l’utilisateur dispose de plusieurs modules d’ajout qui se sont abonnés au même événement, la plate-forme Outlook lance les modules sans ordre particulier.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-211">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="7ff6b-212">Actuellement, seuls cinq modules d’ajout basés sur des événements peuvent être activement en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-212">Currently, only five event-based add-ins can be actively running.</span></span>

<span data-ttu-id="7ff6b-213">L’utilisateur peut passer ou naviguer loin de l’élément de messagerie actuel où l’add-in a commencé à s’exécuter.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-213">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="7ff6b-214">L’add-in qui a été lancé terminera son opération en arrière-plan.</span><span class="sxs-lookup"><span data-stu-id="7ff6b-214">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="7ff6b-215">Certaines Office.js qui modifient ou modifient l’interface utilisateur ne sont pas autorisées à partir d’add-ins basés sur des événements. Voici les API bloquées :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-215">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="7ff6b-216">Sous `OfficeRuntime.auth` :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-216">Under `OfficeRuntime.auth`:</span></span>
  - <span data-ttu-id="7ff6b-217">`getAccessToken`(Windows seulement)</span><span class="sxs-lookup"><span data-stu-id="7ff6b-217">`getAccessToken` (Windows only)</span></span>
- <span data-ttu-id="7ff6b-218">Sous `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-218">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="7ff6b-219">Sous `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-219">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="7ff6b-220">Sous `Office.context.mailbox.item` :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-220">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="7ff6b-221">Sous `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="7ff6b-221">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="7ff6b-222">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7ff6b-222">See also</span></span>

- [<span data-ttu-id="7ff6b-223">Manifestes de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="7ff6b-223">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="7ff6b-224">Comment débobug les modules d’add-in basés sur les événements</span><span class="sxs-lookup"><span data-stu-id="7ff6b-224">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
