---
title: Configurer votre complément Outlook pour l’activation basée sur des événements
description: Découvrez comment configurer votre complément Outlook pour l’activation basée sur des événements.
ms.topic: article
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: 07790ee84693596f4873bc04d53c1e76c3825b4d
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076790"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a><span data-ttu-id="b6a92-103">Configurer votre complément Outlook pour l’activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="b6a92-103">Configure your Outlook add-in for event-based activation</span></span>

<span data-ttu-id="b6a92-104">Sans la fonctionnalité d’activation basée sur des événements, un utilisateur doit lancer explicitement un complément pour effectuer ses tâches.</span><span class="sxs-lookup"><span data-stu-id="b6a92-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="b6a92-105">Cette fonctionnalité permet à votre application d’exécuter des tâches basées sur certains événements, en particulier pour les opérations qui s’appliquent à chaque élément.</span><span class="sxs-lookup"><span data-stu-id="b6a92-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="b6a92-106">Vous pouvez également intégrer le volet Des tâches et la fonctionnalité sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b6a92-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="b6a92-107">À la fin de cette walkthrough, vous aurez un add-in qui s’exécute chaque fois qu’un nouvel élément est créé et définit l’objet.</span><span class="sxs-lookup"><span data-stu-id="b6a92-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!NOTE]
> <span data-ttu-id="b6a92-108">La prise en charge de cette fonctionnalité a été introduite dans [l’ensemble de conditions requises 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="b6a92-108">Support for this feature was introduced in [requirement set 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="b6a92-109">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="b6a92-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-events"></a><span data-ttu-id="b6a92-110">Événements pris en charge</span><span class="sxs-lookup"><span data-stu-id="b6a92-110">Supported events</span></span>

<span data-ttu-id="b6a92-111">Pour l’instant, les événements suivants sont pris en charge sur le web et sur Windows.</span><span class="sxs-lookup"><span data-stu-id="b6a92-111">At present, the following events are supported on the web and on Windows.</span></span>

|<span data-ttu-id="b6a92-112">Événement</span><span class="sxs-lookup"><span data-stu-id="b6a92-112">Event</span></span>|<span data-ttu-id="b6a92-113">Description</span><span class="sxs-lookup"><span data-stu-id="b6a92-113">Description</span></span>|<span data-ttu-id="b6a92-114">Minimum</span><span class="sxs-lookup"><span data-stu-id="b6a92-114">Minimum</span></span><br><span data-ttu-id="b6a92-115">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b6a92-115">requirement set</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="b6a92-116">Lors de la composition d’un nouveau message (y compris répondre, répondre à tous et transmettre), mais pas lors de la modification, par exemple, d’un brouillon.</span><span class="sxs-lookup"><span data-stu-id="b6a92-116">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="b6a92-117">1.10</span><span class="sxs-lookup"><span data-stu-id="b6a92-117">1.10</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="b6a92-118">Lors de la création d’un rendez-vous, mais pas de la modification d’un rendez-vous existant.</span><span class="sxs-lookup"><span data-stu-id="b6a92-118">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="b6a92-119">1.10</span><span class="sxs-lookup"><span data-stu-id="b6a92-119">1.10</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="b6a92-120">Lors de l’ajout ou de la suppression de pièces jointes lors de la composition d’un message.</span><span class="sxs-lookup"><span data-stu-id="b6a92-120">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="b6a92-121">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b6a92-121">Preview</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="b6a92-122">Lors de l’ajout ou de la suppression de pièces jointes lors de la composition d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b6a92-122">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="b6a92-123">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b6a92-123">Preview</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="b6a92-124">Lors de l’ajout ou de la suppression de destinataires lors de la composition d’un message.</span><span class="sxs-lookup"><span data-stu-id="b6a92-124">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="b6a92-125">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b6a92-125">Preview</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="b6a92-126">Lors de l’ajout ou de la suppression de participants lors de la composition d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b6a92-126">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="b6a92-127">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b6a92-127">Preview</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="b6a92-128">Lors de la modification de la date et de l’heure lors de la composition d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b6a92-128">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="b6a92-129">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b6a92-129">Preview</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="b6a92-130">Lors de l’ajout, de la modification ou de la suppression des détails de la récurrence lors de la composition d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b6a92-130">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="b6a92-131">Si la date/l’heure est modifiée, `OnAppointmentTimeChanged` l’événement est également déclenché.</span><span class="sxs-lookup"><span data-stu-id="b6a92-131">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="b6a92-132">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b6a92-132">Preview</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="b6a92-133">Lors du rejet d’une notification lors de la composition d’un élément de message ou de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b6a92-133">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="b6a92-134">Seul le add-in qui a ajouté la notification sera averti.</span><span class="sxs-lookup"><span data-stu-id="b6a92-134">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="b6a92-135">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b6a92-135">Preview</span></span>|

> [!IMPORTANT]
> <span data-ttu-id="b6a92-136">Les événements toujours en prévisualisation sont disponibles uniquement avec un abonnement Microsoft 365 dans Outlook sur le web et Windows.</span><span class="sxs-lookup"><span data-stu-id="b6a92-136">Events still in preview are only available with a Microsoft 365 subscription in Outlook on the web and on Windows.</span></span> <span data-ttu-id="b6a92-137">Pour plus d’informations, voir [La prévisualisation](#how-to-preview) dans cet article.</span><span class="sxs-lookup"><span data-stu-id="b6a92-137">For more details, see [How to preview](#how-to-preview) in this article.</span></span> <span data-ttu-id="b6a92-138">Les événements d’aperçu ne doivent pas être utilisés dans les modules de production.</span><span class="sxs-lookup"><span data-stu-id="b6a92-138">Preview events shouldn't be used in production add-ins.</span></span>

### <a name="how-to-preview"></a><span data-ttu-id="b6a92-139">Comment prévisualiser</span><span class="sxs-lookup"><span data-stu-id="b6a92-139">How to preview</span></span>

<span data-ttu-id="b6a92-140">Nous vous invitons à tester les événements maintenant en prévisualisation !</span><span class="sxs-lookup"><span data-stu-id="b6a92-140">We invite you to try out the events now in preview!</span></span> <span data-ttu-id="b6a92-141">Faites-nous part de vos scénarios et de la façon dont nous pouvons les améliorer en nous faisant part de vos commentaires GitHub (voir la **section** Commentaires à la fin de cette page).</span><span class="sxs-lookup"><span data-stu-id="b6a92-141">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="b6a92-142">Pour afficher un aperçu de ces événements :</span><span class="sxs-lookup"><span data-stu-id="b6a92-142">To preview these events:</span></span>

- <span data-ttu-id="b6a92-143">Par Outlook sur le web :</span><span class="sxs-lookup"><span data-stu-id="b6a92-143">For Outlook on the web:</span></span>
  - <span data-ttu-id="b6a92-144">[Configurez la version ciblée sur votre Microsoft 365 client.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="b6a92-144">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="b6a92-145">Référencez **la bibliothèque** bêta sur le CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="b6a92-145">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="b6a92-146">Le [fichier de définition de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) type pour la compilation et la IntelliSense TypeScript se trouve aux CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="b6a92-146">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="b6a92-147">Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="b6a92-147">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="b6a92-148">Pour Outlook sur Windows :</span><span class="sxs-lookup"><span data-stu-id="b6a92-148">For Outlook on Windows:</span></span>
  - <span data-ttu-id="b6a92-149">La build minimale requise est 16.0.14026.20000.</span><span class="sxs-lookup"><span data-stu-id="b6a92-149">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="b6a92-150">Rejoignez le [Office Insider pour](https://insider.office.com) accéder à Office versions bêta.</span><span class="sxs-lookup"><span data-stu-id="b6a92-150">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="b6a92-151">Configurez le Registre.</span><span class="sxs-lookup"><span data-stu-id="b6a92-151">Configure the registry.</span></span> <span data-ttu-id="b6a92-152">Outlook inclut une copie locale des versions de production et bêta de Office.js au lieu de charger à partir du CDN.</span><span class="sxs-lookup"><span data-stu-id="b6a92-152">Outlook includes a local copy of the production and beta versions of Office.js instead of loading from the CDN.</span></span> <span data-ttu-id="b6a92-153">Par défaut, la copie de production locale de l’API est référencé.</span><span class="sxs-lookup"><span data-stu-id="b6a92-153">By default, the local production copy of the API is referenced.</span></span> <span data-ttu-id="b6a92-154">Pour basculer vers la copie bêta locale des API JavaScript Outlook, vous devez ajouter cette entrée de Registre, sinon les API bêta risquent de ne pas être trouvées.</span><span class="sxs-lookup"><span data-stu-id="b6a92-154">To switch to the local beta copy of the Outlook JavaScript APIs, you need to add this registry entry, otherwise beta APIs may not be found.</span></span>
    1. <span data-ttu-id="b6a92-155">Créez la clé de `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` Registre.</span><span class="sxs-lookup"><span data-stu-id="b6a92-155">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="b6a92-156">Ajoutez une entrée nommée `EnableBetaAPIsInJavaScript` et définissez la valeur sur `1` .</span><span class="sxs-lookup"><span data-stu-id="b6a92-156">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span> <span data-ttu-id="b6a92-157">L’image suivante indique à quoi doit ressembler le registre.</span><span class="sxs-lookup"><span data-stu-id="b6a92-157">The following image shows what the registry should look like.</span></span>

        ![Capture d’écran de l’éditeur du Registre avec une valeur de clé de Registre EnableBetaAPIsInJavaScript.](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a><span data-ttu-id="b6a92-159">Configuration de votre environnement</span><span class="sxs-lookup"><span data-stu-id="b6a92-159">Set up your environment</span></span>

<span data-ttu-id="b6a92-160">[Complétez Outlook démarrage](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rapide qui crée un projet de compl?ment avec le générateur Yeoman pour Office compl?ments.</span><span class="sxs-lookup"><span data-stu-id="b6a92-160">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="b6a92-161">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="b6a92-161">Configure the manifest</span></span>

<span data-ttu-id="b6a92-162">Pour activer l’activation basée sur des événements de votre complément, vous devez configurer l’élément [Runtimes](../reference/manifest/runtimes.md) et le point d’extension [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) dans le nœud `VersionOverridesV1_1` du manifeste.</span><span class="sxs-lookup"><span data-stu-id="b6a92-162">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="b6a92-163">Pour l’instant, `DesktopFormFactor` est le seul facteur de forme pris en charge.</span><span class="sxs-lookup"><span data-stu-id="b6a92-163">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="b6a92-164">Dans votre éditeur de code, ouvrez le projet de démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="b6a92-164">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="b6a92-165">Ouvrez **lemanifest.xml** situé à la racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="b6a92-165">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="b6a92-166">Sélectionnez l’intégralité du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le par le `<VersionOverrides>` code XML suivant, puis enregistrez vos modifications.</span><span class="sxs-lookup"><span data-stu-id="b6a92-166">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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

<span data-ttu-id="b6a92-167">Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web utilise un fichier HTML qui peut référencer le même fichier JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b6a92-167">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="b6a92-168">Vous devez fournir des références à ces deux fichiers dans le nœud du manifeste, car la plateforme Outlook détermine en fin de compte s’il faut utiliser du code HTML ou JavaScript en fonction du `Resources` client Outlook.</span><span class="sxs-lookup"><span data-stu-id="b6a92-168">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="b6a92-169">En tant que tel, pour configurer la gestion des événements, fournissez l’emplacement du code HTML dans l’élément, puis, dans son élément enfant, fournissez l’emplacement du fichier JavaScript indiqué ou référencé par le `Runtime` `Override` code HTML.</span><span class="sxs-lookup"><span data-stu-id="b6a92-169">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="b6a92-170">Pour en savoir plus sur les manifestes de Outlook de votre Outlook, consultez la Outlook [des manifestes de modules.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="b6a92-170">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="b6a92-171">Implémenter la gestion des événements</span><span class="sxs-lookup"><span data-stu-id="b6a92-171">Implement event handling</span></span>

<span data-ttu-id="b6a92-172">Vous devez implémenter la gestion de vos événements sélectionnés.</span><span class="sxs-lookup"><span data-stu-id="b6a92-172">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="b6a92-173">Dans ce scénario, vous allez ajouter la gestion de la composition de nouveaux éléments.</span><span class="sxs-lookup"><span data-stu-id="b6a92-173">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="b6a92-174">À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="b6a92-174">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="b6a92-175">Après la `action` fonction, insérez les fonctions JavaScript suivantes.</span><span class="sxs-lookup"><span data-stu-id="b6a92-175">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="b6a92-176">Ajoutez le code JavaScript suivant à la fin du fichier.</span><span class="sxs-lookup"><span data-stu-id="b6a92-176">Add the following JavaScript code at the end of the file.</span></span>

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. <span data-ttu-id="b6a92-177">Enregistrez vos modifications.</span><span class="sxs-lookup"><span data-stu-id="b6a92-177">Save your changes.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b6a92-178">Windows : actuellement, les importations ne sont pas pris en charge dans le fichier JavaScript où vous implémentez la gestion de l’activation basée sur des événements.</span><span class="sxs-lookup"><span data-stu-id="b6a92-178">Windows: At present, imports are not supported in the JavaScript file where you implement the handling for event-based activation.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="b6a92-179">Essayez</span><span class="sxs-lookup"><span data-stu-id="b6a92-179">Try it out</span></span>

1. <span data-ttu-id="b6a92-180">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="b6a92-180">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="b6a92-181">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution) et votre complément est chargé.</span><span class="sxs-lookup"><span data-stu-id="b6a92-181">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="b6a92-182">Si votre application [n’a](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) pas été automatiquement chargé de manière test, suivez les instructions du chargement de version test des Outlook pour tester le chargement de version test du Outlook.</span><span class="sxs-lookup"><span data-stu-id="b6a92-182">If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="b6a92-183">Dans Outlook sur le web, créez un message.</span><span class="sxs-lookup"><span data-stu-id="b6a92-183">In Outlook on the web, create a new message.</span></span>

    ![Capture d’écran d’une fenêtre de message Outlook sur le web avec l’objet de la composition.](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="b6a92-185">Dans Outlook sur Windows, créez un message.</span><span class="sxs-lookup"><span data-stu-id="b6a92-185">In Outlook on Windows, create a new message.</span></span>

    ![Capture d’écran d’une fenêtre de message Outlook sur Windows avec l’objet définie sur composition.](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="b6a92-187">Si vous exécutez votre add-in à partir de localhost et que vous voyez l’erreur « Nous sommes désolés, nous n’avons pas pu accéder à *{votre-add-in-name-here}*».</span><span class="sxs-lookup"><span data-stu-id="b6a92-187">If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*.</span></span> <span data-ttu-id="b6a92-188">Assurez-vous que vous avez une connexion réseau.</span><span class="sxs-lookup"><span data-stu-id="b6a92-188">Make sure you have a network connection.</span></span> <span data-ttu-id="b6a92-189">Si le problème persiste, veuillez essayer à nouveau plus tard. », vous devrez peut-être activer une exemption de bouclisation.</span><span class="sxs-lookup"><span data-stu-id="b6a92-189">If the problem continues, please try again later.", you may need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="b6a92-190">Fermez Outlook.</span><span class="sxs-lookup"><span data-stu-id="b6a92-190">Close Outlook.</span></span>
    > 1. <span data-ttu-id="b6a92-191">Ouvrez **le Gestionnaire des tâches** et assurez-vous que le processus **msoadfsb.exe** n’est pas en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="b6a92-191">Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.</span></span>
    > 1. <span data-ttu-id="b6a92-192">Exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="b6a92-192">Run the following command.</span></span>
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. <span data-ttu-id="b6a92-193">Redémarrez Outlook.</span><span class="sxs-lookup"><span data-stu-id="b6a92-193">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="b6a92-194">Debug</span><span class="sxs-lookup"><span data-stu-id="b6a92-194">Debug</span></span>

<span data-ttu-id="b6a92-195">Lorsque vous modifiez la gestion des événements de lancement dans votre add-in, vous devez savoir que :</span><span class="sxs-lookup"><span data-stu-id="b6a92-195">As you make changes to launch-event handling in your add-in, you should be aware that:</span></span>

- <span data-ttu-id="b6a92-196">Si vous avez mis à jour le manifeste, [supprimez-le, puis chargez-le](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) de nouveau.</span><span class="sxs-lookup"><span data-stu-id="b6a92-196">If you updated the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) then sideload it again.</span></span>
- <span data-ttu-id="b6a92-197">Si vous avez apporté des modifications à des fichiers autres que le manifeste, fermez et rouvrez Outlook sur Windows ou actualisez l’onglet du navigateur en cours d’exécution Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="b6a92-197">If you made changes to files other than the manifest, close and reopen Outlook on Windows, or refresh the browser tab running Outlook on the web.</span></span>

<span data-ttu-id="b6a92-198">Lors de l’implémentation de vos propres fonctionnalités, vous devrez peut-être déboguer votre code.</span><span class="sxs-lookup"><span data-stu-id="b6a92-198">While implementing your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="b6a92-199">Pour obtenir des instructions sur le débogage de l’activation de complément basée sur des événements, voir [Déboguer](debug-autolaunch.md)votre complément basé sur Outlook événement.</span><span class="sxs-lookup"><span data-stu-id="b6a92-199">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

<span data-ttu-id="b6a92-200">La journalisation runtime est également disponible pour cette fonctionnalité sur Windows.</span><span class="sxs-lookup"><span data-stu-id="b6a92-200">Runtime logging is also available for this feature on Windows.</span></span> <span data-ttu-id="b6a92-201">Pour plus d’informations, voir [Déboguer votre add-in avec la journalisation runtime.](../testing/runtime-logging.md#runtime-logging-on-windows)</span><span class="sxs-lookup"><span data-stu-id="b6a92-201">For more information, see [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows).</span></span>

## <a name="deploy-to-users"></a><span data-ttu-id="b6a92-202">Déployer pour les utilisateurs</span><span class="sxs-lookup"><span data-stu-id="b6a92-202">Deploy to users</span></span>

<span data-ttu-id="b6a92-203">Vous pouvez déployer des add-ins basés sur des événements en chargeant le manifeste via le Centre d’administration Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="b6a92-203">You can deploy event-based add-ins by uploading the manifest through the Microsoft 365 admin center.</span></span> <span data-ttu-id="b6a92-204">Dans le portail d’administration, **développez la** section Paramètres dans le volet de navigation, puis sélectionnez **Applications intégrées.**</span><span class="sxs-lookup"><span data-stu-id="b6a92-204">In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.</span></span> <span data-ttu-id="b6a92-205">Dans la page **Applications intégrées,** sélectionnez l Télécharger **d’applications personnalisées.**</span><span class="sxs-lookup"><span data-stu-id="b6a92-205">On the **Integrated apps** page, choose the **Upload custom apps** action.</span></span>

![Capture d’écran de la page Applications intégrées sur le Centre d’administration Microsoft 365, y compris l’action Télécharger d’applications personnalisées.](../images/outlook-deploy-event-based-add-ins.png)

<span data-ttu-id="b6a92-207">Magasins AppSource et inclients : la possibilité de déployer des compléments basés sur des événements ou de mettre à jour des compléments existants pour inclure la fonctionnalité d’activation basée sur des événements devrait être disponible prochainement.</span><span class="sxs-lookup"><span data-stu-id="b6a92-207">AppSource and inclient stores: The ability to deploy event-based add-ins or update existing add-ins to include the event-based activation feature should be available soon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b6a92-208">Les add-ins basés sur des événements sont limités aux déploiements gérés par l’administrateur uniquement.</span><span class="sxs-lookup"><span data-stu-id="b6a92-208">Event-based add-ins are restricted to admin-managed deployments only.</span></span> <span data-ttu-id="b6a92-209">Pour l’instant, les utilisateurs ne peuvent pas obtenir de add-ins basés sur des événements à partir d’AppSource ou de magasins inclients.</span><span class="sxs-lookup"><span data-stu-id="b6a92-209">For now, users can't get event-based add-ins from AppSource or inclient stores.</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="b6a92-210">Comportement et limitations de l’activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="b6a92-210">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="b6a92-211">Les handlers d’événements de lancement de modules sont censés être de courte durée, légers et aussi peu invasifs que possible.</span><span class="sxs-lookup"><span data-stu-id="b6a92-211">Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible.</span></span> <span data-ttu-id="b6a92-212">Après l’activation, votre complément prendra un délai d’environ 300 secondes, durée maximale autorisée pour l’exécution de compléments basés sur des événements. Pour signaler que votre add-in a terminé le traitement d’un événement de lancement, nous vous recommandons d’avoir le handler associé qui appelle la `event.completed` méthode.</span><span class="sxs-lookup"><span data-stu-id="b6a92-212">After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, we recommend you have the associated handler call the `event.completed` method.</span></span> <span data-ttu-id="b6a92-213">(Notez que le code inclus après `event.completed` l’instruction n’est pas garanti pour s’exécuter.) Chaque fois qu’un événement géré par votre add-in est déclenché, celui-ci est réactivé et exécute le handler d’événements associé, et la fenêtre d’délai est réinitialisée.</span><span class="sxs-lookup"><span data-stu-id="b6a92-213">(Note that code included after the `event.completed` statement is not guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset.</span></span> <span data-ttu-id="b6a92-214">Le add-in se termine à l’issue de son utilisation, ou l’utilisateur ferme la fenêtre de composition ou envoie l’élément.</span><span class="sxs-lookup"><span data-stu-id="b6a92-214">The add-in ends after it times out, or the user closes the compose window or sends the item.</span></span>

<span data-ttu-id="b6a92-215">Si l’utilisateur a plusieurs add-ins abonnés au même événement, la plateforme Outlook lance les modules dans un ordre particulier.</span><span class="sxs-lookup"><span data-stu-id="b6a92-215">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="b6a92-216">Actuellement, seuls cinq add-ins basés sur des événements peuvent être activement en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="b6a92-216">Currently, only five event-based add-ins can be actively running.</span></span>

<span data-ttu-id="b6a92-217">L’utilisateur peut basculer ou naviguer à partir de l’élément de messagerie actuel où le module a commencé à s’exécute.</span><span class="sxs-lookup"><span data-stu-id="b6a92-217">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="b6a92-218">Le module qui a été lancé terminera son opération en arrière-plan.</span><span class="sxs-lookup"><span data-stu-id="b6a92-218">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="b6a92-219">Les importations ne sont pas pris en charge dans le fichier JavaScript où vous implémentez la gestion de l’activation basée sur des événements dans Windows client.</span><span class="sxs-lookup"><span data-stu-id="b6a92-219">Imports are not supported in the JavaScript file where you implement the handling for event-based activation in the Windows client.</span></span>

<span data-ttu-id="b6a92-220">Certaines Office.js API qui modifient ou modifient l’interface utilisateur ne sont pas autorisées à partir des add-ins basés sur des événements. Les API bloquées sont les suivantes :</span><span class="sxs-lookup"><span data-stu-id="b6a92-220">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="b6a92-221">Sous `OfficeRuntime.auth` :</span><span class="sxs-lookup"><span data-stu-id="b6a92-221">Under `OfficeRuntime.auth`:</span></span>
  - <span data-ttu-id="b6a92-222">`getAccessToken`(Windows uniquement)</span><span class="sxs-lookup"><span data-stu-id="b6a92-222">`getAccessToken` (Windows only)</span></span>
- <span data-ttu-id="b6a92-223">Sous `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="b6a92-223">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="b6a92-224">Sous `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="b6a92-224">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="b6a92-225">Sous `Office.context.mailbox.item` :</span><span class="sxs-lookup"><span data-stu-id="b6a92-225">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="b6a92-226">Sous `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="b6a92-226">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="b6a92-227">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b6a92-227">See also</span></span>

- [<span data-ttu-id="b6a92-228">Manifestes de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="b6a92-228">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="b6a92-229">Comment déboguer des add-ins basés sur des événements</span><span class="sxs-lookup"><span data-stu-id="b6a92-229">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
- <span data-ttu-id="b6a92-230">Exemple PnP : utiliser Outlook activation basée sur un événement [pour définir la signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)</span><span class="sxs-lookup"><span data-stu-id="b6a92-230">PnP sample: [Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)</span></span>