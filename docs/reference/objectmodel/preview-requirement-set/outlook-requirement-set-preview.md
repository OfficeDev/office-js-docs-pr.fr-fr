---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Les fonctionnalités et les API qui sont actuellement en préversion pour les compléments Outlook.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: d91105e0cfbb97dc1a239e40b1c81adc4e76988b
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626595"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="59f44-103">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="59f44-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="59f44-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="59f44-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="59f44-105">Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="59f44-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="59f44-106">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="59f44-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="59f44-107">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="59f44-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="59f44-108">Vous pouvez afficher un aperçu des fonctionnalités dans Outlook sur le Web en [configurant la version ciblée sur votre client Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="59f44-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="59f44-109">« Configurer l’accès en aperçu » est indiqué sur cette page pour les fonctionnalités applicables.</span><span class="sxs-lookup"><span data-stu-id="59f44-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="59f44-110">Pour les autres fonctionnalités, vous pouvez demander l’accès à des bits d’aperçu pour Outlook sur le Web à l’aide de votre compte Microsoft 365 en remplissant et envoyant [ce formulaire](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="59f44-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="59f44-111">« Demander un accès en aperçu » est indiqué sur ces fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="59f44-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="59f44-112">L’ensemble de conditions requises pour l’aperçu inclut toutes les fonctionnalités de l' [ensemble de conditions requises 1,9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span><span class="sxs-lookup"><span data-stu-id="59f44-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="59f44-113">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="59f44-113">Features in preview</span></span>

<span data-ttu-id="59f44-114">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="59f44-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="59f44-115">Activation des compléments sur les éléments protégés par la gestion des droits relatifs à l’information (IRM)</span><span class="sxs-lookup"><span data-stu-id="59f44-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="59f44-116">Les compléments peuvent désormais être activés sur les éléments protégés par IRM.</span><span class="sxs-lookup"><span data-stu-id="59f44-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="59f44-117">Pour activer cette fonctionnalité, un administrateur client doit activer le droit d' `OBJMODEL` utilisation en définissant l’option autoriser la stratégie personnalisée d' **accès par programme** dans Office.</span><span class="sxs-lookup"><span data-stu-id="59f44-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="59f44-118">Pour plus d’informations [, voir droits et descriptions d’utilisation](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) .</span><span class="sxs-lookup"><span data-stu-id="59f44-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="59f44-119">**Disponible dans**: Outlook sur Windows, en commençant par Build 13229,10000 (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="59f44-120">Propriétés de calendrier supplémentaires</span><span class="sxs-lookup"><span data-stu-id="59f44-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="59f44-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="59f44-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="59f44-122">Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée entière d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="59f44-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="59f44-123">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="59f44-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="59f44-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="59f44-125">Ajout d’un nouvel objet qui représente le critère de diffusion d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="59f44-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="59f44-126">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="59f44-127">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="59f44-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="59f44-128">Ajout d’une nouvelle propriété qui indique si un rendez-vous est un événement d’une journée entière.</span><span class="sxs-lookup"><span data-stu-id="59f44-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="59f44-129">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="59f44-130">Office. Context. Mailbox. Item. Sensitivity</span><span class="sxs-lookup"><span data-stu-id="59f44-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="59f44-131">Ajout d’une nouvelle propriété qui représente le critère de diffusion d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="59f44-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="59f44-132">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="59f44-133">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="59f44-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="59f44-134">Ajout d’une nouvelle énumération `AppointmentSensitivityType` qui représente les options de critère de diffusion disponibles sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="59f44-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="59f44-135">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="59f44-136">Activation basée sur un événement</span><span class="sxs-lookup"><span data-stu-id="59f44-136">Event-based activation</span></span>

<span data-ttu-id="59f44-137">Prise en charge supplémentaire de la fonctionnalité d’activation basée sur un événement dans les compléments Outlook. Pour en savoir plus, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="59f44-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="59f44-138">Point d’extension LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="59f44-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="59f44-139">Ajout `LaunchEvent` de la prise en charge du point d’extension au manifeste.</span><span class="sxs-lookup"><span data-stu-id="59f44-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="59f44-140">Il configure les fonctionnalités d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="59f44-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="59f44-141">**Disponible dans**: Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="59f44-141">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="59f44-142">Élément de manifeste LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="59f44-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="59f44-143">Ajout `LaunchEvents` de l’élément à manifest.</span><span class="sxs-lookup"><span data-stu-id="59f44-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="59f44-144">Il prend en charge la configuration de la fonctionnalité d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="59f44-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="59f44-145">**Disponible dans**: Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="59f44-145">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="59f44-146">Élément de manifeste runtimes</span><span class="sxs-lookup"><span data-stu-id="59f44-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="59f44-147">Ajout de la prise en charge d’Outlook à l' `Runtimes` élément de manifeste.</span><span class="sxs-lookup"><span data-stu-id="59f44-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="59f44-148">Il fait référence aux fichiers HTML et JavaScript nécessaires à la fonctionnalité d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="59f44-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="59f44-149">**Disponible dans**: Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="59f44-149">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="59f44-150">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="59f44-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="59f44-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="59f44-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="59f44-152">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="59f44-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="59f44-153">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (classique)</span><span class="sxs-lookup"><span data-stu-id="59f44-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="59f44-154">Signature de courrier électronique</span><span class="sxs-lookup"><span data-stu-id="59f44-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="59f44-155">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="59f44-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="59f44-156">Ajout d’une nouvelle fonction à l' `Body` objet qui ajoute ou remplace la signature dans le corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="59f44-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="59f44-157">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="59f44-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="59f44-158">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="59f44-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="59f44-159">Ajout d’une fonction qui désactive la signature client pour la boîte aux lettres d’envoi en mode composition.</span><span class="sxs-lookup"><span data-stu-id="59f44-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="59f44-160">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="59f44-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="59f44-161">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="59f44-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="59f44-162">Ajout d’une nouvelle fonction qui obtient le type de composition d’un message en mode composition.</span><span class="sxs-lookup"><span data-stu-id="59f44-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="59f44-163">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="59f44-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="59f44-164">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="59f44-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="59f44-165">Ajout d’une fonction qui vérifie si la signature client est activée sur l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="59f44-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="59f44-166">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="59f44-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="59f44-167">Office. MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="59f44-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="59f44-168">Ajout d’une nouvelle énumération `ComposeType` disponible en mode composition.</span><span class="sxs-lookup"><span data-stu-id="59f44-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="59f44-169">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="59f44-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="59f44-170">Messages de notification avec actions</span><span class="sxs-lookup"><span data-stu-id="59f44-170">Notification messages with actions</span></span>

<span data-ttu-id="59f44-171">Cette fonctionnalité permet à votre complément d’inclure un message de notification avec une action personnalisée en plus de l’action **Ignorer** par défaut.</span><span class="sxs-lookup"><span data-stu-id="59f44-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="59f44-172">Office. NotificationMessageDetails. actions</span><span class="sxs-lookup"><span data-stu-id="59f44-172">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="59f44-173">Ajout d’une nouvelle propriété qui vous permet d’ajouter une `InsightMessage` notification avec une action personnalisée.</span><span class="sxs-lookup"><span data-stu-id="59f44-173">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="59f44-174">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="59f44-174">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="59f44-175">Office. NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="59f44-175">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="59f44-176">Ajout d’un nouvel objet dans lequel vous définissez une action personnalisée pour votre `InsightMessage` notification.</span><span class="sxs-lookup"><span data-stu-id="59f44-176">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="59f44-177">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="59f44-177">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="59f44-178">Office. MailboxEnums. ActionType</span><span class="sxs-lookup"><span data-stu-id="59f44-178">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="59f44-179">Ajout d’une nouvelle énumération `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="59f44-179">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="59f44-180">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="59f44-180">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="59f44-181">Office. MailboxEnums. ItemNotificationMessageType. InsightMessage</span><span class="sxs-lookup"><span data-stu-id="59f44-181">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="59f44-182">Ajout d’un nouveau type `InsightMessage` à l' `ItemNotificationMessageType` énumération.</span><span class="sxs-lookup"><span data-stu-id="59f44-182">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="59f44-183">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)</span><span class="sxs-lookup"><span data-stu-id="59f44-183">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="59f44-184">Thème Office</span><span class="sxs-lookup"><span data-stu-id="59f44-184">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="59f44-185">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="59f44-185">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="59f44-186">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="59f44-186">Added ability to get Office theme.</span></span>

<span data-ttu-id="59f44-187">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-187">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="59f44-188">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="59f44-188">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="59f44-189">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="59f44-189">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="59f44-190">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-190">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="59f44-191">Données de session</span><span class="sxs-lookup"><span data-stu-id="59f44-191">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="59f44-192">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="59f44-192">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="59f44-193">Ajout d’un nouvel objet qui représente les données de session d’un élément.</span><span class="sxs-lookup"><span data-stu-id="59f44-193">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="59f44-194">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-194">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="59f44-195">Office. Context. Mailbox. Item. sessionData</span><span class="sxs-lookup"><span data-stu-id="59f44-195">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="59f44-196">Ajout d’une nouvelle propriété pour gérer les données de session d’un élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="59f44-196">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="59f44-197">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="59f44-197">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

## <a name="see-also"></a><span data-ttu-id="59f44-198">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="59f44-198">See also</span></span>

- [<span data-ttu-id="59f44-199">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="59f44-199">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="59f44-200">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="59f44-200">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="59f44-201">Prise en main</span><span class="sxs-lookup"><span data-stu-id="59f44-201">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="59f44-202">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="59f44-202">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
