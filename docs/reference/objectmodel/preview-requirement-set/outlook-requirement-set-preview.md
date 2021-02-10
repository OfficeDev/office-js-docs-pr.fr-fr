---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Fonctionnalités et API actuellement en prévisualisation pour les add-ins Outlook.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 92ba3510af0c8b9ebdf9ca4368c889b821a9cb3b
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173954"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="ddb00-103">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="ddb00-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="ddb00-104">Le sous-ensemble de l’API de l’API JavaScript pour Outlook inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un application Outlook.</span><span class="sxs-lookup"><span data-stu-id="ddb00-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ddb00-105">Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="ddb00-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="ddb00-106">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="ddb00-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="ddb00-107">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="ddb00-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="ddb00-108">Vous pourrez peut-être afficher un aperçu des fonctionnalités dans Outlook sur le web en configurant la version ciblée [sur votre client Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="ddb00-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="ddb00-109">Cette page indique « Configurer l’accès en aperçu » pour les fonctionnalités applicables.</span><span class="sxs-lookup"><span data-stu-id="ddb00-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="ddb00-110">Pour d’autres fonctionnalités, vous pouvez demander l’accès aux bits d’aperçu pour Outlook sur le web à l’aide de votre compte Microsoft 365 en remplissant et en envoyant [ce formulaire.](https://aka.ms/OWAPreview)</span><span class="sxs-lookup"><span data-stu-id="ddb00-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="ddb00-111">« Demander l’accès en prévisualisation » est indiqué sur ces fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="ddb00-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="ddb00-112">L’ensemble de conditions requises preview inclut toutes les fonctionnalités de l’ensemble de conditions [requises 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span><span class="sxs-lookup"><span data-stu-id="ddb00-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="ddb00-113">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="ddb00-113">Features in preview</span></span>

<span data-ttu-id="ddb00-114">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="ddb00-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="ddb00-115">Activation de compléments sur des éléments protégés par la Gestion des droits de l’information (IRM)</span><span class="sxs-lookup"><span data-stu-id="ddb00-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="ddb00-116">Les add-ins peuvent désormais être activés sur les éléments protégés par IRM.</span><span class="sxs-lookup"><span data-stu-id="ddb00-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="ddb00-117">Pour activer cette fonctionnalité, un administrateur client doit activer le droit d’utilisation en paramètres de stratégie personnalisée Autoriser l’accès par programmation `OBJMODEL` dans Office. </span><span class="sxs-lookup"><span data-stu-id="ddb00-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="ddb00-118">Pour plus [d’informations, voir droits d’utilisation et descriptions.](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions)</span><span class="sxs-lookup"><span data-stu-id="ddb00-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="ddb00-119">**Disponible dans**: Outlook sur Windows, à partir de la build 13229.10000 (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ddb00-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="ddb00-120">Propriétés de calendrier supplémentaires</span><span class="sxs-lookup"><span data-stu-id="ddb00-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="ddb00-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="ddb00-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="ddb00-122">Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée d’un rendez-vous en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="ddb00-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="ddb00-123">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ddb00-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="ddb00-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="ddb00-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="ddb00-125">Ajout d’un nouvel objet qui représente la sensibilité d’un rendez-vous en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="ddb00-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="ddb00-126">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ddb00-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="ddb00-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="ddb00-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="ddb00-128">Ajout d’une nouvelle propriété qui représente si un rendez-vous est un événement d’une journée.</span><span class="sxs-lookup"><span data-stu-id="ddb00-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="ddb00-129">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ddb00-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="ddb00-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="ddb00-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="ddb00-131">Ajout d’une nouvelle propriété qui représente la sensibilité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="ddb00-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="ddb00-132">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ddb00-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="ddb00-133">Office.MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="ddb00-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="ddb00-134">Ajout d’une nouvelle enum `AppointmentSensitivityType` qui représente les options de sensibilité disponibles sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="ddb00-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="ddb00-135">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ddb00-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="ddb00-136">Activation basée sur un événement</span><span class="sxs-lookup"><span data-stu-id="ddb00-136">Event-based activation</span></span>

<span data-ttu-id="ddb00-137">Prise en charge supplémentaire de la fonctionnalité d’activation basée sur des événements dans les compléments Outlook. Pour [plus d’informations,](../../../outlook/autolaunch.md) voir Configurer votre complément Outlook pour l’activation basée sur des événements.</span><span class="sxs-lookup"><span data-stu-id="ddb00-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="ddb00-138">Point d’extension LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="ddb00-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="ddb00-139">Ajout de `LaunchEvent` la prise en charge du point d’extension au manifeste.</span><span class="sxs-lookup"><span data-stu-id="ddb00-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="ddb00-140">Il configure la fonctionnalité d’activation basée sur des événements.</span><span class="sxs-lookup"><span data-stu-id="ddb00-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="ddb00-141">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="ddb00-141">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="ddb00-142">Élément manifeste LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="ddb00-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="ddb00-143">Ajout `LaunchEvents` d’un élément au manifeste.</span><span class="sxs-lookup"><span data-stu-id="ddb00-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="ddb00-144">Il prend en charge la configuration de la fonctionnalité d’activation basée sur des événements.</span><span class="sxs-lookup"><span data-stu-id="ddb00-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="ddb00-145">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="ddb00-145">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="ddb00-146">Élément manifeste Runtimes</span><span class="sxs-lookup"><span data-stu-id="ddb00-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="ddb00-147">Ajout de la prise en charge d’Outlook à `Runtimes` l’élément manifeste.</span><span class="sxs-lookup"><span data-stu-id="ddb00-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="ddb00-148">Il fait référence aux fichiers HTML et JavaScript nécessaires pour la fonctionnalité d’activation basée sur des événements.</span><span class="sxs-lookup"><span data-stu-id="ddb00-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="ddb00-149">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="ddb00-149">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="ddb00-150">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="ddb00-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="ddb00-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="ddb00-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="ddb00-152">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="ddb00-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="ddb00-153">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="ddb00-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="ddb00-154">Signature électronique</span><span class="sxs-lookup"><span data-stu-id="ddb00-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="ddb00-155">Office.context.mailbox.item.body.setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="ddb00-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="ddb00-156">Ajout d’une nouvelle fonction à l’objet qui ajoute ou remplace la signature dans le corps de l’élément `Body` en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="ddb00-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="ddb00-157">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="ddb00-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="ddb00-158">Office.context.mailbox.item.disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="ddb00-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="ddb00-159">Ajout d’une nouvelle fonction qui désactive la signature du client pour la boîte aux lettres d’envoi en mode composition.</span><span class="sxs-lookup"><span data-stu-id="ddb00-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="ddb00-160">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="ddb00-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="ddb00-161">Office.context.mailbox.item.getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="ddb00-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="ddb00-162">Ajout d’une nouvelle fonction qui obtient le type de composition d’un message en mode composition.</span><span class="sxs-lookup"><span data-stu-id="ddb00-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="ddb00-163">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="ddb00-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="ddb00-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="ddb00-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="ddb00-165">Ajout d’une nouvelle fonction qui vérifie si la signature du client est activée sur l’élément en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="ddb00-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="ddb00-166">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="ddb00-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="ddb00-167">Office.MailboxEnums.ComposeType</span><span class="sxs-lookup"><span data-stu-id="ddb00-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="ddb00-168">Ajout d’une nouvelle enum `ComposeType` disponible en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="ddb00-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="ddb00-169">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="ddb00-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="ddb00-170">Messages de notification avec actions</span><span class="sxs-lookup"><span data-stu-id="ddb00-170">Notification messages with actions</span></span>

<span data-ttu-id="ddb00-171">Cette fonctionnalité permet à votre add-in d’inclure un message de notification avec une action personnalisée en plus de l’action d’ignorer **par** défaut.</span><span class="sxs-lookup"><span data-stu-id="ddb00-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span> <span data-ttu-id="ddb00-172">Dans Outlook sur le web moderne, cette fonctionnalité est disponible en mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="ddb00-172">In modern Outlook on the web, this feature is available in Compose mode only.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="ddb00-173">Office.NotificationMessageDetails.actions</span><span class="sxs-lookup"><span data-stu-id="ddb00-173">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="ddb00-174">Ajout d’une nouvelle propriété qui vous permet d’ajouter une `InsightMessage` notification avec une action personnalisée.</span><span class="sxs-lookup"><span data-stu-id="ddb00-174">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="ddb00-175">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="ddb00-175">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="ddb00-176">Office.NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="ddb00-176">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="ddb00-177">Ajout d’un nouvel objet dans lequel vous définissez une action personnalisée pour votre `InsightMessage` notification.</span><span class="sxs-lookup"><span data-stu-id="ddb00-177">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="ddb00-178">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="ddb00-178">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="ddb00-179">Office.MailboxEnums.ActionType</span><span class="sxs-lookup"><span data-stu-id="ddb00-179">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="ddb00-180">Ajout d’une nouvelle enum `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="ddb00-180">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="ddb00-181">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="ddb00-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="ddb00-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span><span class="sxs-lookup"><span data-stu-id="ddb00-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="ddb00-183">Ajout d’un nouveau type `InsightMessage` à `ItemNotificationMessageType` l’enum.</span><span class="sxs-lookup"><span data-stu-id="ddb00-183">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="ddb00-184">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="ddb00-184">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="ddb00-185">Thème Office</span><span class="sxs-lookup"><span data-stu-id="ddb00-185">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="ddb00-186">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="ddb00-186">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="ddb00-187">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="ddb00-187">Added ability to get Office theme.</span></span>

<span data-ttu-id="ddb00-188">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ddb00-188">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="ddb00-189">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="ddb00-189">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="ddb00-190">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="ddb00-190">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="ddb00-191">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ddb00-191">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="ddb00-192">Données de session</span><span class="sxs-lookup"><span data-stu-id="ddb00-192">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="ddb00-193">Office.SessionData</span><span class="sxs-lookup"><span data-stu-id="ddb00-193">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="ddb00-194">Ajout d’un nouvel objet qui représente les données de session d’un élément.</span><span class="sxs-lookup"><span data-stu-id="ddb00-194">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="ddb00-195">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="ddb00-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="ddb00-196">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="ddb00-196">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="ddb00-197">Ajout d’une nouvelle propriété pour gérer les données de session d’un élément en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="ddb00-197">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="ddb00-198">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="ddb00-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="ddb00-199">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ddb00-199">See also</span></span>

- [<span data-ttu-id="ddb00-200">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="ddb00-200">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="ddb00-201">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="ddb00-201">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="ddb00-202">Prise en main</span><span class="sxs-lookup"><span data-stu-id="ddb00-202">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="ddb00-203">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="ddb00-203">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
