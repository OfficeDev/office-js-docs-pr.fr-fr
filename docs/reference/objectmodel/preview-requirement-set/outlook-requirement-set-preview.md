---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Les fonctionnalités et les API qui sont actuellement en préversion pour les compléments Outlook.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 600aad32c394d35e62f4024808b185e8a9abe5e8
ms.sourcegitcommit: 09a8683ff29cf06d0d1d822be83cf0798f1ccdf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/01/2020
ms.locfileid: "44471344"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="2affd-103">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="2affd-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="2affd-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2affd-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2affd-105">Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="2affd-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="2affd-106">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="2affd-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="2affd-107">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="2affd-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="2affd-108">Vous pouvez afficher un aperçu des fonctionnalités dans Outlook sur le Web en [configurant la version ciblée sur votre client Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="2affd-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="2affd-109">« Configurer l’accès en aperçu » est indiqué sur cette page pour les fonctionnalités applicables.</span><span class="sxs-lookup"><span data-stu-id="2affd-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="2affd-110">Pour les autres fonctionnalités, vous pouvez demander l’accès à des bits d’aperçu pour Outlook sur le Web à l’aide de votre compte Microsoft 365 en remplissant et envoyant [ce formulaire](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="2affd-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="2affd-111">« Demander un accès en aperçu » est indiqué sur ces fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="2affd-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="2affd-112">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="2affd-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="2affd-113">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="2affd-113">Features in preview</span></span>

<span data-ttu-id="2affd-114">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="2affd-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="2affd-115">Propriétés de calendrier supplémentaires</span><span class="sxs-lookup"><span data-stu-id="2affd-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="2affd-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="2affd-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="2affd-117">Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée entière d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2affd-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="2affd-118">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="2affd-118">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="2affd-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="2affd-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="2affd-120">Ajout d’un nouvel objet qui représente le critère de diffusion d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2affd-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="2affd-121">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="2affd-121">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="2affd-122">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="2affd-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="2affd-123">Ajout d’une nouvelle propriété qui indique si un rendez-vous est un événement d’une journée entière.</span><span class="sxs-lookup"><span data-stu-id="2affd-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="2affd-124">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="2affd-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="2affd-125">Office. Context. Mailbox. Item. Sensitivity</span><span class="sxs-lookup"><span data-stu-id="2affd-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="2affd-126">Ajout d’une nouvelle propriété qui représente le critère de diffusion d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2affd-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="2affd-127">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="2affd-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="2affd-128">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="2affd-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="2affd-129">Ajout d’une nouvelle énumération `AppointmentSensitivityType` qui représente les options de critère de diffusion disponibles sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2affd-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="2affd-130">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="2affd-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="2affd-131">Ajouter à l’envoi</span><span class="sxs-lookup"><span data-stu-id="2affd-131">Append on send</span></span>

<span data-ttu-id="2affd-132">Pour en savoir plus sur l’utilisation de la fonctionnalité Ajout à l’envoi, consultez la rubrique [implémenter Append lors de l’envoi dans votre complément Outlook](../../../outlook/append-on-send.md).</span><span class="sxs-lookup"><span data-stu-id="2affd-132">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="2affd-133">Office. Context. Mailbox. Item. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="2affd-133">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="2affd-134">Ajout d’une nouvelle fonction à l' `Body` objet qui ajoute des données à la fin du corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2affd-134">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="2affd-135">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="2affd-135">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="2affd-136">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="2affd-136">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="2affd-137">Ajout d’un nouvel élément au manifeste dans lequel l' `AppendOnSend` autorisation étendue doit être incluse dans la collection des autorisations étendues.</span><span class="sxs-lookup"><span data-stu-id="2affd-137">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="2affd-138">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="2affd-138">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="2affd-139">Activation basée sur les événements</span><span class="sxs-lookup"><span data-stu-id="2affd-139">Event-based activation</span></span>

<span data-ttu-id="2affd-140">Prise en charge supplémentaire de la fonctionnalité d’activation basée sur un événement dans les compléments Outlook. Pour en savoir plus, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="2affd-140">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="2affd-141">Point d’extension LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="2affd-141">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="2affd-142">Ajout `LaunchEvent` de la prise en charge du point d’extension au manifeste.</span><span class="sxs-lookup"><span data-stu-id="2affd-142">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="2affd-143">Il configure les fonctionnalités d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="2affd-143">It configures event-based activation functionality.</span></span>

<span data-ttu-id="2affd-144">**Disponible dans**: Outlook sur le Web (moderne, [demander un accès en aperçu](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="2affd-144">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="2affd-145">Élément de manifeste LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="2affd-145">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="2affd-146">Ajout `LaunchEvents` de l’élément à manifest.</span><span class="sxs-lookup"><span data-stu-id="2affd-146">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="2affd-147">Il prend en charge la configuration de la fonctionnalité d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="2affd-147">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="2affd-148">**Disponible dans**: Outlook sur le Web (moderne, [demander un accès en aperçu](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="2affd-148">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="2affd-149">Élément de manifeste runtimes</span><span class="sxs-lookup"><span data-stu-id="2affd-149">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="2affd-150">Ajout de la prise en charge d’Outlook à l' `Runtimes` élément de manifeste.</span><span class="sxs-lookup"><span data-stu-id="2affd-150">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="2affd-151">Il fait référence aux fichiers HTML et JavaScript nécessaires à la fonctionnalité d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="2affd-151">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="2affd-152">**Disponible dans**: Outlook sur le Web (moderne, [demander un accès en aperçu](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="2affd-152">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="2affd-153">Obtenir toutes les propriétés personnalisées</span><span class="sxs-lookup"><span data-stu-id="2affd-153">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="2affd-154">CustomProperties. getAll</span><span class="sxs-lookup"><span data-stu-id="2affd-154">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

<span data-ttu-id="2affd-155">Ajout d’une nouvelle fonction à l' `CustomProperties` objet qui obtient toutes les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="2affd-155">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="2affd-156">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne), Outlook sur Mac (connecté à l’abonnement Office 365), Outlook sur Android, Outlook sur iOS</span><span class="sxs-lookup"><span data-stu-id="2affd-156">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="2affd-157">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="2affd-157">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="2affd-158">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="2affd-158">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="2affd-159">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="2affd-159">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="2affd-160">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="2affd-160">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="2affd-161">Signature de courrier électronique</span><span class="sxs-lookup"><span data-stu-id="2affd-161">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="2affd-162">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="2affd-162">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="2affd-163">Ajout d’une nouvelle fonction à l' `Body` objet qui ajoute ou remplace la signature dans le corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2affd-163">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="2affd-164">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="2affd-164">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="2affd-165">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="2affd-165">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="2affd-166">Ajout d’une fonction qui désactive la signature client pour la boîte aux lettres d’envoi en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2affd-166">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="2affd-167">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="2affd-167">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="2affd-168">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="2affd-168">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="2affd-169">Ajout d’une nouvelle fonction qui obtient le type de composition d’un message en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2affd-169">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="2affd-170">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="2affd-170">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="2affd-171">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="2affd-171">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="2affd-172">Ajout d’une fonction qui vérifie si la signature client est activée sur l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2affd-172">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="2affd-173">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="2affd-173">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="2affd-174">Office. MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="2affd-174">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="2affd-175">Ajout d’une nouvelle énumération `ComposeType` disponible en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2affd-175">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="2affd-176">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="2affd-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="2affd-177">Thème Office</span><span class="sxs-lookup"><span data-stu-id="2affd-177">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="2affd-178">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="2affd-178">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="2affd-179">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="2affd-179">Added ability to get Office theme.</span></span>

<span data-ttu-id="2affd-180">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="2affd-180">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="2affd-181">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="2affd-181">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="2affd-182">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="2affd-182">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="2affd-183">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="2affd-183">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="2affd-184">Authentification unique (SSO)</span><span class="sxs-lookup"><span data-stu-id="2affd-184">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="2affd-185">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="2affd-185">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="2affd-186">Ajout d’un accès à `getAccessToken`, qui permet aux compléments d’[obtenir un jeton d’accès](../../../outlook/authenticate-a-user-with-an-sso-token.md) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2affd-186">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="2affd-187">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (moderne), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="2affd-187">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="2affd-188">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2affd-188">See also</span></span>

- [<span data-ttu-id="2affd-189">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2affd-189">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="2affd-190">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2affd-190">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="2affd-191">Prise en main</span><span class="sxs-lookup"><span data-stu-id="2affd-191">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="2affd-192">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="2affd-192">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
