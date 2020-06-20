---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Les fonctionnalités et les API qui sont actuellement en préversion pour les compléments Outlook.
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: d165d6ff82edf66034bb90ea40d522a23f919191
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778661"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="bba69-103">Ensemble de conditions requises de l’API du complément Outlook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="bba69-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="bba69-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="bba69-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bba69-105">Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="bba69-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="bba69-106">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="bba69-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="bba69-107">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="bba69-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="bba69-108">Vous pouvez afficher un aperçu des fonctionnalités dans Outlook sur le Web en [configurant la version ciblée sur votre client Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="bba69-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="bba69-109">« Configurer l’accès en aperçu » est indiqué sur cette page pour les fonctionnalités applicables.</span><span class="sxs-lookup"><span data-stu-id="bba69-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="bba69-110">Pour les autres fonctionnalités, vous pouvez demander l’accès à des bits d’aperçu pour Outlook sur le Web à l’aide de votre compte Microsoft 365 en remplissant et envoyant [ce formulaire](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="bba69-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="bba69-111">« Demander un accès en aperçu » est indiqué sur ces fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="bba69-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="bba69-112">L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="bba69-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="bba69-113">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="bba69-113">Features in preview</span></span>

<span data-ttu-id="bba69-114">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="bba69-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="bba69-115">Propriétés de calendrier supplémentaires</span><span class="sxs-lookup"><span data-stu-id="bba69-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="bba69-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="bba69-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="bba69-117">Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée entière d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bba69-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="bba69-118">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-118">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="bba69-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="bba69-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="bba69-120">Ajout d’un nouvel objet qui représente le critère de diffusion d’un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bba69-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="bba69-121">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-121">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="bba69-122">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="bba69-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="bba69-123">Ajout d’une nouvelle propriété qui indique si un rendez-vous est un événement d’une journée entière.</span><span class="sxs-lookup"><span data-stu-id="bba69-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="bba69-124">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="bba69-125">Office. Context. Mailbox. Item. Sensitivity</span><span class="sxs-lookup"><span data-stu-id="bba69-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="bba69-126">Ajout d’une nouvelle propriété qui représente le critère de diffusion d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bba69-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="bba69-127">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="bba69-128">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="bba69-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="bba69-129">Ajout d’une nouvelle énumération `AppointmentSensitivityType` qui représente les options de critère de diffusion disponibles sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bba69-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="bba69-130">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="bba69-131">Ajouter à l’envoi</span><span class="sxs-lookup"><span data-stu-id="bba69-131">Append on send</span></span>

<span data-ttu-id="bba69-132">Pour en savoir plus sur l’utilisation de la fonctionnalité Ajout à l’envoi, consultez la rubrique [implémenter Append lors de l’envoi dans votre complément Outlook](../../../outlook/append-on-send.md).</span><span class="sxs-lookup"><span data-stu-id="bba69-132">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="bba69-133">Office. Context. Mailbox. Item. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-133">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="bba69-134">Ajout d’une nouvelle fonction à l' `Body` objet qui ajoute des données à la fin du corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bba69-134">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="bba69-135">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bba69-135">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="bba69-136">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="bba69-136">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="bba69-137">Ajout d’un nouvel élément au manifeste dans lequel l' `AppendOnSend` autorisation étendue doit être incluse dans la collection des autorisations étendues.</span><span class="sxs-lookup"><span data-stu-id="bba69-137">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="bba69-138">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bba69-138">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="async-versions-of-display-apis"></a><span data-ttu-id="bba69-139">Versions Async des `display` API</span><span class="sxs-lookup"><span data-stu-id="bba69-139">Async versions of `display` APIs</span></span>

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[<span data-ttu-id="bba69-140">Office. Context. Mailbox. displayAppointmentFormAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-140">Office.context.mailbox.displayAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentformasync-itemid--options--callback-)

<span data-ttu-id="bba69-141">Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un rendez-vous existant.</span><span class="sxs-lookup"><span data-stu-id="bba69-141">Added a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="bba69-142">Il s’agit de la version asynchrone de la `displayAppointmentForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="bba69-142">This is the async version of the `displayAppointmentForm` method.</span></span>

<span data-ttu-id="bba69-143">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-143">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[<span data-ttu-id="bba69-144">Office. Context. Mailbox. displayMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-144">Office.context.mailbox.displayMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageformasync-itemid--options--callback-)

<span data-ttu-id="bba69-145">Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="bba69-145">Added a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="bba69-146">Il s’agit de la version asynchrone de la `displayMessageForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="bba69-146">This is the async version of the `displayMessageForm` method.</span></span>

<span data-ttu-id="bba69-147">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-147">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[<span data-ttu-id="bba69-148">Office. Context. Mailbox. displayNewAppointmentFormAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-148">Office.context.mailbox.displayNewAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentformasync-parameters--options--callback-)

<span data-ttu-id="bba69-149">Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bba69-149">Added a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="bba69-150">Il s’agit de la version asynchrone de la `displayNewAppointmentForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="bba69-150">This is the async version of the `displayNewAppointmentForm` method.</span></span>

<span data-ttu-id="bba69-151">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-151">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[<span data-ttu-id="bba69-152">Office. Context. Mailbox. displayNewMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-152">Office.context.mailbox.displayNewMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageformasync-parameters--options--callback-)

<span data-ttu-id="bba69-153">Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="bba69-153">Added a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="bba69-154">Il s’agit de la version asynchrone de la `displayNewMessageForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="bba69-154">This is the async version of the `displayNewMessageForm` method.</span></span>

<span data-ttu-id="bba69-155">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-155">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[<span data-ttu-id="bba69-156">Office. Context. Mailbox. Item. displayReplyAllFormAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-156">Office.context.mailbox.item.displayReplyAllFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bba69-157">Ajout d’une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre à tous » en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="bba69-157">Added a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="bba69-158">Il s’agit de la version asynchrone de la `displayReplyAllForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="bba69-158">This is the async version of the `displayReplyAllForm` method.</span></span>

<span data-ttu-id="bba69-159">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-159">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[<span data-ttu-id="bba69-160">Office. Context. Mailbox. Item. displayReplyFormAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-160">Office.context.mailbox.item.displayReplyFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bba69-161">Ajout d’une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre » en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="bba69-161">Added a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="bba69-162">Il s’agit de la version asynchrone de la `displayReplyForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="bba69-162">This is the async version of the `displayReplyForm` method.</span></span>

<span data-ttu-id="bba69-163">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-163">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="bba69-164">Activation basée sur les événements</span><span class="sxs-lookup"><span data-stu-id="bba69-164">Event-based activation</span></span>

<span data-ttu-id="bba69-165">Prise en charge supplémentaire de la fonctionnalité d’activation basée sur un événement dans les compléments Outlook. Pour en savoir plus, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="bba69-165">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="bba69-166">Point d’extension LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="bba69-166">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="bba69-167">Ajout `LaunchEvent` de la prise en charge du point d’extension au manifeste.</span><span class="sxs-lookup"><span data-stu-id="bba69-167">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="bba69-168">Il configure les fonctionnalités d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="bba69-168">It configures event-based activation functionality.</span></span>

<span data-ttu-id="bba69-169">**Disponible dans**: Outlook sur le Web (moderne, [demander un accès en aperçu](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="bba69-169">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="bba69-170">Élément de manifeste LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="bba69-170">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="bba69-171">Ajout `LaunchEvents` de l’élément à manifest.</span><span class="sxs-lookup"><span data-stu-id="bba69-171">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="bba69-172">Il prend en charge la configuration de la fonctionnalité d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="bba69-172">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="bba69-173">**Disponible dans**: Outlook sur le Web (moderne, [demander un accès en aperçu](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="bba69-173">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="bba69-174">Élément de manifeste runtimes</span><span class="sxs-lookup"><span data-stu-id="bba69-174">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="bba69-175">Ajout de la prise en charge d’Outlook à l' `Runtimes` élément de manifeste.</span><span class="sxs-lookup"><span data-stu-id="bba69-175">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="bba69-176">Il fait référence aux fichiers HTML et JavaScript nécessaires à la fonctionnalité d’activation basée sur les événements.</span><span class="sxs-lookup"><span data-stu-id="bba69-176">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="bba69-177">**Disponible dans**: Outlook sur le Web (moderne, [demander un accès en aperçu](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="bba69-177">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="bba69-178">Obtenir toutes les propriétés personnalisées</span><span class="sxs-lookup"><span data-stu-id="bba69-178">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="bba69-179">CustomProperties. getAll</span><span class="sxs-lookup"><span data-stu-id="bba69-179">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

<span data-ttu-id="bba69-180">Ajout d’une nouvelle fonction à l' `CustomProperties` objet qui obtient toutes les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="bba69-180">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="bba69-181">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne), Outlook sur Mac (connecté à l’abonnement Office 365), Outlook sur Android, Outlook sur iOS</span><span class="sxs-lookup"><span data-stu-id="bba69-181">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="bba69-182">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="bba69-182">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="bba69-183">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-183">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bba69-184">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="bba69-184">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="bba69-185">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="bba69-185">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="bba69-186">Signature de courrier électronique</span><span class="sxs-lookup"><span data-stu-id="bba69-186">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="bba69-187">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-187">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="bba69-188">Ajout d’une nouvelle fonction à l' `Body` objet qui ajoute ou remplace la signature dans le corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bba69-188">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="bba69-189">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bba69-189">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="bba69-190">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-190">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bba69-191">Ajout d’une fonction qui désactive la signature client pour la boîte aux lettres d’envoi en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bba69-191">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="bba69-192">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bba69-192">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="bba69-193">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-193">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="bba69-194">Ajout d’une nouvelle fonction qui obtient le type de composition d’un message en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bba69-194">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="bba69-195">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bba69-195">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="bba69-196">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="bba69-196">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bba69-197">Ajout d’une fonction qui vérifie si la signature client est activée sur l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bba69-197">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="bba69-198">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bba69-198">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="bba69-199">Office. MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="bba69-199">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="bba69-200">Ajout d’une nouvelle énumération `ComposeType` disponible en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bba69-200">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="bba69-201">**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bba69-201">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="bba69-202">Thème Office</span><span class="sxs-lookup"><span data-stu-id="bba69-202">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="bba69-203">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="bba69-203">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="bba69-204">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="bba69-204">Added ability to get Office theme.</span></span>

<span data-ttu-id="bba69-205">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-205">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="bba69-206">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="bba69-206">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="bba69-207">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="bba69-207">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="bba69-208">**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="bba69-208">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="bba69-209">Authentification unique (SSO)</span><span class="sxs-lookup"><span data-stu-id="bba69-209">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="bba69-210">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="bba69-210">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="bba69-211">Ajout d’un accès à `getAccessToken`, qui permet aux compléments d’[obtenir un jeton d’accès](../../../outlook/authenticate-a-user-with-an-sso-token.md) pour l’API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="bba69-211">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="bba69-212">**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (moderne), Outlook sur le web (classique)</span><span class="sxs-lookup"><span data-stu-id="bba69-212">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="bba69-213">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bba69-213">See also</span></span>

- [<span data-ttu-id="bba69-214">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="bba69-214">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="bba69-215">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="bba69-215">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="bba69-216">Prise en main</span><span class="sxs-lookup"><span data-stu-id="bba69-216">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="bba69-217">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="bba69-217">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
