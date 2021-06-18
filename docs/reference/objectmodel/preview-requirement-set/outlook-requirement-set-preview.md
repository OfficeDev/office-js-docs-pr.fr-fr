---
title: Outlook conditions requises de l’API du module de prévisualisation du add-in
description: Fonctionnalités et API actuellement en prévisualisation pour Outlook de recherche.
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: f9d8afc2b4347a8fb13f8ab98a163fb63968123f
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007761"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="b2ad5-103">Outlook conditions requises de l’API du module de prévisualisation du add-in</span><span class="sxs-lookup"><span data-stu-id="b2ad5-103">Outlook add-in API preview requirement set</span></span>

<span data-ttu-id="b2ad5-104">Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b2ad5-105">Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="b2ad5-106">Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="b2ad5-107">Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="b2ad5-108">Vous pourrez peut-être prévisualiser les fonctionnalités dans Outlook sur le web en configurant la version [ciblée sur votre Microsoft 365 client.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="b2ad5-109">Cette page indique « Configurer l’accès en aperçu » pour les fonctionnalités applicables.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="b2ad5-110">Pour d’autres fonctionnalités, vous pouvez demander l’accès aux bits d’aperçu pour Outlook sur le web à l’aide de votre compte Microsoft 365 en complétant et en envoyant [ce formulaire.](https://aka.ms/OWAPreview)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="b2ad5-111">« Demander l’accès en prévisualisation » est indiqué sur ces fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="b2ad5-112">L’ensemble de conditions requises de prévisualisation inclut toutes les fonctionnalités de l’ensemble de conditions [requises 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="b2ad5-112">The preview requirement set includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="b2ad5-113">Fonctionnalités (aperçu) :</span><span class="sxs-lookup"><span data-stu-id="b2ad5-113">Features in preview</span></span>

<span data-ttu-id="b2ad5-114">Les fonctionnalités suivantes sont disponibles en aperçu.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="b2ad5-115">Activation de complément sur des éléments protégés par la Gestion des droits de l’information (IRM)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="b2ad5-116">Les add-ins peuvent désormais être activés sur les éléments protégés par IRM.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="b2ad5-117">Pour activer cette fonctionnalité, un administrateur client doit activer le droit d’utilisation en paramètres de stratégie personnalisée Autoriser l’accès par programme `OBJMODEL` dans Office. </span><span class="sxs-lookup"><span data-stu-id="b2ad5-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="b2ad5-118">Pour plus [d’informations, voir droits d’utilisation et descriptions.](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="b2ad5-119">**Disponible dans**: Outlook sur Windows, à partir de la build 13229.10000 (connectée à Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="b2ad5-120">Propriétés de calendrier supplémentaires</span><span class="sxs-lookup"><span data-stu-id="b2ad5-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="b2ad5-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="b2ad5-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="b2ad5-122">Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée d’un rendez-vous en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="b2ad5-123">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="b2ad5-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="b2ad5-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="b2ad5-125">Ajout d’un nouvel objet qui représente la sensibilité d’un rendez-vous en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="b2ad5-126">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="b2ad5-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="b2ad5-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="b2ad5-128">Ajout d’une nouvelle propriété qui représente si un rendez-vous est un événement d’une journée.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="b2ad5-129">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="b2ad5-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="b2ad5-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="b2ad5-131">Ajout d’une nouvelle propriété qui représente la sensibilité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="b2ad5-132">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="b2ad5-133">Office. MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="b2ad5-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="b2ad5-134">Ajout d’une nouvelle enum `AppointmentSensitivityType` qui représente les options de sensibilité disponibles sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="b2ad5-135">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="b2ad5-136">Activation basée sur un événement</span><span class="sxs-lookup"><span data-stu-id="b2ad5-136">Event-based activation</span></span>

<span data-ttu-id="b2ad5-137">Cette fonctionnalité a été publiée dans [l’ensemble de conditions requises 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="b2ad5-137">This feature was released in [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="b2ad5-138">Toutefois, des événements supplémentaires sont désormais disponibles en prévisualisation.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-138">However, additional events are now available in preview.</span></span> <span data-ttu-id="b2ad5-139">Pour en savoir plus, reportez-vous aux [événements pris en charge.](../../../outlook/autolaunch.md#supported-events)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-139">To learn more, refer to [Supported events](../../../outlook/autolaunch.md#supported-events).</span></span>

<span data-ttu-id="b2ad5-140">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="b2ad5-141">Intégration avec les messages actionnables</span><span class="sxs-lookup"><span data-stu-id="b2ad5-141">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="b2ad5-142">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="b2ad5-142">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="b2ad5-143">Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="b2ad5-143">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="b2ad5-144">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-144">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="b2ad5-145">Thème Office</span><span class="sxs-lookup"><span data-stu-id="b2ad5-145">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="b2ad5-146">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="b2ad5-146">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="b2ad5-147">Ajout de la possibilité d’obtenir un thème Office.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-147">Added ability to get Office theme.</span></span>

<span data-ttu-id="b2ad5-148">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="b2ad5-149">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="b2ad5-149">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="b2ad5-150">Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-150">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="b2ad5-151">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-151">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="b2ad5-152">Données de session</span><span class="sxs-lookup"><span data-stu-id="b2ad5-152">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="b2ad5-153">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="b2ad5-153">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="b2ad5-154">Ajout d’un nouvel objet qui représente les données de session d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-154">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="b2ad5-155">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-155">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="b2ad5-156">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="b2ad5-156">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="b2ad5-157">Ajout d’une nouvelle propriété pour gérer les données de session d’un élément en mode Composition.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-157">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="b2ad5-158">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-158">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="shared-mailboxes"></a><span data-ttu-id="b2ad5-159">Boîtes aux lettres partagées</span><span class="sxs-lookup"><span data-stu-id="b2ad5-159">Shared mailboxes</span></span>

<span data-ttu-id="b2ad5-160">La prise en charge des fonctionnalités pour les dossiers partagés (autrement dit, l’accès délégué) a été publiée dans l’ensemble de conditions [requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="b2ad5-160">Feature support for shared folders (that is, delegate access) was released in [requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="b2ad5-161">Toutefois, la prise en charge des boîtes aux lettres partagées est désormais disponible en prévisualisation.</span><span class="sxs-lookup"><span data-stu-id="b2ad5-161">However, support for shared mailboxes is now available in preview.</span></span> <span data-ttu-id="b2ad5-162">Pour plus d’informations, voir [Activer les dossiers partagés et les scénarios de boîtes aux lettres partagées.](../../../outlook/delegate-access.md)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-162">To learn more, refer to [Enable shared folders and shared mailbox scenarios](../../../outlook/delegate-access.md).</span></span>

<span data-ttu-id="b2ad5-163">**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)</span><span class="sxs-lookup"><span data-stu-id="b2ad5-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="b2ad5-164">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b2ad5-164">See also</span></span>

- [<span data-ttu-id="b2ad5-165">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="b2ad5-165">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="b2ad5-166">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="b2ad5-166">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="b2ad5-167">Prise en main</span><span class="sxs-lookup"><span data-stu-id="b2ad5-167">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="b2ad5-168">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="b2ad5-168">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
