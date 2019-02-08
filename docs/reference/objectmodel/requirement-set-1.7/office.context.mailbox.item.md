---
title: Office.Context.Mailbox.Item - exigence défini 1.7
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: 6ac795d426cf80071d7b83d5e10714f4d3a6036b
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701889"
---
# <a name="item"></a><span data-ttu-id="18ec6-102">élément</span><span class="sxs-lookup"><span data-stu-id="18ec6-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="18ec6-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="18ec6-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="18ec6-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="18ec6-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-106">Requirements</span></span>

|<span data-ttu-id="18ec6-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-107">Requirement</span></span>|<span data-ttu-id="18ec6-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-110">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-110">1.0</span></span>|
|[<span data-ttu-id="18ec6-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="18ec6-112">Restricted</span></span>|
|[<span data-ttu-id="18ec6-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="18ec6-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="18ec6-115">Members and methods</span></span>

| <span data-ttu-id="18ec6-116">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-116">Member</span></span> | <span data-ttu-id="18ec6-117">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="18ec6-118">attachments</span><span class="sxs-lookup"><span data-stu-id="18ec6-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="18ec6-119">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-119">Member</span></span> |
| [<span data-ttu-id="18ec6-120">bcc</span><span class="sxs-lookup"><span data-stu-id="18ec6-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="18ec6-121">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-121">Member</span></span> |
| [<span data-ttu-id="18ec6-122">body</span><span class="sxs-lookup"><span data-stu-id="18ec6-122">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="18ec6-123">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-123">Member</span></span> |
| [<span data-ttu-id="18ec6-124">cc</span><span class="sxs-lookup"><span data-stu-id="18ec6-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="18ec6-125">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-125">Member</span></span> |
| [<span data-ttu-id="18ec6-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="18ec6-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="18ec6-127">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-127">Member</span></span> |
| [<span data-ttu-id="18ec6-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="18ec6-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="18ec6-129">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-129">Member</span></span> |
| [<span data-ttu-id="18ec6-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="18ec6-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="18ec6-131">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-131">Member</span></span> |
| [<span data-ttu-id="18ec6-132">end</span><span class="sxs-lookup"><span data-stu-id="18ec6-132">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="18ec6-133">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-133">Member</span></span> |
| [<span data-ttu-id="18ec6-134">from</span><span class="sxs-lookup"><span data-stu-id="18ec6-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="18ec6-135">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-135">Member</span></span> |
| [<span data-ttu-id="18ec6-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="18ec6-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="18ec6-137">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-137">Member</span></span> |
| [<span data-ttu-id="18ec6-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="18ec6-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="18ec6-139">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-139">Member</span></span> |
| [<span data-ttu-id="18ec6-140">itemId</span><span class="sxs-lookup"><span data-stu-id="18ec6-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="18ec6-141">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-141">Member</span></span> |
| [<span data-ttu-id="18ec6-142">itemType</span><span class="sxs-lookup"><span data-stu-id="18ec6-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="18ec6-143">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-143">Member</span></span> |
| [<span data-ttu-id="18ec6-144">location</span><span class="sxs-lookup"><span data-stu-id="18ec6-144">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="18ec6-145">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-145">Member</span></span> |
| [<span data-ttu-id="18ec6-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="18ec6-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="18ec6-147">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-147">Member</span></span> |
| [<span data-ttu-id="18ec6-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="18ec6-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="18ec6-149">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-149">Member</span></span> |
| [<span data-ttu-id="18ec6-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="18ec6-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="18ec6-151">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-151">Member</span></span> |
| [<span data-ttu-id="18ec6-152">organizer</span><span class="sxs-lookup"><span data-stu-id="18ec6-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="18ec6-153">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-153">Member</span></span> |
| [<span data-ttu-id="18ec6-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="18ec6-154">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="18ec6-155">Member</span><span class="sxs-lookup"><span data-stu-id="18ec6-155">Member</span></span> |
| [<span data-ttu-id="18ec6-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="18ec6-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="18ec6-157">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-157">Member</span></span> |
| [<span data-ttu-id="18ec6-158">sender</span><span class="sxs-lookup"><span data-stu-id="18ec6-158">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="18ec6-159">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-159">Member</span></span> |
| [<span data-ttu-id="18ec6-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="18ec6-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="18ec6-161">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-161">Member</span></span> |
| [<span data-ttu-id="18ec6-162">start</span><span class="sxs-lookup"><span data-stu-id="18ec6-162">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="18ec6-163">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-163">Member</span></span> |
| [<span data-ttu-id="18ec6-164">subject</span><span class="sxs-lookup"><span data-stu-id="18ec6-164">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="18ec6-165">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-165">Member</span></span> |
| [<span data-ttu-id="18ec6-166">to</span><span class="sxs-lookup"><span data-stu-id="18ec6-166">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="18ec6-167">Membre</span><span class="sxs-lookup"><span data-stu-id="18ec6-167">Member</span></span> |
| [<span data-ttu-id="18ec6-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="18ec6-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="18ec6-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-169">Method</span></span> |
| [<span data-ttu-id="18ec6-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="18ec6-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="18ec6-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-171">Method</span></span> |
| [<span data-ttu-id="18ec6-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="18ec6-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="18ec6-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-173">Method</span></span> |
| [<span data-ttu-id="18ec6-174">close</span><span class="sxs-lookup"><span data-stu-id="18ec6-174">close</span></span>](#close) | <span data-ttu-id="18ec6-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-175">Method</span></span> |
| [<span data-ttu-id="18ec6-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="18ec6-176">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="18ec6-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-177">Method</span></span> |
| [<span data-ttu-id="18ec6-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="18ec6-178">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="18ec6-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-179">Method</span></span> |
| [<span data-ttu-id="18ec6-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="18ec6-180">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="18ec6-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-181">Method</span></span> |
| [<span data-ttu-id="18ec6-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="18ec6-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="18ec6-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-183">Method</span></span> |
| [<span data-ttu-id="18ec6-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="18ec6-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="18ec6-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-185">Method</span></span> |
| [<span data-ttu-id="18ec6-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="18ec6-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="18ec6-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-187">Method</span></span> |
| [<span data-ttu-id="18ec6-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="18ec6-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="18ec6-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-189">Method</span></span> |
| [<span data-ttu-id="18ec6-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="18ec6-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="18ec6-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-191">Method</span></span> |
| [<span data-ttu-id="18ec6-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="18ec6-192">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="18ec6-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-193">Method</span></span> |
| [<span data-ttu-id="18ec6-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="18ec6-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="18ec6-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-195">Method</span></span> |
| [<span data-ttu-id="18ec6-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="18ec6-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="18ec6-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-197">Method</span></span> |
| [<span data-ttu-id="18ec6-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="18ec6-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="18ec6-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-199">Method</span></span> |
| [<span data-ttu-id="18ec6-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="18ec6-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="18ec6-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-201">Method</span></span> |
| [<span data-ttu-id="18ec6-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="18ec6-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="18ec6-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-203">Method</span></span> |
| [<span data-ttu-id="18ec6-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="18ec6-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="18ec6-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="18ec6-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="18ec6-206">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-206">Example</span></span>

<span data-ttu-id="18ec6-207">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="18ec6-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="18ec6-208">Membres</span><span class="sxs-lookup"><span data-stu-id="18ec6-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="18ec6-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="18ec6-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="18ec6-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-212">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="18ec6-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="18ec6-213">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="18ec6-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-214">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-214">Type:</span></span>

*   <span data-ttu-id="18ec6-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="18ec6-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-216">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-216">Requirements</span></span>

|<span data-ttu-id="18ec6-217">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-217">Requirement</span></span>|<span data-ttu-id="18ec6-218">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-219">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-220">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-220">1.0</span></span>|
|[<span data-ttu-id="18ec6-221">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-222">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-224">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-225">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-225">Example</span></span>

<span data-ttu-id="18ec6-226">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="18ec6-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="18ec6-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18ec6-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="18ec6-228">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="18ec6-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="18ec6-229">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-230">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-230">Type:</span></span>

*   [<span data-ttu-id="18ec6-231">Destinataires</span><span class="sxs-lookup"><span data-stu-id="18ec6-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="18ec6-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-232">Requirements</span></span>

|<span data-ttu-id="18ec6-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-233">Requirement</span></span>|<span data-ttu-id="18ec6-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-236">1.1</span><span class="sxs-lookup"><span data-stu-id="18ec6-236">1.1</span></span>|
|[<span data-ttu-id="18ec6-237">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-238">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-240">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-241">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="18ec6-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="18ec6-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="18ec6-243">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-244">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-244">Type:</span></span>

*   [<span data-ttu-id="18ec6-245">Corps</span><span class="sxs-lookup"><span data-stu-id="18ec6-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="18ec6-246">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-246">Requirements</span></span>

|<span data-ttu-id="18ec6-247">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-247">Requirement</span></span>|<span data-ttu-id="18ec6-248">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-249">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-250">1.1</span><span class="sxs-lookup"><span data-stu-id="18ec6-250">1.1</span></span>|
|[<span data-ttu-id="18ec6-251">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-252">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-253">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-254">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-254">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="18ec6-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18ec6-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="18ec6-256">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="18ec6-256">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="18ec6-257">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="18ec6-257">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-258">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-258">Read mode</span></span>

<span data-ttu-id="18ec6-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-261">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-261">Compose mode</span></span>

<span data-ttu-id="18ec6-262">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="18ec6-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-263">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-263">Type:</span></span>

*   <span data-ttu-id="18ec6-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18ec6-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-265">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-265">Requirements</span></span>

|<span data-ttu-id="18ec6-266">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-266">Requirement</span></span>|<span data-ttu-id="18ec6-267">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-268">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-269">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-269">1.0</span></span>|
|[<span data-ttu-id="18ec6-270">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-271">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-272">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-273">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-273">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-274">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-274">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="18ec6-275">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="18ec6-275">(nullable) conversationId :String</span></span>

<span data-ttu-id="18ec6-276">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="18ec6-276">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="18ec6-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="18ec6-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-281">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-281">Type:</span></span>

*   <span data-ttu-id="18ec6-282">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-282">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-283">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-283">Requirements</span></span>

|<span data-ttu-id="18ec6-284">Requirement</span><span class="sxs-lookup"><span data-stu-id="18ec6-284">Requirement</span></span>|<span data-ttu-id="18ec6-285">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-286">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-287">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-287">1.0</span></span>|
|[<span data-ttu-id="18ec6-288">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-288">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-289">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-290">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-290">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-291">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-291">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="18ec6-292">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="18ec6-292">dateTimeCreated :Date</span></span>

<span data-ttu-id="18ec6-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-295">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-295">Type:</span></span>

*   <span data-ttu-id="18ec6-296">Date</span><span class="sxs-lookup"><span data-stu-id="18ec6-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-297">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-297">Requirements</span></span>

|<span data-ttu-id="18ec6-298">Requirement</span><span class="sxs-lookup"><span data-stu-id="18ec6-298">Requirement</span></span>|<span data-ttu-id="18ec6-299">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-300">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-301">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-301">1.0</span></span>|
|[<span data-ttu-id="18ec6-302">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-303">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-304">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-305">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-306">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-306">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="18ec6-307">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="18ec6-307">dateTimeModified :Date</span></span>

<span data-ttu-id="18ec6-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-310">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-310">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-311">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-311">Type:</span></span>

*   <span data-ttu-id="18ec6-312">Date</span><span class="sxs-lookup"><span data-stu-id="18ec6-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-313">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-313">Requirements</span></span>

|<span data-ttu-id="18ec6-314">Requirement</span><span class="sxs-lookup"><span data-stu-id="18ec6-314">Requirement</span></span>|<span data-ttu-id="18ec6-315">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-316">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-317">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-317">1.0</span></span>|
|[<span data-ttu-id="18ec6-318">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-318">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-319">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-320">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-320">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-321">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-322">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-322">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="18ec6-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="18ec6-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="18ec6-324">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="18ec6-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-327">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-327">Read mode</span></span>

<span data-ttu-id="18ec6-328">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-328">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-329">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-329">Compose mode</span></span>

<span data-ttu-id="18ec6-330">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="18ec6-331">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="18ec6-331">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-332">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-332">Type:</span></span>

*   <span data-ttu-id="18ec6-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="18ec6-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-334">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-334">Requirements</span></span>

|<span data-ttu-id="18ec6-335">Requirement</span><span class="sxs-lookup"><span data-stu-id="18ec6-335">Requirement</span></span>|<span data-ttu-id="18ec6-336">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-337">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-338">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-338">1.0</span></span>|
|[<span data-ttu-id="18ec6-339">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-340">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-341">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-342">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-343">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-343">Example</span></span>

<span data-ttu-id="18ec6-344">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-344">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="18ec6-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="18ec6-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="18ec6-346">Permet d’obtenir l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="18ec6-346">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="18ec6-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-349">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-350">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-350">Read mode</span></span>

<span data-ttu-id="18ec6-351">La propriété `from` renvoie un objet `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-351">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-352">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-352">Compose mode</span></span>

<span data-ttu-id="18ec6-353">La propriété `from` renvoie un objet `From` qui fournit une méthode pour obtenir la valeur from.</span><span class="sxs-lookup"><span data-stu-id="18ec6-353">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="18ec6-354">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-354">Type:</span></span>

*   <span data-ttu-id="18ec6-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="18ec6-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-356">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-356">Requirements</span></span>

|<span data-ttu-id="18ec6-357">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-357">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="18ec6-358">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-359">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-359">1.0</span></span>|<span data-ttu-id="18ec6-360">1.7</span><span class="sxs-lookup"><span data-stu-id="18ec6-360">1.7</span></span>|
|[<span data-ttu-id="18ec6-361">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-362">ReadItem</span></span>|<span data-ttu-id="18ec6-363">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-363">ReadWriteItem</span></span>|
|[<span data-ttu-id="18ec6-364">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-365">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-365">Read</span></span>|<span data-ttu-id="18ec6-366">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-366">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="18ec6-367">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="18ec6-367">internetMessageId :String</span></span>

<span data-ttu-id="18ec6-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-370">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-370">Type:</span></span>

*   <span data-ttu-id="18ec6-371">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-372">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-372">Requirements</span></span>

|<span data-ttu-id="18ec6-373">Requirement</span><span class="sxs-lookup"><span data-stu-id="18ec6-373">Requirement</span></span>|<span data-ttu-id="18ec6-374">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-375">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-376">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-376">1.0</span></span>|
|[<span data-ttu-id="18ec6-377">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-378">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-379">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-380">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-381">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-381">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="18ec6-382">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="18ec6-382">itemClass :String</span></span>

<span data-ttu-id="18ec6-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="18ec6-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="18ec6-387">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-387">Type</span></span>|<span data-ttu-id="18ec6-388">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-388">Description</span></span>|<span data-ttu-id="18ec6-389">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="18ec6-389">item class</span></span>|
|---|---|---|
|<span data-ttu-id="18ec6-390">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="18ec6-390">Appointment items</span></span>|<span data-ttu-id="18ec6-391">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-391">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="18ec6-392">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="18ec6-392">Message items</span></span>|<span data-ttu-id="18ec6-393">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="18ec6-393">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="18ec6-394">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-394">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-395">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-395">Type:</span></span>

*   <span data-ttu-id="18ec6-396">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-397">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-397">Requirements</span></span>

|<span data-ttu-id="18ec6-398">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-398">Requirement</span></span>|<span data-ttu-id="18ec6-399">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-400">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-401">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-401">1.0</span></span>|
|[<span data-ttu-id="18ec6-402">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-403">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-404">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-405">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-406">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-406">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="18ec6-407">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="18ec6-407">(nullable) itemId :String</span></span>

<span data-ttu-id="18ec6-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-410">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="18ec6-410">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="18ec6-411">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="18ec6-411">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="18ec6-412">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="18ec6-412">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="18ec6-413">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="18ec6-413">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="18ec6-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-416">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-416">Type:</span></span>

*   <span data-ttu-id="18ec6-417">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-418">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-418">Requirements</span></span>

|<span data-ttu-id="18ec6-419">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-419">Requirement</span></span>|<span data-ttu-id="18ec6-420">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-421">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-422">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-422">1.0</span></span>|
|[<span data-ttu-id="18ec6-423">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-424">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-425">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-426">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-427">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-427">Example</span></span>

<span data-ttu-id="18ec6-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="18ec6-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="18ec6-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="18ec6-431">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="18ec6-431">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="18ec6-432">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-432">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-433">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-433">Type:</span></span>

*   [<span data-ttu-id="18ec6-434">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="18ec6-434">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="18ec6-435">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-435">Requirements</span></span>

|<span data-ttu-id="18ec6-436">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-436">Requirement</span></span>|<span data-ttu-id="18ec6-437">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-438">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-439">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-439">1.0</span></span>|
|[<span data-ttu-id="18ec6-440">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-441">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-442">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-443">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-444">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-444">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="18ec6-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="18ec6-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="18ec6-446">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-446">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-447">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-447">Read mode</span></span>

<span data-ttu-id="18ec6-448">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-448">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-449">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-449">Compose mode</span></span>

<span data-ttu-id="18ec6-450">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-450">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-451">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-451">Type:</span></span>

*   <span data-ttu-id="18ec6-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="18ec6-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-453">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-453">Requirements</span></span>

|<span data-ttu-id="18ec6-454">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-454">Requirement</span></span>|<span data-ttu-id="18ec6-455">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-456">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-457">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-457">1.0</span></span>|
|[<span data-ttu-id="18ec6-458">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-459">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-460">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-461">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-461">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-462">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-462">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="18ec6-463">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="18ec6-463">normalizedSubject :String</span></span>

<span data-ttu-id="18ec6-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="18ec6-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="18ec6-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-468">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-468">Type:</span></span>

*   <span data-ttu-id="18ec6-469">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-469">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-470">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-470">Requirements</span></span>

|<span data-ttu-id="18ec6-471">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-471">Requirement</span></span>|<span data-ttu-id="18ec6-472">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-473">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-473">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-474">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-474">1.0</span></span>|
|[<span data-ttu-id="18ec6-475">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-476">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-477">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-478">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-478">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-479">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-479">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="18ec6-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="18ec6-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="18ec6-481">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-481">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-482">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-482">Type:</span></span>

*   [<span data-ttu-id="18ec6-483">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="18ec6-483">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="18ec6-484">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-484">Requirements</span></span>

|<span data-ttu-id="18ec6-485">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-485">Requirement</span></span>|<span data-ttu-id="18ec6-486">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-487">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-488">1.3</span><span class="sxs-lookup"><span data-stu-id="18ec6-488">1.3</span></span>|
|[<span data-ttu-id="18ec6-489">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-490">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-491">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-492">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-492">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="18ec6-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18ec6-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="18ec6-494">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-494">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="18ec6-495">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="18ec6-495">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-496">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-496">Read mode</span></span>

<span data-ttu-id="18ec6-497">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="18ec6-497">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-498">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-498">Compose mode</span></span>

<span data-ttu-id="18ec6-499">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="18ec6-499">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-500">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-500">Type:</span></span>

*   <span data-ttu-id="18ec6-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18ec6-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-502">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-502">Requirements</span></span>

|<span data-ttu-id="18ec6-503">Requirement</span><span class="sxs-lookup"><span data-stu-id="18ec6-503">Requirement</span></span>|<span data-ttu-id="18ec6-504">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-505">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-506">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-506">1.0</span></span>|
|[<span data-ttu-id="18ec6-507">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-508">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-509">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-510">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-510">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-511">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-511">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="18ec6-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="18ec6-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="18ec6-513">Permet d’obtenir l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="18ec6-513">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-514">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-514">Read mode</span></span>

<span data-ttu-id="18ec6-515">La propriété `organizer` renvoie un objet [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="18ec6-515">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-516">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-516">Compose mode</span></span>

<span data-ttu-id="18ec6-517">La propriété `organizer` renvoie un objet [Organizer](/javascript/api/outlook_1_7/office.organizer) qui fournit une méthode pour obtenir la valeur organizer.</span><span class="sxs-lookup"><span data-stu-id="18ec6-517">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-518">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-518">Type:</span></span>

*   <span data-ttu-id="18ec6-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="18ec6-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-520">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-520">Requirements</span></span>

|<span data-ttu-id="18ec6-521">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-521">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="18ec6-522">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-523">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-523">1.0</span></span>|<span data-ttu-id="18ec6-524">1.7</span><span class="sxs-lookup"><span data-stu-id="18ec6-524">1.7</span></span>|
|[<span data-ttu-id="18ec6-525">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-525">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-526">ReadItem</span></span>|<span data-ttu-id="18ec6-527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-527">ReadWriteItem</span></span>|
|[<span data-ttu-id="18ec6-528">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-529">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-529">Read</span></span>|<span data-ttu-id="18ec6-530">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-530">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-531">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-531">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="18ec6-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="18ec6-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="18ec6-533">Permet d’obtenir ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="18ec6-534">Permet d’obtenir la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="18ec6-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="18ec6-535">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="18ec6-536">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="18ec6-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="18ec6-537">La propriété `recurrence` renvoie un objet [périodicité](/javascript/api/outlook_1_7/office.recurrence) pour des demandes de réunions ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="18ec6-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="18ec6-538">La valeur `null` est renvoyée pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="18ec6-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="18ec6-539">La valeur `undefined` est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="18ec6-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="18ec6-540">Remarque : les demandes de réunion ont une valeur `itemClass` d’IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="18ec6-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="18ec6-541">Remarque : si l’objet de périodicité est `null`, cela indique que l’objet est un rendez-vous unique ou une demande de réunion de rendez-vous unique, et NON une partie d’une série.</span><span class="sxs-lookup"><span data-stu-id="18ec6-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-542">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-542">Type:</span></span>

* [<span data-ttu-id="18ec6-543">Recurrence</span><span class="sxs-lookup"><span data-stu-id="18ec6-543">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="18ec6-544">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-544">Requirement</span></span>|<span data-ttu-id="18ec6-545">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-546">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-547">1.7</span><span class="sxs-lookup"><span data-stu-id="18ec6-547">1.7</span></span>|
|[<span data-ttu-id="18ec6-548">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-548">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-549">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-550">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-550">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-551">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-551">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="18ec6-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18ec6-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="18ec6-553">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-553">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="18ec6-554">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="18ec6-554">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-555">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-555">Read mode</span></span>

<span data-ttu-id="18ec6-556">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="18ec6-556">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-557">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-557">Compose mode</span></span>

<span data-ttu-id="18ec6-558">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="18ec6-558">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-559">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-559">Type:</span></span>

*   <span data-ttu-id="18ec6-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18ec6-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-561">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-561">Requirements</span></span>

|<span data-ttu-id="18ec6-562">Requirement</span><span class="sxs-lookup"><span data-stu-id="18ec6-562">Requirement</span></span>|<span data-ttu-id="18ec6-563">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-564">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-565">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-565">1.0</span></span>|
|[<span data-ttu-id="18ec6-566">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-567">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-569">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-570">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="18ec6-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="18ec6-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="18ec6-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="18ec6-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-576">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-576">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-577">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-577">Type:</span></span>

*   [<span data-ttu-id="18ec6-578">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="18ec6-578">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="18ec6-579">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-579">Requirements</span></span>

|<span data-ttu-id="18ec6-580">Requirement</span><span class="sxs-lookup"><span data-stu-id="18ec6-580">Requirement</span></span>|<span data-ttu-id="18ec6-581">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-582">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-583">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-583">1.0</span></span>|
|[<span data-ttu-id="18ec6-584">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-585">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-586">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-587">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-587">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-588">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-588">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="18ec6-589">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="18ec6-589">(nullable) seriesId :String</span></span>

<span data-ttu-id="18ec6-590">Permet d’obtenir l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="18ec6-590">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="18ec6-591">Dans OWA et Outlook, `seriesId` renvoie l’identificateur de services web Exchange (EWS) de l’élément (series) parent auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="18ec6-591">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="18ec6-592">Dans iOS et Android, `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="18ec6-592">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-593">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="18ec6-593">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="18ec6-594">La propriété `seriesId` n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="18ec6-594">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="18ec6-595">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="18ec6-595">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="18ec6-596">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="18ec6-596">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="18ec6-597">La propriété `seriesId` renvoie `null` pour les éléments qui n’ont pas d’élément parent, tels que des rendez-vous uniques, des éléments de séries ou des demandes de réunion, et renvoie `undefined` pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="18ec6-597">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-598">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-598">Type:</span></span>

* <span data-ttu-id="18ec6-599">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-599">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-600">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-600">Requirements</span></span>

|<span data-ttu-id="18ec6-601">Requirement</span><span class="sxs-lookup"><span data-stu-id="18ec6-601">Requirement</span></span>|<span data-ttu-id="18ec6-602">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-603">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-604">1.7</span><span class="sxs-lookup"><span data-stu-id="18ec6-604">1.7</span></span>|
|[<span data-ttu-id="18ec6-605">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-606">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-607">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-608">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-609">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-609">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="18ec6-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="18ec6-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="18ec6-611">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-611">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="18ec6-p130">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-614">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-614">Read mode</span></span>

<span data-ttu-id="18ec6-615">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-615">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-616">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-616">Compose mode</span></span>

<span data-ttu-id="18ec6-617">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-617">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="18ec6-618">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="18ec6-618">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-619">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-619">Type:</span></span>

*   <span data-ttu-id="18ec6-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="18ec6-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-621">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-621">Requirements</span></span>

|<span data-ttu-id="18ec6-622">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-622">Requirement</span></span>|<span data-ttu-id="18ec6-623">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-624">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-625">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-625">1.0</span></span>|
|[<span data-ttu-id="18ec6-626">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-627">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-627">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-628">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-629">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-629">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-630">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-630">Example</span></span>

<span data-ttu-id="18ec6-631">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-631">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="18ec6-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="18ec6-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="18ec6-633">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-633">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="18ec6-634">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="18ec6-634">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-635">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-635">Read mode</span></span>

<span data-ttu-id="18ec6-p131">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-638">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-638">Compose mode</span></span>

<span data-ttu-id="18ec6-639">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="18ec6-639">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="18ec6-640">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-640">Type:</span></span>

*   <span data-ttu-id="18ec6-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="18ec6-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-642">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-642">Requirements</span></span>

|<span data-ttu-id="18ec6-643">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-643">Requirement</span></span>|<span data-ttu-id="18ec6-644">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-645">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-646">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-646">1.0</span></span>|
|[<span data-ttu-id="18ec6-647">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-648">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-649">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-650">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-650">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="18ec6-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18ec6-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="18ec6-652">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="18ec6-652">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="18ec6-653">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="18ec6-653">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18ec6-654">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-654">Read mode</span></span>

<span data-ttu-id="18ec6-p133">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18ec6-657">Mode composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-657">Compose mode</span></span>

<span data-ttu-id="18ec6-658">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="18ec6-658">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="18ec6-659">Type :</span><span class="sxs-lookup"><span data-stu-id="18ec6-659">Type:</span></span>

*   <span data-ttu-id="18ec6-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18ec6-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-661">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-661">Requirements</span></span>

|<span data-ttu-id="18ec6-662">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-662">Requirement</span></span>|<span data-ttu-id="18ec6-663">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-663">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-664">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-664">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-665">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-665">1.0</span></span>|
|[<span data-ttu-id="18ec6-666">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-666">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-667">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-667">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-668">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-668">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-669">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-669">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-670">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-670">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="18ec6-671">Méthodes</span><span class="sxs-lookup"><span data-stu-id="18ec6-671">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="18ec6-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18ec6-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="18ec6-673">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="18ec6-673">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="18ec6-674">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="18ec6-674">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="18ec6-675">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="18ec6-675">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-676">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-676">Parameters:</span></span>
|<span data-ttu-id="18ec6-677">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-677">Name</span></span>|<span data-ttu-id="18ec6-678">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-678">Type</span></span>|<span data-ttu-id="18ec6-679">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-679">Attributes</span></span>|<span data-ttu-id="18ec6-680">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-680">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="18ec6-681">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-681">String</span></span>||<span data-ttu-id="18ec6-p134">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="18ec6-684">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-684">String</span></span>||<span data-ttu-id="18ec6-p135">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="18ec6-687">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-687">Object</span></span>|<span data-ttu-id="18ec6-688">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-688">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-689">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-689">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18ec6-690">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-690">Object</span></span>|<span data-ttu-id="18ec6-691">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-691">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-692">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-692">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="18ec6-693">Boolean</span><span class="sxs-lookup"><span data-stu-id="18ec6-693">Boolean</span></span>|<span data-ttu-id="18ec6-694">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-694">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-695">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-695">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="18ec6-696">fonction</span><span class="sxs-lookup"><span data-stu-id="18ec6-696">function</span></span>|<span data-ttu-id="18ec6-697">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-697">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-698">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-698">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="18ec6-699">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-699">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="18ec6-700">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="18ec6-700">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="18ec6-701">Erreurs</span><span class="sxs-lookup"><span data-stu-id="18ec6-701">Errors</span></span>

|<span data-ttu-id="18ec6-702">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="18ec6-702">Error code</span></span>|<span data-ttu-id="18ec6-703">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-703">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="18ec6-704">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="18ec6-704">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="18ec6-705">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="18ec6-705">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="18ec6-706">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-706">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-707">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-707">Requirements</span></span>

|<span data-ttu-id="18ec6-708">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-708">Requirement</span></span>|<span data-ttu-id="18ec6-709">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-710">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-711">1.1</span><span class="sxs-lookup"><span data-stu-id="18ec6-711">1.1</span></span>|
|[<span data-ttu-id="18ec6-712">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-713">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-713">ReadWriteItem</span></span>|
|[<span data-ttu-id="18ec6-714">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-715">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-715">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="18ec6-716">Exemples</span><span class="sxs-lookup"><span data-stu-id="18ec6-716">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="18ec6-717">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="18ec6-717">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="18ec6-718">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18ec6-718">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="18ec6-719">Ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="18ec6-719">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="18ec6-720">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-720">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-721">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-721">Parameters:</span></span>

| <span data-ttu-id="18ec6-722">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-722">Name</span></span> | <span data-ttu-id="18ec6-723">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-723">Type</span></span> | <span data-ttu-id="18ec6-724">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-724">Attributes</span></span> | <span data-ttu-id="18ec6-725">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-725">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="18ec6-726">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="18ec6-726">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="18ec6-727">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="18ec6-727">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="18ec6-728">Fonction</span><span class="sxs-lookup"><span data-stu-id="18ec6-728">Function</span></span> || <span data-ttu-id="18ec6-p136">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="18ec6-732">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-732">Object</span></span> | <span data-ttu-id="18ec6-733">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-733">&lt;optional&gt;</span></span> | <span data-ttu-id="18ec6-734">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-734">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="18ec6-735">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-735">Object</span></span> | <span data-ttu-id="18ec6-736">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-736">&lt;optional&gt;</span></span> | <span data-ttu-id="18ec6-737">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-737">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="18ec6-738">fonction</span><span class="sxs-lookup"><span data-stu-id="18ec6-738">function</span></span>| <span data-ttu-id="18ec6-739">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-739">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-740">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-740">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-741">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-741">Requirements</span></span>

|<span data-ttu-id="18ec6-742">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-742">Requirement</span></span>| <span data-ttu-id="18ec6-743">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-743">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-744">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-744">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18ec6-745">1.7</span><span class="sxs-lookup"><span data-stu-id="18ec6-745">1.7</span></span> |
|[<span data-ttu-id="18ec6-746">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-746">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18ec6-747">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-747">ReadItem</span></span> |
|[<span data-ttu-id="18ec6-748">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-748">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18ec6-749">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-749">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="18ec6-750">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-750">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="18ec6-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18ec6-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="18ec6-752">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-752">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="18ec6-p137">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="18ec6-756">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="18ec6-756">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="18ec6-757">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="18ec6-757">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-758">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-758">Parameters:</span></span>

|<span data-ttu-id="18ec6-759">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-759">Name</span></span>|<span data-ttu-id="18ec6-760">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-760">Type</span></span>|<span data-ttu-id="18ec6-761">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-761">Attributes</span></span>|<span data-ttu-id="18ec6-762">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-762">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="18ec6-763">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-763">String</span></span>||<span data-ttu-id="18ec6-p138">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="18ec6-766">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-766">String</span></span>||<span data-ttu-id="18ec6-p139">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="18ec6-769">Object</span><span class="sxs-lookup"><span data-stu-id="18ec6-769">Object</span></span>|<span data-ttu-id="18ec6-770">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-770">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-771">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-771">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18ec6-772">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-772">Object</span></span>|<span data-ttu-id="18ec6-773">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-773">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-774">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-774">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18ec6-775">fonction</span><span class="sxs-lookup"><span data-stu-id="18ec6-775">function</span></span>|<span data-ttu-id="18ec6-776">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-776">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-777">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="18ec6-778">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-778">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="18ec6-779">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="18ec6-779">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="18ec6-780">Erreurs</span><span class="sxs-lookup"><span data-stu-id="18ec6-780">Errors</span></span>

|<span data-ttu-id="18ec6-781">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="18ec6-781">Error code</span></span>|<span data-ttu-id="18ec6-782">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-782">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="18ec6-783">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-783">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-784">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-784">Requirements</span></span>

|<span data-ttu-id="18ec6-785">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-785">Requirement</span></span>|<span data-ttu-id="18ec6-786">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-786">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-787">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-787">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-788">1.1</span><span class="sxs-lookup"><span data-stu-id="18ec6-788">1.1</span></span>|
|[<span data-ttu-id="18ec6-789">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-789">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-790">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-790">ReadWriteItem</span></span>|
|[<span data-ttu-id="18ec6-791">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-791">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-792">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-792">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-793">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-793">Example</span></span>

<span data-ttu-id="18ec6-794">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-794">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="18ec6-795">close()</span><span class="sxs-lookup"><span data-stu-id="18ec6-795">close()</span></span>

<span data-ttu-id="18ec6-796">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="18ec6-796">Closes the current item that is being composed.</span></span>

<span data-ttu-id="18ec6-p140">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-799">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-799">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="18ec6-800">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="18ec6-800">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-801">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-801">Requirements</span></span>

|<span data-ttu-id="18ec6-802">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-802">Requirement</span></span>|<span data-ttu-id="18ec6-803">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-803">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-804">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-804">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-805">1.3</span><span class="sxs-lookup"><span data-stu-id="18ec6-805">1.3</span></span>|
|[<span data-ttu-id="18ec6-806">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-806">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-807">Restreinte</span><span class="sxs-lookup"><span data-stu-id="18ec6-807">Restricted</span></span>|
|[<span data-ttu-id="18ec6-808">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-808">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-809">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-809">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="18ec6-810">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="18ec6-810">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="18ec6-811">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="18ec6-811">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-812">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-812">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18ec6-813">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-813">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="18ec6-814">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="18ec6-814">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="18ec6-p141">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-818">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-818">Parameters:</span></span>

|<span data-ttu-id="18ec6-819">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-819">Name</span></span>|<span data-ttu-id="18ec6-820">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-820">Type</span></span>|<span data-ttu-id="18ec6-821">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-821">Attributes</span></span>|<span data-ttu-id="18ec6-822">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-822">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="18ec6-823">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="18ec6-823">String &#124; Object</span></span>||<span data-ttu-id="18ec6-p142">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="18ec6-826">**OU**</span><span class="sxs-lookup"><span data-stu-id="18ec6-826">**OR**</span></span><br/><span data-ttu-id="18ec6-p143">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="18ec6-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="18ec6-829">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-829">String</span></span>|<span data-ttu-id="18ec6-830">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-830">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="18ec6-833">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-833">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="18ec6-834">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-834">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-835">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-835">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="18ec6-836">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-836">String</span></span>||<span data-ttu-id="18ec6-p145">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="18ec6-839">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-839">String</span></span>||<span data-ttu-id="18ec6-840">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="18ec6-840">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="18ec6-841">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-841">String</span></span>||<span data-ttu-id="18ec6-p146">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="18ec6-844">Boolean</span><span class="sxs-lookup"><span data-stu-id="18ec6-844">Boolean</span></span>||<span data-ttu-id="18ec6-p147">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="18ec6-847">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-847">String</span></span>||<span data-ttu-id="18ec6-p148">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="18ec6-851">function</span><span class="sxs-lookup"><span data-stu-id="18ec6-851">function</span></span>|<span data-ttu-id="18ec6-852">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-852">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-853">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-853">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-854">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-854">Requirements</span></span>

|<span data-ttu-id="18ec6-855">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-855">Requirement</span></span>|<span data-ttu-id="18ec6-856">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-857">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-858">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-858">1.0</span></span>|
|[<span data-ttu-id="18ec6-859">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-859">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-860">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-861">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-861">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-862">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-862">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="18ec6-863">Exemples</span><span class="sxs-lookup"><span data-stu-id="18ec6-863">Examples</span></span>

<span data-ttu-id="18ec6-864">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-864">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="18ec6-865">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="18ec6-865">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="18ec6-866">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="18ec6-866">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="18ec6-867">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="18ec6-867">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="18ec6-868">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-868">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="18ec6-869">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-869">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="18ec6-870">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="18ec6-870">displayReplyForm(formData)</span></span>

<span data-ttu-id="18ec6-871">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="18ec6-871">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-872">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-872">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18ec6-873">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-873">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="18ec6-874">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="18ec6-874">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="18ec6-p149">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-878">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-878">Parameters:</span></span>

|<span data-ttu-id="18ec6-879">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-879">Name</span></span>|<span data-ttu-id="18ec6-880">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-880">Type</span></span>|<span data-ttu-id="18ec6-881">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-881">Attributes</span></span>|<span data-ttu-id="18ec6-882">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-882">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="18ec6-883">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="18ec6-883">String &#124; Object</span></span>||<span data-ttu-id="18ec6-p150">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="18ec6-886">**OU**</span><span class="sxs-lookup"><span data-stu-id="18ec6-886">**OR**</span></span><br/><span data-ttu-id="18ec6-p151">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="18ec6-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="18ec6-889">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-889">String</span></span>|<span data-ttu-id="18ec6-890">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-890">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-p152">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="18ec6-893">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-893">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="18ec6-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-894">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-895">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-895">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="18ec6-896">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-896">String</span></span>||<span data-ttu-id="18ec6-p153">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="18ec6-899">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-899">String</span></span>||<span data-ttu-id="18ec6-900">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="18ec6-900">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="18ec6-901">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-901">String</span></span>||<span data-ttu-id="18ec6-p154">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="18ec6-904">Boolean</span><span class="sxs-lookup"><span data-stu-id="18ec6-904">Boolean</span></span>||<span data-ttu-id="18ec6-p155">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="18ec6-907">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-907">String</span></span>||<span data-ttu-id="18ec6-p156">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="18ec6-911">function</span><span class="sxs-lookup"><span data-stu-id="18ec6-911">function</span></span>|<span data-ttu-id="18ec6-912">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-912">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-913">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-913">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-914">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-914">Requirements</span></span>

|<span data-ttu-id="18ec6-915">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-915">Requirement</span></span>|<span data-ttu-id="18ec6-916">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-917">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-918">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-918">1.0</span></span>|
|[<span data-ttu-id="18ec6-919">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-919">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-920">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-921">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-921">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-922">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-922">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="18ec6-923">Exemples</span><span class="sxs-lookup"><span data-stu-id="18ec6-923">Examples</span></span>

<span data-ttu-id="18ec6-924">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-924">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="18ec6-925">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="18ec6-925">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="18ec6-926">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="18ec6-926">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="18ec6-927">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="18ec6-927">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="18ec6-928">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-928">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="18ec6-929">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-929">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="18ec6-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="18ec6-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="18ec6-931">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="18ec6-931">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-932">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-932">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-933">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-933">Requirements</span></span>

|<span data-ttu-id="18ec6-934">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-934">Requirement</span></span>|<span data-ttu-id="18ec6-935">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-936">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-937">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-937">1.0</span></span>|
|[<span data-ttu-id="18ec6-938">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-939">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-940">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-941">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-941">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18ec6-942">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="18ec6-942">Returns:</span></span>

<span data-ttu-id="18ec6-943">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="18ec6-943">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="18ec6-944">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-944">Example</span></span>

<span data-ttu-id="18ec6-945">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="18ec6-945">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="18ec6-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="18ec6-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="18ec6-947">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="18ec6-947">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-948">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-949">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-949">Parameters:</span></span>

|<span data-ttu-id="18ec6-950">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-950">Name</span></span>|<span data-ttu-id="18ec6-951">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-951">Type</span></span>|<span data-ttu-id="18ec6-952">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-952">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="18ec6-953">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="18ec6-953">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="18ec6-954">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="18ec6-954">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-955">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-955">Requirements</span></span>

|<span data-ttu-id="18ec6-956">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-956">Requirement</span></span>|<span data-ttu-id="18ec6-957">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-958">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-959">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-959">1.0</span></span>|
|[<span data-ttu-id="18ec6-960">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-961">Restreinte</span><span class="sxs-lookup"><span data-stu-id="18ec6-961">Restricted</span></span>|
|[<span data-ttu-id="18ec6-962">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-963">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18ec6-964">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="18ec6-964">Returns:</span></span>

<span data-ttu-id="18ec6-965">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="18ec6-965">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="18ec6-966">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="18ec6-966">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="18ec6-967">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-967">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="18ec6-968">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="18ec6-968">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="18ec6-969">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="18ec6-969">Value of `entityType`</span></span>|<span data-ttu-id="18ec6-970">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="18ec6-970">Type of objects in returned array</span></span>|<span data-ttu-id="18ec6-971">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="18ec6-971">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="18ec6-972">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-972">String</span></span>|<span data-ttu-id="18ec6-973">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="18ec6-973">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="18ec6-974">Contact</span><span class="sxs-lookup"><span data-stu-id="18ec6-974">Contact</span></span>|<span data-ttu-id="18ec6-975">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="18ec6-975">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="18ec6-976">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-976">String</span></span>|<span data-ttu-id="18ec6-977">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="18ec6-977">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="18ec6-978">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="18ec6-978">MeetingSuggestion</span></span>|<span data-ttu-id="18ec6-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="18ec6-979">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="18ec6-980">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="18ec6-980">PhoneNumber</span></span>|<span data-ttu-id="18ec6-981">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="18ec6-981">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="18ec6-982">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="18ec6-982">TaskSuggestion</span></span>|<span data-ttu-id="18ec6-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="18ec6-983">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="18ec6-984">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-984">String</span></span>|<span data-ttu-id="18ec6-985">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="18ec6-985">**Restricted**</span></span>|

<span data-ttu-id="18ec6-986">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="18ec6-986">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="18ec6-987">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-987">Example</span></span>

<span data-ttu-id="18ec6-988">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="18ec6-988">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="18ec6-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="18ec6-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="18ec6-990">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="18ec6-990">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-991">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-991">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18ec6-992">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="18ec6-992">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-993">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-993">Parameters:</span></span>

|<span data-ttu-id="18ec6-994">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-994">Name</span></span>|<span data-ttu-id="18ec6-995">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-995">Type</span></span>|<span data-ttu-id="18ec6-996">object</span><span class="sxs-lookup"><span data-stu-id="18ec6-996">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="18ec6-997">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-997">String</span></span>|<span data-ttu-id="18ec6-998">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="18ec6-998">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-999">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-999">Requirements</span></span>

|<span data-ttu-id="18ec6-1000">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1000">Requirement</span></span>|<span data-ttu-id="18ec6-1001">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1002">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-1003">1.0</span></span>|
|[<span data-ttu-id="18ec6-1004">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1004">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1005">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-1006">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1006">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1007">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-1007">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18ec6-1008">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1008">Returns:</span></span>

<span data-ttu-id="18ec6-p158">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="18ec6-1011">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="18ec6-1011">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="18ec6-1012">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="18ec6-1012">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="18ec6-1013">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1013">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-1014">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1014">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18ec6-p159">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="18ec6-1018">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1018">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="18ec6-1019">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1019">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="18ec6-p160">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-1023">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1023">Requirements</span></span>

|<span data-ttu-id="18ec6-1024">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1024">Requirement</span></span>|<span data-ttu-id="18ec6-1025">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1026">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-1027">1.0</span></span>|
|[<span data-ttu-id="18ec6-1028">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1029">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-1030">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1031">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-1031">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18ec6-1032">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1032">Returns:</span></span>

<span data-ttu-id="18ec6-p161">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="18ec6-1035">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="18ec6-1035">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="18ec6-1036">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1036">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="18ec6-1037">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-1037">Example</span></span>

<span data-ttu-id="18ec6-1038">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1038">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="18ec6-1039">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="18ec6-1039">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="18ec6-1040">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1040">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-1041">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1041">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18ec6-1042">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1042">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="18ec6-p162">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-1045">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1045">Parameters:</span></span>

|<span data-ttu-id="18ec6-1046">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-1046">Name</span></span>|<span data-ttu-id="18ec6-1047">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-1047">Type</span></span>|<span data-ttu-id="18ec6-1048">object</span><span class="sxs-lookup"><span data-stu-id="18ec6-1048">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="18ec6-1049">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-1049">String</span></span>|<span data-ttu-id="18ec6-1050">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1050">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-1051">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1051">Requirements</span></span>

|<span data-ttu-id="18ec6-1052">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1052">Requirement</span></span>|<span data-ttu-id="18ec6-1053">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1054">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1055">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-1055">1.0</span></span>|
|[<span data-ttu-id="18ec6-1056">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1056">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1057">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1057">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-1058">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1058">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1059">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-1059">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18ec6-1060">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1060">Returns:</span></span>

<span data-ttu-id="18ec6-1061">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1061">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="18ec6-1062">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="18ec6-1062">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="18ec6-1063">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="18ec6-1063">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="18ec6-1064">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-1064">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="18ec6-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="18ec6-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="18ec6-1066">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1066">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="18ec6-p163">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-1069">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1069">Parameters:</span></span>

|<span data-ttu-id="18ec6-1070">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-1070">Name</span></span>|<span data-ttu-id="18ec6-1071">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-1071">Type</span></span>|<span data-ttu-id="18ec6-1072">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-1072">Attributes</span></span>|<span data-ttu-id="18ec6-1073">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-1073">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="18ec6-1074">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="18ec6-1074">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="18ec6-p164">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="18ec6-1078">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1078">Object</span></span>|<span data-ttu-id="18ec6-1079">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1079">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1080">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1080">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18ec6-1081">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1081">Object</span></span>|<span data-ttu-id="18ec6-1082">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1083">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1083">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18ec6-1084">fonction</span><span class="sxs-lookup"><span data-stu-id="18ec6-1084">function</span></span>||<span data-ttu-id="18ec6-1085">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-1085">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="18ec6-1086">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1086">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="18ec6-1087">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1087">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-1088">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1088">Requirements</span></span>

|<span data-ttu-id="18ec6-1089">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1089">Requirement</span></span>|<span data-ttu-id="18ec6-1090">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1091">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1092">1.2</span><span class="sxs-lookup"><span data-stu-id="18ec6-1092">1.2</span></span>|
|[<span data-ttu-id="18ec6-1093">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="18ec6-1095">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1096">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-1096">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="18ec6-1097">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1097">Returns:</span></span>

<span data-ttu-id="18ec6-1098">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1098">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="18ec6-1099">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="18ec6-1099">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="18ec6-1100">Chaîne</span><span class="sxs-lookup"><span data-stu-id="18ec6-1100">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="18ec6-1101">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-1101">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="18ec6-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="18ec6-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="18ec6-p166">Permet d’obtenir les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="18ec6-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-1105">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1105">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-1106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1106">Requirements</span></span>

|<span data-ttu-id="18ec6-1107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1107">Requirement</span></span>|<span data-ttu-id="18ec6-1108">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1108">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1110">1.6</span><span class="sxs-lookup"><span data-stu-id="18ec6-1110">1.6</span></span>|
|[<span data-ttu-id="18ec6-1111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1112">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1112">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-1113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1114">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-1114">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18ec6-1115">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1115">Returns:</span></span>

<span data-ttu-id="18ec6-1116">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="18ec6-1116">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="18ec6-1117">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-1117">Example</span></span>

<span data-ttu-id="18ec6-1118">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1118">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="18ec6-1119">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="18ec6-1119">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="18ec6-p167">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="18ec6-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-1122">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1122">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18ec6-p168">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="18ec6-1126">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1126">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="18ec6-1127">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1127">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="18ec6-p169">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18ec6-1131">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1131">Requirements</span></span>

|<span data-ttu-id="18ec6-1132">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1132">Requirement</span></span>|<span data-ttu-id="18ec6-1133">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1133">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1134">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1135">1.6</span><span class="sxs-lookup"><span data-stu-id="18ec6-1135">1.6</span></span>|
|[<span data-ttu-id="18ec6-1136">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1136">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1137">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-1138">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1138">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1139">Lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-1139">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18ec6-1140">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1140">Returns:</span></span>

<span data-ttu-id="18ec6-p170">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="18ec6-1143">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-1143">Example</span></span>

<span data-ttu-id="18ec6-1144">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1144">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="18ec6-1145">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="18ec6-1145">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="18ec6-1146">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1146">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="18ec6-p171">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-1150">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1150">Parameters:</span></span>

|<span data-ttu-id="18ec6-1151">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-1151">Name</span></span>|<span data-ttu-id="18ec6-1152">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-1152">Type</span></span>|<span data-ttu-id="18ec6-1153">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-1153">Attributes</span></span>|<span data-ttu-id="18ec6-1154">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-1154">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="18ec6-1155">function</span><span class="sxs-lookup"><span data-stu-id="18ec6-1155">function</span></span>||<span data-ttu-id="18ec6-1156">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="18ec6-1157">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1157">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="18ec6-1158">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1158">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="18ec6-1159">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1159">Object</span></span>|<span data-ttu-id="18ec6-1160">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1161">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1161">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="18ec6-1162">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1162">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-1163">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1163">Requirements</span></span>

|<span data-ttu-id="18ec6-1164">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1164">Requirement</span></span>|<span data-ttu-id="18ec6-1165">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1165">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1166">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1167">1.0</span><span class="sxs-lookup"><span data-stu-id="18ec6-1167">1.0</span></span>|
|[<span data-ttu-id="18ec6-1168">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1169">ReadItem</span></span>|
|[<span data-ttu-id="18ec6-1170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1171">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-1171">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-1172">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-1172">Example</span></span>

<span data-ttu-id="18ec6-p174">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="18ec6-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18ec6-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="18ec6-1177">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1177">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="18ec6-p175">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-1182">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1182">Parameters:</span></span>

|<span data-ttu-id="18ec6-1183">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-1183">Name</span></span>|<span data-ttu-id="18ec6-1184">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-1184">Type</span></span>|<span data-ttu-id="18ec6-1185">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-1185">Attributes</span></span>|<span data-ttu-id="18ec6-1186">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-1186">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="18ec6-1187">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-1187">String</span></span>||<span data-ttu-id="18ec6-1188">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1188">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="18ec6-1189">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1189">Object</span></span>|<span data-ttu-id="18ec6-1190">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1191">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18ec6-1192">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1192">Object</span></span>|<span data-ttu-id="18ec6-1193">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1194">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18ec6-1195">fonction</span><span class="sxs-lookup"><span data-stu-id="18ec6-1195">function</span></span>|<span data-ttu-id="18ec6-1196">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1197">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="18ec6-1198">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="18ec6-1199">Erreurs</span><span class="sxs-lookup"><span data-stu-id="18ec6-1199">Errors</span></span>

|<span data-ttu-id="18ec6-1200">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1200">Error code</span></span>|<span data-ttu-id="18ec6-1201">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="18ec6-1202">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-1203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1203">Requirements</span></span>

|<span data-ttu-id="18ec6-1204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1204">Requirement</span></span>|<span data-ttu-id="18ec6-1205">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="18ec6-1207">1.1</span></span>|
|[<span data-ttu-id="18ec6-1208">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="18ec6-1210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1211">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-1212">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-1212">Example</span></span>

<span data-ttu-id="18ec6-1213">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="18ec6-1213">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="18ec6-1214">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18ec6-1214">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="18ec6-1215">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1215">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="18ec6-1216">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-1217">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1217">Parameters:</span></span>

| <span data-ttu-id="18ec6-1218">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-1218">Name</span></span> | <span data-ttu-id="18ec6-1219">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-1219">Type</span></span> | <span data-ttu-id="18ec6-1220">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-1220">Attributes</span></span> | <span data-ttu-id="18ec6-1221">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="18ec6-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="18ec6-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="18ec6-1223">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1223">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="18ec6-1224">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1224">Object</span></span> | <span data-ttu-id="18ec6-1225">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1225">&lt;optional&gt;</span></span> | <span data-ttu-id="18ec6-1226">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="18ec6-1227">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1227">Object</span></span> | <span data-ttu-id="18ec6-1228">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1228">&lt;optional&gt;</span></span> | <span data-ttu-id="18ec6-1229">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="18ec6-1230">fonction</span><span class="sxs-lookup"><span data-stu-id="18ec6-1230">function</span></span>| <span data-ttu-id="18ec6-1231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1232">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-1232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-1233">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1233">Requirements</span></span>

|<span data-ttu-id="18ec6-1234">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1234">Requirement</span></span>| <span data-ttu-id="18ec6-1235">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1235">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1236">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18ec6-1237">1.7</span><span class="sxs-lookup"><span data-stu-id="18ec6-1237">1.7</span></span> |
|[<span data-ttu-id="18ec6-1238">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18ec6-1239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1239">ReadItem</span></span> |
|[<span data-ttu-id="18ec6-1240">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18ec6-1241">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="18ec6-1241">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="18ec6-1242">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-1242">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="18ec6-1243">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="18ec6-1243">saveAsync([options], callback)</span></span>

<span data-ttu-id="18ec6-1244">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1244">Asynchronously saves an item.</span></span>

<span data-ttu-id="18ec6-p176">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p176">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-1248">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1248">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="18ec6-1249">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1249">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="18ec6-p178">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p178">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="18ec6-1253">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1253">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="18ec6-1254">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1254">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="18ec6-1255">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1255">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="18ec6-1256">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1256">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-1257">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1257">Parameters:</span></span>

|<span data-ttu-id="18ec6-1258">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-1258">Name</span></span>|<span data-ttu-id="18ec6-1259">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-1259">Type</span></span>|<span data-ttu-id="18ec6-1260">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-1260">Attributes</span></span>|<span data-ttu-id="18ec6-1261">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="18ec6-1262">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1262">Object</span></span>|<span data-ttu-id="18ec6-1263">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1264">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18ec6-1265">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1265">Object</span></span>|<span data-ttu-id="18ec6-1266">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1267">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18ec6-1268">fonction</span><span class="sxs-lookup"><span data-stu-id="18ec6-1268">function</span></span>||<span data-ttu-id="18ec6-1269">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-1269">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="18ec6-1270">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1270">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-1271">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1271">Requirements</span></span>

|<span data-ttu-id="18ec6-1272">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1272">Requirement</span></span>|<span data-ttu-id="18ec6-1273">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1274">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1275">1.3</span><span class="sxs-lookup"><span data-stu-id="18ec6-1275">1.3</span></span>|
|[<span data-ttu-id="18ec6-1276">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="18ec6-1278">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1279">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-1279">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="18ec6-1280">範例</span><span class="sxs-lookup"><span data-stu-id="18ec6-1280">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="18ec6-p180">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p180">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="18ec6-1283">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="18ec6-1283">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="18ec6-1284">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1284">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="18ec6-p181">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p181">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18ec6-1288">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="18ec6-1288">Parameters:</span></span>

|<span data-ttu-id="18ec6-1289">Nom</span><span class="sxs-lookup"><span data-stu-id="18ec6-1289">Name</span></span>|<span data-ttu-id="18ec6-1290">Type</span><span class="sxs-lookup"><span data-stu-id="18ec6-1290">Type</span></span>|<span data-ttu-id="18ec6-1291">Attributs</span><span class="sxs-lookup"><span data-stu-id="18ec6-1291">Attributes</span></span>|<span data-ttu-id="18ec6-1292">Description</span><span class="sxs-lookup"><span data-stu-id="18ec6-1292">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="18ec6-1293">String</span><span class="sxs-lookup"><span data-stu-id="18ec6-1293">String</span></span>||<span data-ttu-id="18ec6-p182">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p182">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="18ec6-1297">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1297">Object</span></span>|<span data-ttu-id="18ec6-1298">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1299">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1299">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18ec6-1300">Objet</span><span class="sxs-lookup"><span data-stu-id="18ec6-1300">Object</span></span>|<span data-ttu-id="18ec6-1301">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1301">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-1302">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1302">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="18ec6-1303">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="18ec6-1303">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="18ec6-1304">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="18ec6-1304">&lt;optional&gt;</span></span>|<span data-ttu-id="18ec6-p183">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p183">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="18ec6-p184">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="18ec6-p184">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="18ec6-1309">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="18ec6-1309">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="18ec6-1310">fonction</span><span class="sxs-lookup"><span data-stu-id="18ec6-1310">function</span></span>||<span data-ttu-id="18ec6-1311">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18ec6-1311">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18ec6-1312">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="18ec6-1312">Requirements</span></span>

|<span data-ttu-id="18ec6-1313">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="18ec6-1313">Requirement</span></span>|<span data-ttu-id="18ec6-1314">Valeur</span><span class="sxs-lookup"><span data-stu-id="18ec6-1314">Value</span></span>|
|---|---|
|[<span data-ttu-id="18ec6-1315">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="18ec6-1315">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18ec6-1316">1.2</span><span class="sxs-lookup"><span data-stu-id="18ec6-1316">1.2</span></span>|
|[<span data-ttu-id="18ec6-1317">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="18ec6-1317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18ec6-1318">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18ec6-1318">ReadWriteItem</span></span>|
|[<span data-ttu-id="18ec6-1319">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="18ec6-1319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18ec6-1320">Composition</span><span class="sxs-lookup"><span data-stu-id="18ec6-1320">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18ec6-1321">Exemple</span><span class="sxs-lookup"><span data-stu-id="18ec6-1321">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
