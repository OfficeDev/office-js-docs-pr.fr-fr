---
title: Office.Context.Mailbox.Item - ensemble de conditions requises d’aperçu
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: b4b2ec9c735270d9b1bfca3d1c24ef6b0f1ca1cb
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389598"
---
# <a name="item"></a><span data-ttu-id="4510a-102">élément</span><span class="sxs-lookup"><span data-stu-id="4510a-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4510a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4510a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4510a-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="4510a-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-106">Requirements</span></span>

|<span data-ttu-id="4510a-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-107">Requirement</span></span>|<span data-ttu-id="4510a-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-110">1.0</span></span>|
|[<span data-ttu-id="4510a-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4510a-112">Restricted</span></span>|
|[<span data-ttu-id="4510a-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4510a-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4510a-115">Members and methods</span></span>

| <span data-ttu-id="4510a-116">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-116">Member</span></span> | <span data-ttu-id="4510a-117">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4510a-118">attachments</span><span class="sxs-lookup"><span data-stu-id="4510a-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="4510a-119">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-119">Member</span></span> |
| [<span data-ttu-id="4510a-120">bcc</span><span class="sxs-lookup"><span data-stu-id="4510a-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4510a-121">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-121">Member</span></span> |
| [<span data-ttu-id="4510a-122">body</span><span class="sxs-lookup"><span data-stu-id="4510a-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="4510a-123">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-123">Member</span></span> |
| [<span data-ttu-id="4510a-124">cc</span><span class="sxs-lookup"><span data-stu-id="4510a-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4510a-125">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-125">Member</span></span> |
| [<span data-ttu-id="4510a-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="4510a-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4510a-127">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-127">Member</span></span> |
| [<span data-ttu-id="4510a-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4510a-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4510a-129">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-129">Member</span></span> |
| [<span data-ttu-id="4510a-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4510a-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4510a-131">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-131">Member</span></span> |
| [<span data-ttu-id="4510a-132">end</span><span class="sxs-lookup"><span data-stu-id="4510a-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="4510a-133">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-133">Member</span></span> |
| [<span data-ttu-id="4510a-134">from</span><span class="sxs-lookup"><span data-stu-id="4510a-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="4510a-135">Member</span><span class="sxs-lookup"><span data-stu-id="4510a-135">Member</span></span> |
| [<span data-ttu-id="4510a-136">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="4510a-136">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="4510a-137">Member</span><span class="sxs-lookup"><span data-stu-id="4510a-137">Member</span></span> |
| [<span data-ttu-id="4510a-138">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4510a-138">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4510a-139">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-139">Member</span></span> |
| [<span data-ttu-id="4510a-140">itemClass</span><span class="sxs-lookup"><span data-stu-id="4510a-140">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4510a-141">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-141">Member</span></span> |
| [<span data-ttu-id="4510a-142">itemId</span><span class="sxs-lookup"><span data-stu-id="4510a-142">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4510a-143">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-143">Member</span></span> |
| [<span data-ttu-id="4510a-144">itemType</span><span class="sxs-lookup"><span data-stu-id="4510a-144">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="4510a-145">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-145">Member</span></span> |
| [<span data-ttu-id="4510a-146">location</span><span class="sxs-lookup"><span data-stu-id="4510a-146">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="4510a-147">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-147">Member</span></span> |
| [<span data-ttu-id="4510a-148">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4510a-148">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4510a-149">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-149">Member</span></span> |
| [<span data-ttu-id="4510a-150">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4510a-150">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="4510a-151">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-151">Member</span></span> |
| [<span data-ttu-id="4510a-152">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4510a-152">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4510a-153">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-153">Member</span></span> |
| [<span data-ttu-id="4510a-154">organizer</span><span class="sxs-lookup"><span data-stu-id="4510a-154">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="4510a-155">Member</span><span class="sxs-lookup"><span data-stu-id="4510a-155">Member</span></span> |
| [<span data-ttu-id="4510a-156">recurrence</span><span class="sxs-lookup"><span data-stu-id="4510a-156">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="4510a-157">Member</span><span class="sxs-lookup"><span data-stu-id="4510a-157">Member</span></span> |
| [<span data-ttu-id="4510a-158">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4510a-158">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4510a-159">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-159">Member</span></span> |
| [<span data-ttu-id="4510a-160">sender</span><span class="sxs-lookup"><span data-stu-id="4510a-160">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="4510a-161">Member</span><span class="sxs-lookup"><span data-stu-id="4510a-161">Member</span></span> |
| [<span data-ttu-id="4510a-162">seriesId</span><span class="sxs-lookup"><span data-stu-id="4510a-162">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="4510a-163">Member</span><span class="sxs-lookup"><span data-stu-id="4510a-163">Member</span></span> |
| [<span data-ttu-id="4510a-164">start</span><span class="sxs-lookup"><span data-stu-id="4510a-164">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="4510a-165">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-165">Member</span></span> |
| [<span data-ttu-id="4510a-166">subject</span><span class="sxs-lookup"><span data-stu-id="4510a-166">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="4510a-167">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-167">Member</span></span> |
| [<span data-ttu-id="4510a-168">to</span><span class="sxs-lookup"><span data-stu-id="4510a-168">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4510a-169">Membre</span><span class="sxs-lookup"><span data-stu-id="4510a-169">Member</span></span> |
| [<span data-ttu-id="4510a-170">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-170">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4510a-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-171">Method</span></span> |
| [<span data-ttu-id="4510a-172">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="4510a-172">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="4510a-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-173">Method</span></span> |
| [<span data-ttu-id="4510a-174">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-174">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4510a-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-175">Method</span></span> |
| [<span data-ttu-id="4510a-176">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-176">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4510a-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-177">Method</span></span> |
| [<span data-ttu-id="4510a-178">close</span><span class="sxs-lookup"><span data-stu-id="4510a-178">close</span></span>](#close) | <span data-ttu-id="4510a-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-179">Method</span></span> |
| [<span data-ttu-id="4510a-180">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4510a-180">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="4510a-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-181">Method</span></span> |
| [<span data-ttu-id="4510a-182">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4510a-182">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="4510a-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-183">Method</span></span> |
| [<span data-ttu-id="4510a-184">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-184">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="4510a-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-185">Method</span></span> |
| [<span data-ttu-id="4510a-186">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-186">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="4510a-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-187">Method</span></span> |
| [<span data-ttu-id="4510a-188">getEntities</span><span class="sxs-lookup"><span data-stu-id="4510a-188">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="4510a-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-189">Method</span></span> |
| [<span data-ttu-id="4510a-190">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4510a-190">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="4510a-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-191">Method</span></span> |
| [<span data-ttu-id="4510a-192">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4510a-192">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="4510a-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-193">Method</span></span> |
| [<span data-ttu-id="4510a-194">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-194">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="4510a-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-195">Method</span></span> |
| [<span data-ttu-id="4510a-196">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4510a-196">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4510a-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-197">Method</span></span> |
| [<span data-ttu-id="4510a-198">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4510a-198">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4510a-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-199">Method</span></span> |
| [<span data-ttu-id="4510a-200">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-200">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4510a-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-201">Method</span></span> |
| [<span data-ttu-id="4510a-202">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="4510a-202">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="4510a-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-203">Method</span></span> |
| [<span data-ttu-id="4510a-204">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4510a-204">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="4510a-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-205">Method</span></span> |
| [<span data-ttu-id="4510a-206">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-206">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="4510a-207">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-207">Method</span></span> |
| [<span data-ttu-id="4510a-208">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-208">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4510a-209">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-209">Method</span></span> |
| [<span data-ttu-id="4510a-210">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-210">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4510a-211">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-211">Method</span></span> |
| [<span data-ttu-id="4510a-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4510a-213">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-213">Method</span></span> |
| [<span data-ttu-id="4510a-214">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-214">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4510a-215">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-215">Method</span></span> |
| [<span data-ttu-id="4510a-216">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4510a-216">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4510a-217">Méthode</span><span class="sxs-lookup"><span data-stu-id="4510a-217">Method</span></span> |

### <a name="example"></a><span data-ttu-id="4510a-218">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-218">Example</span></span>

<span data-ttu-id="4510a-219">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="4510a-219">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
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

### <a name="members"></a><span data-ttu-id="4510a-220">Membres</span><span class="sxs-lookup"><span data-stu-id="4510a-220">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="4510a-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4510a-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="4510a-222">Permet d’obtenir les pièces jointes de l’élément sous forme de tableau.</span><span class="sxs-lookup"><span data-stu-id="4510a-222">Gets the item's attachments as an array.</span></span> <span data-ttu-id="4510a-223">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-223">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-224">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="4510a-224">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4510a-225">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="4510a-225">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-226">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-226">Type:</span></span>

*   <span data-ttu-id="4510a-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4510a-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-228">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-228">Requirements</span></span>

|<span data-ttu-id="4510a-229">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-229">Requirement</span></span>|<span data-ttu-id="4510a-230">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-231">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-232">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-232">1.0</span></span>|
|[<span data-ttu-id="4510a-233">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-233">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-234">ReadItem</span></span>|
|[<span data-ttu-id="4510a-235">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-235">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-236">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-236">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-237">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-237">Example</span></span>

<span data-ttu-id="4510a-238">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4510a-238">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4510a-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4510a-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4510a-240">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="4510a-240">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4510a-241">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-241">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-242">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-242">Type:</span></span>

*   [<span data-ttu-id="4510a-243">Destinataires</span><span class="sxs-lookup"><span data-stu-id="4510a-243">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="4510a-244">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-244">Requirements</span></span>

|<span data-ttu-id="4510a-245">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-245">Requirement</span></span>|<span data-ttu-id="4510a-246">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-247">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-248">1.1</span><span class="sxs-lookup"><span data-stu-id="4510a-248">1.1</span></span>|
|[<span data-ttu-id="4510a-249">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-250">ReadItem</span></span>|
|[<span data-ttu-id="4510a-251">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-252">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-252">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-253">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-253">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="4510a-254">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="4510a-254">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="4510a-255">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-255">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-256">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-256">Type:</span></span>

*   [<span data-ttu-id="4510a-257">Corps</span><span class="sxs-lookup"><span data-stu-id="4510a-257">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="4510a-258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-258">Requirements</span></span>

|<span data-ttu-id="4510a-259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-259">Requirement</span></span>|<span data-ttu-id="4510a-260">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-262">1.1</span><span class="sxs-lookup"><span data-stu-id="4510a-262">1.1</span></span>|
|[<span data-ttu-id="4510a-263">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-263">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-264">ReadItem</span></span>|
|[<span data-ttu-id="4510a-265">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-265">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-266">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-266">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4510a-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4510a-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4510a-268">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="4510a-268">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4510a-269">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4510a-269">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-270">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-270">Read mode</span></span>

<span data-ttu-id="4510a-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="4510a-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4510a-273">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-273">Compose mode</span></span>

<span data-ttu-id="4510a-274">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="4510a-274">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-275">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-275">Type:</span></span>

*   <span data-ttu-id="4510a-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4510a-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-277">Requirements</span></span>

|<span data-ttu-id="4510a-278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-278">Requirement</span></span>|<span data-ttu-id="4510a-279">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-281">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-281">1.0</span></span>|
|[<span data-ttu-id="4510a-282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-283">ReadItem</span></span>|
|[<span data-ttu-id="4510a-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-285">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-285">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-286">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-286">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="4510a-287">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="4510a-287">(nullable) conversationId :String</span></span>

<span data-ttu-id="4510a-288">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="4510a-288">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4510a-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="4510a-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4510a-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="4510a-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-293">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-293">Type:</span></span>

*   <span data-ttu-id="4510a-294">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-294">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-295">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-295">Requirements</span></span>

|<span data-ttu-id="4510a-296">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-296">Requirement</span></span>|<span data-ttu-id="4510a-297">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-298">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-299">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-299">1.0</span></span>|
|[<span data-ttu-id="4510a-300">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-300">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-301">ReadItem</span></span>|
|[<span data-ttu-id="4510a-302">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-302">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-303">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-303">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="4510a-304">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="4510a-304">dateTimeCreated :Date</span></span>

<span data-ttu-id="4510a-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-307">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-307">Type:</span></span>

*   <span data-ttu-id="4510a-308">Date</span><span class="sxs-lookup"><span data-stu-id="4510a-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-309">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-309">Requirements</span></span>

|<span data-ttu-id="4510a-310">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-310">Requirement</span></span>|<span data-ttu-id="4510a-311">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-312">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-313">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-313">1.0</span></span>|
|[<span data-ttu-id="4510a-314">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-315">ReadItem</span></span>|
|[<span data-ttu-id="4510a-316">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-317">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-318">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-318">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="4510a-319">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="4510a-319">dateTimeModified :Date</span></span>

<span data-ttu-id="4510a-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-322">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-322">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-323">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-323">Type:</span></span>

*   <span data-ttu-id="4510a-324">Date</span><span class="sxs-lookup"><span data-stu-id="4510a-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-325">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-325">Requirements</span></span>

|<span data-ttu-id="4510a-326">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-326">Requirement</span></span>|<span data-ttu-id="4510a-327">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-328">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-329">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-329">1.0</span></span>|
|[<span data-ttu-id="4510a-330">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-331">ReadItem</span></span>|
|[<span data-ttu-id="4510a-332">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-333">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-334">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-334">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="4510a-335">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4510a-335">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="4510a-336">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4510a-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="4510a-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-339">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-339">Read mode</span></span>

<span data-ttu-id="4510a-340">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="4510a-340">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4510a-341">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-341">Compose mode</span></span>

<span data-ttu-id="4510a-342">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4510a-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4510a-343">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="4510a-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-344">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-344">Type:</span></span>

*   <span data-ttu-id="4510a-345">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4510a-345">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-346">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-346">Requirements</span></span>

|<span data-ttu-id="4510a-347">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-347">Requirement</span></span>|<span data-ttu-id="4510a-348">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-348">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-349">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-349">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-350">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-350">1.0</span></span>|
|[<span data-ttu-id="4510a-351">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-351">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-352">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-352">ReadItem</span></span>|
|[<span data-ttu-id="4510a-353">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-353">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-354">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-354">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-355">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-355">Example</span></span>

<span data-ttu-id="4510a-356">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4510a-356">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="4510a-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="4510a-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="4510a-358">Permet d’obtenir l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="4510a-358">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="4510a-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="4510a-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-361">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4510a-361">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-362">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-362">Read mode</span></span>

<span data-ttu-id="4510a-363">La propriété `from` renvoie un objet `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="4510a-363">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="4510a-364">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-364">Compose mode</span></span>

<span data-ttu-id="4510a-365">La propriété `from` renvoie un objet `From` qui fournit une méthode pour obtenir la valeur from.</span><span class="sxs-lookup"><span data-stu-id="4510a-365">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4510a-366">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-366">Type:</span></span>

*   <span data-ttu-id="4510a-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="4510a-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-368">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-368">Requirements</span></span>

|<span data-ttu-id="4510a-369">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-369">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4510a-370">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-371">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-371">1.0</span></span>|<span data-ttu-id="4510a-372">1.7</span><span class="sxs-lookup"><span data-stu-id="4510a-372">1.7</span></span>|
|[<span data-ttu-id="4510a-373">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-373">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-374">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-374">ReadItem</span></span>|<span data-ttu-id="4510a-375">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4510a-375">ReadWriteItem</span></span>|
|[<span data-ttu-id="4510a-376">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-376">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-377">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-377">Read</span></span>|<span data-ttu-id="4510a-378">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-378">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="4510a-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="4510a-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="4510a-380">Permet d’obtenir ou de définir les en-têtes Internet d’un message.</span><span class="sxs-lookup"><span data-stu-id="4510a-380">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-381">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-381">Type:</span></span>

*   [<span data-ttu-id="4510a-382">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="4510a-382">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="4510a-383">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-383">Requirements</span></span>

|<span data-ttu-id="4510a-384">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-384">Requirement</span></span>|<span data-ttu-id="4510a-385">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-386">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-387">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4510a-387">Preview</span></span>|
|[<span data-ttu-id="4510a-388">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-389">ReadItem</span></span>|
|[<span data-ttu-id="4510a-390">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-391">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-391">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="4510a-392">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="4510a-392">internetMessageId :String</span></span>

<span data-ttu-id="4510a-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-395">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-395">Type:</span></span>

*   <span data-ttu-id="4510a-396">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-397">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-397">Requirements</span></span>

|<span data-ttu-id="4510a-398">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-398">Requirement</span></span>|<span data-ttu-id="4510a-399">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-400">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-401">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-401">1.0</span></span>|
|[<span data-ttu-id="4510a-402">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-403">ReadItem</span></span>|
|[<span data-ttu-id="4510a-404">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-405">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-406">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-406">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="4510a-407">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="4510a-407">itemClass :String</span></span>

<span data-ttu-id="4510a-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4510a-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="4510a-412">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-412">Type</span></span>|<span data-ttu-id="4510a-413">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-413">Description</span></span>|<span data-ttu-id="4510a-414">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="4510a-414">item class</span></span>|
|---|---|---|
|<span data-ttu-id="4510a-415">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="4510a-415">Appointment items</span></span>|<span data-ttu-id="4510a-416">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="4510a-416">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="4510a-417">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="4510a-417">Message items</span></span>|<span data-ttu-id="4510a-418">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="4510a-418">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="4510a-419">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="4510a-419">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-420">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-420">Type:</span></span>

*   <span data-ttu-id="4510a-421">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-421">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-422">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-422">Requirements</span></span>

|<span data-ttu-id="4510a-423">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-423">Requirement</span></span>|<span data-ttu-id="4510a-424">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-425">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-426">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-426">1.0</span></span>|
|[<span data-ttu-id="4510a-427">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-428">ReadItem</span></span>|
|[<span data-ttu-id="4510a-429">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-430">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-430">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-431">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-431">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4510a-432">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="4510a-432">(nullable) itemId :String</span></span>

<span data-ttu-id="4510a-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-435">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4510a-435">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4510a-436">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="4510a-436">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4510a-437">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4510a-437">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4510a-438">Pour plus d’informations, consultez la rubrique [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="4510a-438">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4510a-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-441">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-441">Type:</span></span>

*   <span data-ttu-id="4510a-442">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-443">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-443">Requirements</span></span>

|<span data-ttu-id="4510a-444">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-444">Requirement</span></span>|<span data-ttu-id="4510a-445">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-446">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-447">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-447">1.0</span></span>|
|[<span data-ttu-id="4510a-448">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-449">ReadItem</span></span>|
|[<span data-ttu-id="4510a-450">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-451">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-452">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-452">Example</span></span>

<span data-ttu-id="4510a-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4510a-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="4510a-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="4510a-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="4510a-456">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="4510a-456">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4510a-457">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-457">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-458">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-458">Type:</span></span>

*   [<span data-ttu-id="4510a-459">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4510a-459">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="4510a-460">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-460">Requirements</span></span>

|<span data-ttu-id="4510a-461">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-461">Requirement</span></span>|<span data-ttu-id="4510a-462">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-463">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-464">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-464">1.0</span></span>|
|[<span data-ttu-id="4510a-465">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-465">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-466">ReadItem</span></span>|
|[<span data-ttu-id="4510a-467">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-467">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-468">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-468">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-469">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-469">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="4510a-470">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="4510a-470">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="4510a-471">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-471">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-472">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-472">Read mode</span></span>

<span data-ttu-id="4510a-473">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-473">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4510a-474">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-474">Compose mode</span></span>

<span data-ttu-id="4510a-475">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-475">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-476">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-476">Type:</span></span>

*   <span data-ttu-id="4510a-477">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="4510a-477">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-478">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-478">Requirements</span></span>

|<span data-ttu-id="4510a-479">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-479">Requirement</span></span>|<span data-ttu-id="4510a-480">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-481">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-482">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-482">1.0</span></span>|
|[<span data-ttu-id="4510a-483">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-484">ReadItem</span></span>|
|[<span data-ttu-id="4510a-485">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-486">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-487">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-487">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4510a-488">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="4510a-488">normalizedSubject :String</span></span>

<span data-ttu-id="4510a-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4510a-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="4510a-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-493">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-493">Type:</span></span>

*   <span data-ttu-id="4510a-494">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-494">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-495">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-495">Requirements</span></span>

|<span data-ttu-id="4510a-496">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-496">Requirement</span></span>|<span data-ttu-id="4510a-497">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-497">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-498">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-499">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-499">1.0</span></span>|
|[<span data-ttu-id="4510a-500">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-500">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-501">ReadItem</span></span>|
|[<span data-ttu-id="4510a-502">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-502">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-503">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-503">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-504">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-504">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="4510a-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="4510a-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="4510a-506">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-506">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-507">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-507">Type:</span></span>

*   [<span data-ttu-id="4510a-508">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4510a-508">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="4510a-509">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-509">Requirements</span></span>

|<span data-ttu-id="4510a-510">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-510">Requirement</span></span>|<span data-ttu-id="4510a-511">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-511">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-512">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-512">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-513">1.3</span><span class="sxs-lookup"><span data-stu-id="4510a-513">1.3</span></span>|
|[<span data-ttu-id="4510a-514">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-514">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-515">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-515">ReadItem</span></span>|
|[<span data-ttu-id="4510a-516">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-516">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-517">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-517">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4510a-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4510a-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4510a-519">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="4510a-519">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4510a-520">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4510a-520">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-521">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-521">Read mode</span></span>

<span data-ttu-id="4510a-522">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="4510a-522">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4510a-523">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-523">Compose mode</span></span>

<span data-ttu-id="4510a-524">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="4510a-524">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-525">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-525">Type:</span></span>

*   <span data-ttu-id="4510a-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4510a-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-527">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-527">Requirements</span></span>

|<span data-ttu-id="4510a-528">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-528">Requirement</span></span>|<span data-ttu-id="4510a-529">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-529">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-530">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-531">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-531">1.0</span></span>|
|[<span data-ttu-id="4510a-532">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-532">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-533">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-533">ReadItem</span></span>|
|[<span data-ttu-id="4510a-534">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-534">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-535">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-535">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-536">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-536">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="4510a-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="4510a-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="4510a-538">Permet d’obtenir l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="4510a-538">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-539">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-539">Read mode</span></span>

<span data-ttu-id="4510a-540">La propriété `organizer` renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="4510a-540">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4510a-541">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-541">Compose mode</span></span>

<span data-ttu-id="4510a-542">La propriété `organizer` renvoie un objet [Organizer](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur organizer.</span><span class="sxs-lookup"><span data-stu-id="4510a-542">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-543">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-543">Type:</span></span>

*   <span data-ttu-id="4510a-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="4510a-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-545">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-545">Requirements</span></span>

|<span data-ttu-id="4510a-546">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-546">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4510a-547">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-547">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-548">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-548">1.0</span></span>|<span data-ttu-id="4510a-549">1.7</span><span class="sxs-lookup"><span data-stu-id="4510a-549">1.7</span></span>|
|[<span data-ttu-id="4510a-550">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-551">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-551">ReadItem</span></span>|<span data-ttu-id="4510a-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4510a-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="4510a-553">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-554">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-554">Read</span></span>|<span data-ttu-id="4510a-555">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-556">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-556">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="4510a-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="4510a-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="4510a-558">Permet d’obtenir ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-558">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="4510a-559">Permet d’obtenir la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="4510a-559">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="4510a-560">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-560">Read and compose modes for appointment items.</span></span> <span data-ttu-id="4510a-561">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="4510a-561">Read mode for meeting request items.</span></span>

<span data-ttu-id="4510a-562">La propriété `recurrence` renvoie un objet [périodicité](/javascript/api/outlook/office.recurrence) pour des demandes de réunions ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="4510a-562">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="4510a-563">La valeur `null` est renvoyée pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="4510a-563">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="4510a-564">La valeur `undefined` est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="4510a-564">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="4510a-565">Remarque : les demandes de réunion ont une valeur `itemClass` d’IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="4510a-565">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="4510a-566">Remarque : si l’objet de périodicité est `null`, cela indique que l’objet est un rendez-vous unique ou une demande de réunion de rendez-vous unique, et NON une partie d’une série.</span><span class="sxs-lookup"><span data-stu-id="4510a-566">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-567">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-567">Type:</span></span>

* [<span data-ttu-id="4510a-568">Recurrence</span><span class="sxs-lookup"><span data-stu-id="4510a-568">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="4510a-569">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-569">Requirement</span></span>|<span data-ttu-id="4510a-570">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-570">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-571">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-571">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-572">1.7</span><span class="sxs-lookup"><span data-stu-id="4510a-572">1.7</span></span>|
|[<span data-ttu-id="4510a-573">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-573">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-574">ReadItem</span></span>|
|[<span data-ttu-id="4510a-575">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-575">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-576">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-576">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4510a-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4510a-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4510a-578">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="4510a-578">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4510a-579">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4510a-579">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-580">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-580">Read mode</span></span>

<span data-ttu-id="4510a-581">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="4510a-581">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4510a-582">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-582">Compose mode</span></span>

<span data-ttu-id="4510a-583">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="4510a-583">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-584">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-584">Type:</span></span>

*   <span data-ttu-id="4510a-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4510a-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-586">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-586">Requirements</span></span>

|<span data-ttu-id="4510a-587">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-587">Requirement</span></span>|<span data-ttu-id="4510a-588">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-588">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-589">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-589">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-590">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-590">1.0</span></span>|
|[<span data-ttu-id="4510a-591">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-591">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-592">ReadItem</span></span>|
|[<span data-ttu-id="4510a-593">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-593">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-594">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-594">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-595">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-595">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="4510a-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4510a-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="4510a-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4510a-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="4510a-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-601">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4510a-601">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-602">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-602">Type:</span></span>

*   [<span data-ttu-id="4510a-603">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4510a-603">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4510a-604">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-604">Requirements</span></span>

|<span data-ttu-id="4510a-605">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-605">Requirement</span></span>|<span data-ttu-id="4510a-606">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-607">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-608">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-608">1.0</span></span>|
|[<span data-ttu-id="4510a-609">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-610">ReadItem</span></span>|
|[<span data-ttu-id="4510a-611">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-612">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-612">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-613">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-613">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="4510a-614">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="4510a-614">(nullable) seriesId :String</span></span>

<span data-ttu-id="4510a-615">Permet d’obtenir l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="4510a-615">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="4510a-616">Dans OWA et Outlook, `seriesId` renvoie l’identificateur de services web Exchange (EWS) de l’élément (series) parent auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="4510a-616">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="4510a-617">Dans iOS et Android, `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="4510a-617">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-618">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4510a-618">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4510a-619">La propriété `seriesId` n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="4510a-619">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="4510a-620">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4510a-620">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4510a-621">Pour plus d’informations, consultez la rubrique [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="4510a-621">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="4510a-622">La propriété `seriesId` renvoie `null` pour les éléments qui n’ont pas d’élément parent, tels que des rendez-vous uniques, des éléments de séries ou des demandes de réunion, et renvoie `undefined` pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="4510a-622">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-623">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-623">Type:</span></span>

* <span data-ttu-id="4510a-624">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-624">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-625">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-625">Requirements</span></span>

|<span data-ttu-id="4510a-626">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-626">Requirement</span></span>|<span data-ttu-id="4510a-627">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-628">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-629">1.7</span><span class="sxs-lookup"><span data-stu-id="4510a-629">1.7</span></span>|
|[<span data-ttu-id="4510a-630">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-631">ReadItem</span></span>|
|[<span data-ttu-id="4510a-632">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-633">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-633">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-634">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-634">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="4510a-635">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4510a-635">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="4510a-636">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-636">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4510a-p130">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="4510a-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-639">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-639">Read mode</span></span>

<span data-ttu-id="4510a-640">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="4510a-640">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4510a-641">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-641">Compose mode</span></span>

<span data-ttu-id="4510a-642">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4510a-642">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4510a-643">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="4510a-643">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-644">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-644">Type:</span></span>

*   <span data-ttu-id="4510a-645">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4510a-645">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-646">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-646">Requirements</span></span>

|<span data-ttu-id="4510a-647">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-647">Requirement</span></span>|<span data-ttu-id="4510a-648">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-649">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-650">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-650">1.0</span></span>|
|[<span data-ttu-id="4510a-651">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-652">ReadItem</span></span>|
|[<span data-ttu-id="4510a-653">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-654">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-655">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-655">Example</span></span>

<span data-ttu-id="4510a-656">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4510a-656">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="4510a-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4510a-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="4510a-658">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4510a-659">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="4510a-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-660">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-660">Read mode</span></span>

<span data-ttu-id="4510a-p131">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="4510a-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="4510a-663">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-663">Compose mode</span></span>

<span data-ttu-id="4510a-664">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="4510a-664">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4510a-665">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-665">Type:</span></span>

*   <span data-ttu-id="4510a-666">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4510a-666">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-667">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-667">Requirements</span></span>

|<span data-ttu-id="4510a-668">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-668">Requirement</span></span>|<span data-ttu-id="4510a-669">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-670">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-671">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-671">1.0</span></span>|
|[<span data-ttu-id="4510a-672">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-673">ReadItem</span></span>|
|[<span data-ttu-id="4510a-674">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-675">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-675">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4510a-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4510a-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4510a-677">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="4510a-677">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4510a-678">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4510a-678">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4510a-679">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-679">Read mode</span></span>

<span data-ttu-id="4510a-p133">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="4510a-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4510a-682">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4510a-682">Compose mode</span></span>

<span data-ttu-id="4510a-683">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="4510a-683">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="4510a-684">Type :</span><span class="sxs-lookup"><span data-stu-id="4510a-684">Type:</span></span>

*   <span data-ttu-id="4510a-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4510a-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-686">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-686">Requirements</span></span>

|<span data-ttu-id="4510a-687">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-687">Requirement</span></span>|<span data-ttu-id="4510a-688">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-688">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-689">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-689">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-690">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-690">1.0</span></span>|
|[<span data-ttu-id="4510a-691">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-691">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-692">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-692">ReadItem</span></span>|
|[<span data-ttu-id="4510a-693">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-693">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-694">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-694">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-695">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-695">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="4510a-696">Méthodes</span><span class="sxs-lookup"><span data-stu-id="4510a-696">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4510a-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4510a-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4510a-698">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="4510a-698">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4510a-699">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="4510a-699">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4510a-700">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4510a-700">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-701">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-701">Parameters:</span></span>
|<span data-ttu-id="4510a-702">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-702">Name</span></span>|<span data-ttu-id="4510a-703">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-703">Type</span></span>|<span data-ttu-id="4510a-704">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-704">Attributes</span></span>|<span data-ttu-id="4510a-705">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-705">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="4510a-706">String</span><span class="sxs-lookup"><span data-stu-id="4510a-706">String</span></span>||<span data-ttu-id="4510a-p134">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="4510a-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4510a-709">String</span><span class="sxs-lookup"><span data-stu-id="4510a-709">String</span></span>||<span data-ttu-id="4510a-p135">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4510a-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4510a-712">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-712">Object</span></span>|<span data-ttu-id="4510a-713">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-713">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-714">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-714">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-715">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-715">Object</span></span>|<span data-ttu-id="4510a-716">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-716">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-717">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-717">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="4510a-718">Boolean</span><span class="sxs-lookup"><span data-stu-id="4510a-718">Boolean</span></span>|<span data-ttu-id="4510a-719">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-719">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-720">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4510a-720">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="4510a-721">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-721">function</span></span>|<span data-ttu-id="4510a-722">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-722">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-723">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-723">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4510a-724">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4510a-724">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4510a-725">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="4510a-725">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4510a-726">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4510a-726">Errors</span></span>

|<span data-ttu-id="4510a-727">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4510a-727">Error code</span></span>|<span data-ttu-id="4510a-728">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-728">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="4510a-729">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="4510a-729">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="4510a-730">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="4510a-730">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4510a-731">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4510a-731">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-732">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-732">Requirements</span></span>

|<span data-ttu-id="4510a-733">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-733">Requirement</span></span>|<span data-ttu-id="4510a-734">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-734">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-735">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-735">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-736">1.1</span><span class="sxs-lookup"><span data-stu-id="4510a-736">1.1</span></span>|
|[<span data-ttu-id="4510a-737">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-737">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-738">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4510a-738">ReadWriteItem</span></span>|
|[<span data-ttu-id="4510a-739">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-739">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-740">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-740">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4510a-741">Exemples</span><span class="sxs-lookup"><span data-stu-id="4510a-741">Examples</span></span>

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

<span data-ttu-id="4510a-742">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="4510a-742">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="4510a-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4510a-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4510a-744">Ajoute un fichier provenant du codage base64 à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="4510a-744">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4510a-745">La méthode `addFileAttachmentFromBase64Async` charge le fichier depuis le codage base64 et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="4510a-745">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="4510a-746">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="4510a-746">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="4510a-747">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4510a-747">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-748">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-748">Parameters:</span></span>
|<span data-ttu-id="4510a-749">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-749">Name</span></span>|<span data-ttu-id="4510a-750">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-750">Type</span></span>|<span data-ttu-id="4510a-751">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-751">Attributes</span></span>|<span data-ttu-id="4510a-752">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-752">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="4510a-753">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-753">String</span></span>||<span data-ttu-id="4510a-754">Contenu codé en base64 d’une image ou d’un fichier à ajouter à un e-mail ou à un événement.</span><span class="sxs-lookup"><span data-stu-id="4510a-754">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="4510a-755">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-755">String</span></span>||<span data-ttu-id="4510a-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4510a-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4510a-758">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-758">Object</span></span>|<span data-ttu-id="4510a-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-759">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-760">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-761">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-761">Object</span></span>|<span data-ttu-id="4510a-762">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-762">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-763">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="4510a-764">Boolean</span><span class="sxs-lookup"><span data-stu-id="4510a-764">Boolean</span></span>|<span data-ttu-id="4510a-765">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-765">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-766">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4510a-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="4510a-767">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-767">function</span></span>|<span data-ttu-id="4510a-768">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-768">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-769">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4510a-770">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4510a-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4510a-771">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="4510a-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4510a-772">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4510a-772">Errors</span></span>

|<span data-ttu-id="4510a-773">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4510a-773">Error code</span></span>|<span data-ttu-id="4510a-774">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="4510a-775">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="4510a-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="4510a-776">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="4510a-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4510a-777">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4510a-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-778">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-778">Requirements</span></span>

|<span data-ttu-id="4510a-779">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-779">Requirement</span></span>|<span data-ttu-id="4510a-780">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-781">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-782">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4510a-782">Preview</span></span>|
|[<span data-ttu-id="4510a-783">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-783">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4510a-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="4510a-785">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-785">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-786">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4510a-787">Exemples</span><span class="sxs-lookup"><span data-stu-id="4510a-787">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4510a-788">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4510a-788">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4510a-789">Ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4510a-789">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4510a-790">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="4510a-790">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-791">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-791">Parameters:</span></span>

| <span data-ttu-id="4510a-792">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-792">Name</span></span> | <span data-ttu-id="4510a-793">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-793">Type</span></span> | <span data-ttu-id="4510a-794">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-794">Attributes</span></span> | <span data-ttu-id="4510a-795">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-795">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4510a-796">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4510a-796">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4510a-797">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="4510a-797">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4510a-798">Fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-798">Function</span></span> || <span data-ttu-id="4510a-p138">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4510a-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4510a-802">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-802">Object</span></span> | <span data-ttu-id="4510a-803">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-803">&lt;optional&gt;</span></span> | <span data-ttu-id="4510a-804">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-804">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4510a-805">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-805">Object</span></span> | <span data-ttu-id="4510a-806">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-806">&lt;optional&gt;</span></span> | <span data-ttu-id="4510a-807">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-807">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4510a-808">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-808">function</span></span>| <span data-ttu-id="4510a-809">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-809">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-810">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-810">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-811">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-811">Requirements</span></span>

|<span data-ttu-id="4510a-812">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-812">Requirement</span></span>| <span data-ttu-id="4510a-813">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-814">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4510a-815">1.7</span><span class="sxs-lookup"><span data-stu-id="4510a-815">1.7</span></span> |
|[<span data-ttu-id="4510a-816">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4510a-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-817">ReadItem</span></span> |
|[<span data-ttu-id="4510a-818">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4510a-819">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-819">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4510a-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4510a-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4510a-821">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-821">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4510a-p139">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4510a-825">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4510a-825">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4510a-826">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="4510a-826">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-827">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-827">Parameters:</span></span>

|<span data-ttu-id="4510a-828">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-828">Name</span></span>|<span data-ttu-id="4510a-829">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-829">Type</span></span>|<span data-ttu-id="4510a-830">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-830">Attributes</span></span>|<span data-ttu-id="4510a-831">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-831">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="4510a-832">String</span><span class="sxs-lookup"><span data-stu-id="4510a-832">String</span></span>||<span data-ttu-id="4510a-p140">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4510a-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4510a-835">String</span><span class="sxs-lookup"><span data-stu-id="4510a-835">String</span></span>||<span data-ttu-id="4510a-p141">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4510a-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4510a-838">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-838">Object</span></span>|<span data-ttu-id="4510a-839">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-839">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-840">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-840">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-841">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-841">Object</span></span>|<span data-ttu-id="4510a-842">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-842">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-843">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-843">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4510a-844">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-844">function</span></span>|<span data-ttu-id="4510a-845">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-845">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-846">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4510a-847">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4510a-847">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4510a-848">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="4510a-848">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4510a-849">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4510a-849">Errors</span></span>

|<span data-ttu-id="4510a-850">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4510a-850">Error code</span></span>|<span data-ttu-id="4510a-851">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-851">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4510a-852">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4510a-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-853">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-853">Requirements</span></span>

|<span data-ttu-id="4510a-854">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-854">Requirement</span></span>|<span data-ttu-id="4510a-855">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-856">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-857">1.1</span><span class="sxs-lookup"><span data-stu-id="4510a-857">1.1</span></span>|
|[<span data-ttu-id="4510a-858">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4510a-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="4510a-860">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-861">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-861">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-862">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-862">Example</span></span>

<span data-ttu-id="4510a-863">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="4510a-863">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

####  <a name="close"></a><span data-ttu-id="4510a-864">close()</span><span class="sxs-lookup"><span data-stu-id="4510a-864">close()</span></span>

<span data-ttu-id="4510a-865">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="4510a-865">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4510a-p142">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="4510a-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-868">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-868">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4510a-869">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="4510a-869">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-870">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-870">Requirements</span></span>

|<span data-ttu-id="4510a-871">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-871">Requirement</span></span>|<span data-ttu-id="4510a-872">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-873">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-874">1.3</span><span class="sxs-lookup"><span data-stu-id="4510a-874">1.3</span></span>|
|[<span data-ttu-id="4510a-875">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-876">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4510a-876">Restricted</span></span>|
|[<span data-ttu-id="4510a-877">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-878">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-878">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="4510a-879">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="4510a-879">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="4510a-880">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4510a-880">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-881">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-881">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4510a-882">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="4510a-882">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4510a-883">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="4510a-883">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4510a-p143">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="4510a-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-887">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-887">Parameters:</span></span>

|<span data-ttu-id="4510a-888">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-888">Name</span></span>|<span data-ttu-id="4510a-889">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-889">Type</span></span>|<span data-ttu-id="4510a-890">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-890">Attributes</span></span>|<span data-ttu-id="4510a-891">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-891">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4510a-892">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4510a-892">String &#124; Object</span></span>||<span data-ttu-id="4510a-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4510a-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4510a-895">**OU**</span><span class="sxs-lookup"><span data-stu-id="4510a-895">**OR**</span></span><br/><span data-ttu-id="4510a-p145">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="4510a-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4510a-898">String</span><span class="sxs-lookup"><span data-stu-id="4510a-898">String</span></span>|<span data-ttu-id="4510a-899">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-899">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4510a-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4510a-902">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-902">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4510a-903">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-903">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-904">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-904">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4510a-905">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-905">String</span></span>||<span data-ttu-id="4510a-p147">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4510a-908">String</span><span class="sxs-lookup"><span data-stu-id="4510a-908">String</span></span>||<span data-ttu-id="4510a-909">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4510a-909">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4510a-910">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-910">String</span></span>||<span data-ttu-id="4510a-p148">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4510a-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4510a-913">Booléen</span><span class="sxs-lookup"><span data-stu-id="4510a-913">Boolean</span></span>||<span data-ttu-id="4510a-p149">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4510a-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4510a-916">String</span><span class="sxs-lookup"><span data-stu-id="4510a-916">String</span></span>||<span data-ttu-id="4510a-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4510a-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4510a-920">function</span><span class="sxs-lookup"><span data-stu-id="4510a-920">function</span></span>|<span data-ttu-id="4510a-921">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-921">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-922">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-923">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-923">Requirements</span></span>

|<span data-ttu-id="4510a-924">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-924">Requirement</span></span>|<span data-ttu-id="4510a-925">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-925">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-926">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-926">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-927">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-927">1.0</span></span>|
|[<span data-ttu-id="4510a-928">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-928">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-929">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-929">ReadItem</span></span>|
|[<span data-ttu-id="4510a-930">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-930">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-931">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-931">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4510a-932">Exemples</span><span class="sxs-lookup"><span data-stu-id="4510a-932">Examples</span></span>

<span data-ttu-id="4510a-933">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="4510a-933">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4510a-934">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="4510a-934">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4510a-935">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="4510a-935">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4510a-936">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="4510a-936">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="4510a-937">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-937">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="4510a-938">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-938">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="4510a-939">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="4510a-939">displayReplyForm(formData)</span></span>

<span data-ttu-id="4510a-940">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4510a-940">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-941">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-941">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4510a-942">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="4510a-942">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4510a-943">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="4510a-943">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4510a-p151">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="4510a-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-947">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-947">Parameters:</span></span>

|<span data-ttu-id="4510a-948">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-948">Name</span></span>|<span data-ttu-id="4510a-949">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-949">Type</span></span>|<span data-ttu-id="4510a-950">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-950">Attributes</span></span>|<span data-ttu-id="4510a-951">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-951">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4510a-952">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4510a-952">String &#124; Object</span></span>||<span data-ttu-id="4510a-p152">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4510a-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4510a-955">**OU**</span><span class="sxs-lookup"><span data-stu-id="4510a-955">**OR**</span></span><br/><span data-ttu-id="4510a-p153">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="4510a-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4510a-958">String</span><span class="sxs-lookup"><span data-stu-id="4510a-958">String</span></span>|<span data-ttu-id="4510a-959">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-959">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4510a-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4510a-962">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-962">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4510a-963">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-963">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-964">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-964">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4510a-965">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-965">String</span></span>||<span data-ttu-id="4510a-p155">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4510a-968">String</span><span class="sxs-lookup"><span data-stu-id="4510a-968">String</span></span>||<span data-ttu-id="4510a-969">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4510a-969">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4510a-970">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-970">String</span></span>||<span data-ttu-id="4510a-p156">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4510a-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4510a-973">Booléen</span><span class="sxs-lookup"><span data-stu-id="4510a-973">Boolean</span></span>||<span data-ttu-id="4510a-p157">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4510a-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4510a-976">String</span><span class="sxs-lookup"><span data-stu-id="4510a-976">String</span></span>||<span data-ttu-id="4510a-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4510a-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4510a-980">function</span><span class="sxs-lookup"><span data-stu-id="4510a-980">function</span></span>|<span data-ttu-id="4510a-981">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-981">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-982">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-983">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-983">Requirements</span></span>

|<span data-ttu-id="4510a-984">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-984">Requirement</span></span>|<span data-ttu-id="4510a-985">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-986">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-987">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-987">1.0</span></span>|
|[<span data-ttu-id="4510a-988">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-989">ReadItem</span></span>|
|[<span data-ttu-id="4510a-990">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-991">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-991">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4510a-992">Exemples</span><span class="sxs-lookup"><span data-stu-id="4510a-992">Examples</span></span>

<span data-ttu-id="4510a-993">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="4510a-993">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4510a-994">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="4510a-994">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4510a-995">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="4510a-995">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4510a-996">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="4510a-996">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="4510a-997">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-997">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="4510a-998">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-998">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="4510a-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="4510a-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="4510a-1000">Permet d’obtenir la pièce jointe spécifiée depuis un message ou un rendez-vous, et la renvoie en tant qu’objet `AttachmentContent`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1000">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="4510a-1001">La méthode `getAttachmentContentAsync` permet d’obtenir la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-1001">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="4510a-1002">Nous vous recommandons de suivre la bonne pratique consistant à utiliser l’identificateur pour récupérer une pièce jointe dans la même session que celle où les objets attachmentIds ont été récupérés avec l’appel `getAttachmentsAsync` ou `item.attachments`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1002">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="4510a-1003">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4510a-1003">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="4510a-1004">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer un formulaire incorporé qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="4510a-1004">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1005">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1005">Parameters:</span></span>

|<span data-ttu-id="4510a-1006">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1006">Name</span></span>|<span data-ttu-id="4510a-1007">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1007">Type</span></span>|<span data-ttu-id="4510a-1008">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1008">Attributes</span></span>|<span data-ttu-id="4510a-1009">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1009">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="4510a-1010">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-1010">String</span></span>||<span data-ttu-id="4510a-1011">Identificateur de la pièce jointe que vous voulez obtenir.</span><span class="sxs-lookup"><span data-stu-id="4510a-1011">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="4510a-1012">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1012">Object</span></span>|<span data-ttu-id="4510a-1013">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1014">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-1015">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1015">Object</span></span>|<span data-ttu-id="4510a-1016">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1017">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4510a-1018">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-1018">function</span></span>|<span data-ttu-id="4510a-1019">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1020">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1021">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1021">Requirements</span></span>

|<span data-ttu-id="4510a-1022">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1022">Requirement</span></span>|<span data-ttu-id="4510a-1023">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1024">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1025">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4510a-1025">Preview</span></span>|
|[<span data-ttu-id="4510a-1026">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1027">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1028">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1029">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1030">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1030">Returns:</span></span>

<span data-ttu-id="4510a-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="4510a-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="4510a-1032">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1032">Example</span></span>

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var options = {asyncContext: {type: result.value[i].attachmentType}};
            getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);  
        }
    }
}

function handleAttachmentsCallback(result) {
    // parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file
    if (result.format == Office.MailboxEnums.AttachmentContentFormat.Base64) {
        // handle file attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.Eml) {
        // handle item attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
        // handle .icalender attachment
    }
    else {
        // handle cloud attachment  
    }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="4510a-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4510a-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="4510a-1034">Permet d’obtenir les pièces jointes de l’élément sous forme de tableau.</span><span class="sxs-lookup"><span data-stu-id="4510a-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="4510a-1035">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="4510a-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1036">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1036">Parameters:</span></span>

|<span data-ttu-id="4510a-1037">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1037">Name</span></span>|<span data-ttu-id="4510a-1038">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1038">Type</span></span>|<span data-ttu-id="4510a-1039">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1039">Attributes</span></span>|<span data-ttu-id="4510a-1040">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4510a-1041">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1041">Object</span></span>|<span data-ttu-id="4510a-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1043">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-1044">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1044">Object</span></span>|<span data-ttu-id="4510a-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1046">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4510a-1047">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-1047">function</span></span>|<span data-ttu-id="4510a-1048">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1049">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1050">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1050">Requirements</span></span>

|<span data-ttu-id="4510a-1051">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1051">Requirement</span></span>|<span data-ttu-id="4510a-1052">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1053">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1054">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4510a-1054">Preview</span></span>|
|[<span data-ttu-id="4510a-1055">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1056">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1057">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1058">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1059">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1059">Returns:</span></span>

<span data-ttu-id="4510a-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4510a-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="4510a-1061">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1061">Example</span></span>

<span data-ttu-id="4510a-1062">L’exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4510a-1062">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);  
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var _att = result.value [i];
            outputString += "<BR>" + i + ". Name: ";
            outputString += _att.name;
            outputString += "<BR>ID: " + _att.id;
            outputString += "<BR>contentType: " + _att.contentType;
            outputString += "<BR>size: " + _att.size;
            outputString += "<BR>attachmentType: " + _att.attachmentType;
            outputString += "<BR>isInline: " + _att.isInline;
        }
    }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="4510a-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4510a-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="4510a-1064">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4510a-1064">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1065">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-1065">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-1066">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1066">Requirements</span></span>

|<span data-ttu-id="4510a-1067">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1067">Requirement</span></span>|<span data-ttu-id="4510a-1068">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1069">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-1070">1.0</span></span>|
|[<span data-ttu-id="4510a-1071">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1072">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1073">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1074">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1075">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1075">Returns:</span></span>

<span data-ttu-id="4510a-1076">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4510a-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4510a-1077">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1077">Example</span></span>

<span data-ttu-id="4510a-1078">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4510a-1078">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="4510a-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4510a-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4510a-1080">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4510a-1080">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1081">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-1081">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1082">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1082">Parameters:</span></span>

|<span data-ttu-id="4510a-1083">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1083">Name</span></span>|<span data-ttu-id="4510a-1084">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1084">Type</span></span>|<span data-ttu-id="4510a-1085">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="4510a-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4510a-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="4510a-1087">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="4510a-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1088">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1088">Requirements</span></span>

|<span data-ttu-id="4510a-1089">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1089">Requirement</span></span>|<span data-ttu-id="4510a-1090">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1091">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-1092">1.0</span></span>|
|[<span data-ttu-id="4510a-1093">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1094">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4510a-1094">Restricted</span></span>|
|[<span data-ttu-id="4510a-1095">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1096">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1097">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1097">Returns:</span></span>

<span data-ttu-id="4510a-1098">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="4510a-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4510a-1099">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="4510a-1099">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4510a-1100">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4510a-1101">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="4510a-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="4510a-1102">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="4510a-1102">Value of `entityType`</span></span>|<span data-ttu-id="4510a-1103">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="4510a-1103">Type of objects in returned array</span></span>|<span data-ttu-id="4510a-1104">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="4510a-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="4510a-1105">String</span><span class="sxs-lookup"><span data-stu-id="4510a-1105">String</span></span>|<span data-ttu-id="4510a-1106">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4510a-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="4510a-1107">Contact</span><span class="sxs-lookup"><span data-stu-id="4510a-1107">Contact</span></span>|<span data-ttu-id="4510a-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4510a-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="4510a-1109">String</span><span class="sxs-lookup"><span data-stu-id="4510a-1109">String</span></span>|<span data-ttu-id="4510a-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4510a-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="4510a-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4510a-1111">MeetingSuggestion</span></span>|<span data-ttu-id="4510a-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4510a-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="4510a-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4510a-1113">PhoneNumber</span></span>|<span data-ttu-id="4510a-1114">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4510a-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="4510a-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4510a-1115">TaskSuggestion</span></span>|<span data-ttu-id="4510a-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4510a-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="4510a-1117">String</span><span class="sxs-lookup"><span data-stu-id="4510a-1117">String</span></span>|<span data-ttu-id="4510a-1118">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4510a-1118">**Restricted**</span></span>|

<span data-ttu-id="4510a-1119">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4510a-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="4510a-1120">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1120">Example</span></span>

<span data-ttu-id="4510a-1121">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4510a-1121">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```javascript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="4510a-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4510a-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4510a-1123">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4510a-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1124">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-1124">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4510a-1125">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="4510a-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1126">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1126">Parameters:</span></span>

|<span data-ttu-id="4510a-1127">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1127">Name</span></span>|<span data-ttu-id="4510a-1128">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1128">Type</span></span>|<span data-ttu-id="4510a-1129">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4510a-1130">String</span><span class="sxs-lookup"><span data-stu-id="4510a-1130">String</span></span>|<span data-ttu-id="4510a-1131">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="4510a-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1132">Requirements</span></span>

|<span data-ttu-id="4510a-1133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1133">Requirement</span></span>|<span data-ttu-id="4510a-1134">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-1136">1.0</span></span>|
|[<span data-ttu-id="4510a-1137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1138">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1140">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1141">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1141">Returns:</span></span>

<span data-ttu-id="4510a-p162">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="4510a-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4510a-1144">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4510a-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="4510a-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4510a-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="4510a-1146">Récupère les données d’initialisation transmises quand le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="4510a-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1147">Cette méthode est uniquement prise en charge par Outlook 2016 ou version ultérieure pour Windows (versions en un clic supérieures à 16.0.8413.1000) et Outlook sur le web pour Office 365.</span><span class="sxs-lookup"><span data-stu-id="4510a-1147">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1148">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1148">Parameters:</span></span>
|<span data-ttu-id="4510a-1149">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1149">Name</span></span>|<span data-ttu-id="4510a-1150">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1150">Type</span></span>|<span data-ttu-id="4510a-1151">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1151">Attributes</span></span>|<span data-ttu-id="4510a-1152">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4510a-1153">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1153">Object</span></span>|<span data-ttu-id="4510a-1154">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1155">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-1156">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1156">Object</span></span>|<span data-ttu-id="4510a-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1158">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4510a-1159">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-1159">function</span></span>|<span data-ttu-id="4510a-1160">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1161">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4510a-1162">En cas de réussite, les données d’initialisation sont fournies dans la propriété `asyncResult.value` sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="4510a-1162">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="4510a-1163">S’il n’existe aucun contexte d’initialisation, l’objet `asyncResult` contient un objet `Error` dont la propriété `code` est définie sur `9020` et la propriété `name` sur `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1164">Requirements</span></span>

|<span data-ttu-id="4510a-1165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1165">Requirement</span></span>|<span data-ttu-id="4510a-1166">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1167">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1168">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4510a-1168">Preview</span></span>|
|[<span data-ttu-id="4510a-1169">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1170">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1172">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-1173">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1173">Example</span></span>

```javascript
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="4510a-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4510a-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4510a-1175">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4510a-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1176">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-1176">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4510a-p163">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="4510a-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4510a-1180">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="4510a-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4510a-1181">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4510a-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-1185">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1185">Requirements</span></span>

|<span data-ttu-id="4510a-1186">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1186">Requirement</span></span>|<span data-ttu-id="4510a-1187">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1188">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-1189">1.0</span></span>|
|[<span data-ttu-id="4510a-1190">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1191">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1192">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1193">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1194">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1194">Returns:</span></span>

<span data-ttu-id="4510a-p165">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="4510a-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="4510a-1197">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4510a-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4510a-1198">Object</span><span class="sxs-lookup"><span data-stu-id="4510a-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4510a-1199">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1199">Example</span></span>

<span data-ttu-id="4510a-1200">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="4510a-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4510a-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="4510a-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4510a-1202">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4510a-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1203">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4510a-1204">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="4510a-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4510a-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="4510a-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1207">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1207">Parameters:</span></span>

|<span data-ttu-id="4510a-1208">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1208">Name</span></span>|<span data-ttu-id="4510a-1209">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1209">Type</span></span>|<span data-ttu-id="4510a-1210">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4510a-1211">String</span><span class="sxs-lookup"><span data-stu-id="4510a-1211">String</span></span>|<span data-ttu-id="4510a-1212">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="4510a-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1213">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1213">Requirements</span></span>

|<span data-ttu-id="4510a-1214">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1214">Requirement</span></span>|<span data-ttu-id="4510a-1215">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1216">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-1217">1.0</span></span>|
|[<span data-ttu-id="4510a-1218">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1219">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1220">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1221">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1222">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1222">Returns:</span></span>

<span data-ttu-id="4510a-1223">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4510a-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="4510a-1224">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4510a-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4510a-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4510a-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4510a-1226">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4510a-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4510a-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4510a-1228">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4510a-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4510a-p167">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="4510a-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1231">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1231">Parameters:</span></span>

|<span data-ttu-id="4510a-1232">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1232">Name</span></span>|<span data-ttu-id="4510a-1233">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1233">Type</span></span>|<span data-ttu-id="4510a-1234">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1234">Attributes</span></span>|<span data-ttu-id="4510a-1235">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="4510a-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4510a-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4510a-p168">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="4510a-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="4510a-1240">Object</span><span class="sxs-lookup"><span data-stu-id="4510a-1240">Object</span></span>|<span data-ttu-id="4510a-1241">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1242">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-1243">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1243">Object</span></span>|<span data-ttu-id="4510a-1244">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1245">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4510a-1246">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-1246">function</span></span>||<span data-ttu-id="4510a-1247">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4510a-1248">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4510a-1249">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1249">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1250">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1250">Requirements</span></span>

|<span data-ttu-id="4510a-1251">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1251">Requirement</span></span>|<span data-ttu-id="4510a-1252">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1253">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="4510a-1254">1.2</span></span>|
|[<span data-ttu-id="4510a-1255">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="4510a-1257">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1258">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1259">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1259">Returns:</span></span>

<span data-ttu-id="4510a-1260">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="4510a-1261">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4510a-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4510a-1262">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4510a-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4510a-1263">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1263">Example</span></span>

```javascript
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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="4510a-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4510a-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="4510a-p170">Permet d’obtenir les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4510a-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1267">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-1267">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-1268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1268">Requirements</span></span>

|<span data-ttu-id="4510a-1269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1269">Requirement</span></span>|<span data-ttu-id="4510a-1270">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="4510a-1272">1.6</span></span>|
|[<span data-ttu-id="4510a-1273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1274">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1276">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1277">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1277">Returns:</span></span>

<span data-ttu-id="4510a-1278">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4510a-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4510a-1279">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1279">Example</span></span>

<span data-ttu-id="4510a-1280">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4510a-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="4510a-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4510a-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="4510a-p171">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4510a-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1284">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4510a-1284">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4510a-p172">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="4510a-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4510a-1288">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="4510a-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4510a-1289">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4510a-p173">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4510a-1293">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1293">Requirements</span></span>

|<span data-ttu-id="4510a-1294">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1294">Requirement</span></span>|<span data-ttu-id="4510a-1295">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1296">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="4510a-1297">1.6</span></span>|
|[<span data-ttu-id="4510a-1298">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1299">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1300">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1301">Lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4510a-1302">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4510a-1302">Returns:</span></span>

<span data-ttu-id="4510a-p174">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="4510a-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="4510a-1305">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1305">Example</span></span>

<span data-ttu-id="4510a-1306">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="4510a-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="4510a-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4510a-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="4510a-1308">Permet d’obtenir les propriétés du rendez-vous ou du message sélectionné dans une boîte aux lettres, un calendrier ou un dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="4510a-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1309">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1309">Parameters:</span></span>

|<span data-ttu-id="4510a-1310">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1310">Name</span></span>|<span data-ttu-id="4510a-1311">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1311">Type</span></span>|<span data-ttu-id="4510a-1312">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1312">Attributes</span></span>|<span data-ttu-id="4510a-1313">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4510a-1314">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1314">Object</span></span>|<span data-ttu-id="4510a-1315">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1316">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-1317">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1317">Object</span></span>|<span data-ttu-id="4510a-1318">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1319">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4510a-1320">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-1320">function</span></span>||<span data-ttu-id="4510a-1321">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4510a-1322">Les propriétés partagées sont fournies sous la forme d’un objet [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1322">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4510a-1323">Cet objet peut être utilisé pour obtenir des propriétés partagées de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1324">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1324">Requirements</span></span>

|<span data-ttu-id="4510a-1325">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1325">Requirement</span></span>|<span data-ttu-id="4510a-1326">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1327">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1328">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4510a-1328">Preview</span></span>|
|[<span data-ttu-id="4510a-1329">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1330">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1331">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1332">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-1333">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4510a-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4510a-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4510a-1335">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4510a-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4510a-p176">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="4510a-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1339">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1339">Parameters:</span></span>

|<span data-ttu-id="4510a-1340">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1340">Name</span></span>|<span data-ttu-id="4510a-1341">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1341">Type</span></span>|<span data-ttu-id="4510a-1342">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1342">Attributes</span></span>|<span data-ttu-id="4510a-1343">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="4510a-1344">function</span><span class="sxs-lookup"><span data-stu-id="4510a-1344">function</span></span>||<span data-ttu-id="4510a-1345">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4510a-1346">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4510a-1347">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="4510a-1347">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="4510a-1348">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1348">Object</span></span>|<span data-ttu-id="4510a-1349">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1350">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1350">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4510a-1351">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1352">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1352">Requirements</span></span>

|<span data-ttu-id="4510a-1353">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1353">Requirement</span></span>|<span data-ttu-id="4510a-1354">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1355">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="4510a-1356">1.0</span></span>|
|[<span data-ttu-id="4510a-1357">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1358">ReadItem</span></span>|
|[<span data-ttu-id="4510a-1359">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1360">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-1361">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1361">Example</span></span>

<span data-ttu-id="4510a-p179">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="4510a-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4510a-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4510a-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4510a-1366">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4510a-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4510a-1367">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="4510a-1368">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="4510a-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="4510a-1369">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4510a-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="4510a-1370">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer un formulaire incorporé qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="4510a-1370">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1371">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1371">Parameters:</span></span>

|<span data-ttu-id="4510a-1372">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1372">Name</span></span>|<span data-ttu-id="4510a-1373">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1373">Type</span></span>|<span data-ttu-id="4510a-1374">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1374">Attributes</span></span>|<span data-ttu-id="4510a-1375">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="4510a-1376">String</span><span class="sxs-lookup"><span data-stu-id="4510a-1376">String</span></span>||<span data-ttu-id="4510a-1377">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="4510a-1377">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="4510a-1378">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1378">Object</span></span>|<span data-ttu-id="4510a-1379">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1379">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1380">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-1380">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-1381">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1381">Object</span></span>|<span data-ttu-id="4510a-1382">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1382">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1383">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1383">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4510a-1384">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-1384">function</span></span>|<span data-ttu-id="4510a-1385">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1386">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1386">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4510a-1387">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="4510a-1387">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4510a-1388">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4510a-1388">Errors</span></span>

|<span data-ttu-id="4510a-1389">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4510a-1389">Error code</span></span>|<span data-ttu-id="4510a-1390">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1390">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="4510a-1391">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="4510a-1391">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1392">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1392">Requirements</span></span>

|<span data-ttu-id="4510a-1393">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1393">Requirement</span></span>|<span data-ttu-id="4510a-1394">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1394">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1395">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1396">1.1</span><span class="sxs-lookup"><span data-stu-id="4510a-1396">1.1</span></span>|
|[<span data-ttu-id="4510a-1397">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1398">ReadWriteItem</span></span>|
|[<span data-ttu-id="4510a-1399">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1400">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-1400">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-1401">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1401">Example</span></span>

<span data-ttu-id="4510a-1402">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="4510a-1402">The following code removes an attachment with an identifier of '0'.</span></span>

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4510a-1403">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4510a-1403">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4510a-1404">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4510a-1404">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4510a-1405">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1405">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1406">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1406">Parameters:</span></span>

| <span data-ttu-id="4510a-1407">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1407">Name</span></span> | <span data-ttu-id="4510a-1408">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1408">Type</span></span> | <span data-ttu-id="4510a-1409">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1409">Attributes</span></span> | <span data-ttu-id="4510a-1410">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1410">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4510a-1411">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4510a-1411">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4510a-1412">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="4510a-1412">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="4510a-1413">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1413">Object</span></span> | <span data-ttu-id="4510a-1414">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1414">&lt;optional&gt;</span></span> | <span data-ttu-id="4510a-1415">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-1415">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4510a-1416">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1416">Object</span></span> | <span data-ttu-id="4510a-1417">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1417">&lt;optional&gt;</span></span> | <span data-ttu-id="4510a-1418">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1418">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4510a-1419">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-1419">function</span></span>| <span data-ttu-id="4510a-1420">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1420">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1421">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1421">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1422">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1422">Requirements</span></span>

|<span data-ttu-id="4510a-1423">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1423">Requirement</span></span>| <span data-ttu-id="4510a-1424">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1424">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1425">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4510a-1426">1.7</span><span class="sxs-lookup"><span data-stu-id="4510a-1426">1.7</span></span> |
|[<span data-ttu-id="4510a-1427">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4510a-1428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1428">ReadItem</span></span> |
|[<span data-ttu-id="4510a-1429">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4510a-1430">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4510a-1430">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="4510a-1431">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4510a-1431">saveAsync([options], callback)</span></span>

<span data-ttu-id="4510a-1432">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4510a-1432">Asynchronously saves an item.</span></span>

<span data-ttu-id="4510a-p181">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="4510a-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1436">si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="4510a-1436">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4510a-1437">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="4510a-1437">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4510a-p183">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="4510a-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4510a-1441">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="4510a-1441">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4510a-1442">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="4510a-1442">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="4510a-1443">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="4510a-1443">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="4510a-1444">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="4510a-1444">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1445">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1445">Parameters:</span></span>

|<span data-ttu-id="4510a-1446">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1446">Name</span></span>|<span data-ttu-id="4510a-1447">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1447">Type</span></span>|<span data-ttu-id="4510a-1448">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1448">Attributes</span></span>|<span data-ttu-id="4510a-1449">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1449">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4510a-1450">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1450">Object</span></span>|<span data-ttu-id="4510a-1451">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1451">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1452">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-1452">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-1453">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1453">Object</span></span>|<span data-ttu-id="4510a-1454">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1454">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1455">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1455">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4510a-1456">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-1456">function</span></span>||<span data-ttu-id="4510a-1457">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1457">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4510a-1458">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4510a-1458">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1459">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1459">Requirements</span></span>

|<span data-ttu-id="4510a-1460">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1460">Requirement</span></span>|<span data-ttu-id="4510a-1461">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1461">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1462">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1463">1.3</span><span class="sxs-lookup"><span data-stu-id="4510a-1463">1.3</span></span>|
|[<span data-ttu-id="4510a-1464">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1465">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1465">ReadWriteItem</span></span>|
|[<span data-ttu-id="4510a-1466">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1467">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-1467">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4510a-1468">範例</span><span class="sxs-lookup"><span data-stu-id="4510a-1468">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="4510a-p185">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4510a-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4510a-1471">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4510a-1471">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4510a-1472">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4510a-1472">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4510a-p186">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="4510a-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4510a-1476">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4510a-1476">Parameters:</span></span>

|<span data-ttu-id="4510a-1477">Nom</span><span class="sxs-lookup"><span data-stu-id="4510a-1477">Name</span></span>|<span data-ttu-id="4510a-1478">Type</span><span class="sxs-lookup"><span data-stu-id="4510a-1478">Type</span></span>|<span data-ttu-id="4510a-1479">Attributs</span><span class="sxs-lookup"><span data-stu-id="4510a-1479">Attributes</span></span>|<span data-ttu-id="4510a-1480">Description</span><span class="sxs-lookup"><span data-stu-id="4510a-1480">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="4510a-1481">String</span><span class="sxs-lookup"><span data-stu-id="4510a-1481">String</span></span>||<span data-ttu-id="4510a-p187">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="4510a-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="4510a-1485">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1485">Object</span></span>|<span data-ttu-id="4510a-1486">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1486">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1487">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4510a-1487">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4510a-1488">Objet</span><span class="sxs-lookup"><span data-stu-id="4510a-1488">Object</span></span>|<span data-ttu-id="4510a-1489">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1489">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-1490">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4510a-1490">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="4510a-1491">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4510a-1491">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="4510a-1492">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4510a-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="4510a-p188">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4510a-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4510a-p189">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="4510a-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4510a-1497">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="4510a-1497">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="4510a-1498">fonction</span><span class="sxs-lookup"><span data-stu-id="4510a-1498">function</span></span>||<span data-ttu-id="4510a-1499">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4510a-1499">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4510a-1500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4510a-1500">Requirements</span></span>

|<span data-ttu-id="4510a-1501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4510a-1501">Requirement</span></span>|<span data-ttu-id="4510a-1502">Valeur</span><span class="sxs-lookup"><span data-stu-id="4510a-1502">Value</span></span>|
|---|---|
|[<span data-ttu-id="4510a-1503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4510a-1503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4510a-1504">1.2</span><span class="sxs-lookup"><span data-stu-id="4510a-1504">1.2</span></span>|
|[<span data-ttu-id="4510a-1505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4510a-1505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4510a-1506">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4510a-1506">ReadWriteItem</span></span>|
|[<span data-ttu-id="4510a-1507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4510a-1507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4510a-1508">Composition</span><span class="sxs-lookup"><span data-stu-id="4510a-1508">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4510a-1509">Exemple</span><span class="sxs-lookup"><span data-stu-id="4510a-1509">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
