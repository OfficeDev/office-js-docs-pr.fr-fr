---
title: Office.Context.Mailbox.Item - ensemble de conditions requises d’aperçu
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: 73495cfaceceec5da9c737f31f6ee96a7452dc3c
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701917"
---
# <a name="item"></a><span data-ttu-id="92251-102">élément</span><span class="sxs-lookup"><span data-stu-id="92251-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="92251-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="92251-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="92251-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="92251-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-106">Requirements</span></span>

|<span data-ttu-id="92251-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-107">Requirement</span></span>|<span data-ttu-id="92251-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-110">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-110">1.0</span></span>|
|[<span data-ttu-id="92251-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="92251-112">Restricted</span></span>|
|[<span data-ttu-id="92251-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="92251-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="92251-115">Members and methods</span></span>

| <span data-ttu-id="92251-116">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-116">Member</span></span> | <span data-ttu-id="92251-117">Type</span><span class="sxs-lookup"><span data-stu-id="92251-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="92251-118">attachments</span><span class="sxs-lookup"><span data-stu-id="92251-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="92251-119">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-119">Member</span></span> |
| [<span data-ttu-id="92251-120">bcc</span><span class="sxs-lookup"><span data-stu-id="92251-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="92251-121">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-121">Member</span></span> |
| [<span data-ttu-id="92251-122">body</span><span class="sxs-lookup"><span data-stu-id="92251-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="92251-123">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-123">Member</span></span> |
| [<span data-ttu-id="92251-124">cc</span><span class="sxs-lookup"><span data-stu-id="92251-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="92251-125">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-125">Member</span></span> |
| [<span data-ttu-id="92251-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="92251-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="92251-127">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-127">Member</span></span> |
| [<span data-ttu-id="92251-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="92251-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="92251-129">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-129">Member</span></span> |
| [<span data-ttu-id="92251-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="92251-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="92251-131">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-131">Member</span></span> |
| [<span data-ttu-id="92251-132">end</span><span class="sxs-lookup"><span data-stu-id="92251-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="92251-133">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-133">Member</span></span> |
| [<span data-ttu-id="92251-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="92251-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) | <span data-ttu-id="92251-135">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-135">Member</span></span> |
| [<span data-ttu-id="92251-136">from</span><span class="sxs-lookup"><span data-stu-id="92251-136">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="92251-137">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-137">Member</span></span> |
| [<span data-ttu-id="92251-138">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="92251-138">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="92251-139">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-139">Member</span></span> |
| [<span data-ttu-id="92251-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="92251-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="92251-141">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-141">Member</span></span> |
| [<span data-ttu-id="92251-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="92251-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="92251-143">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-143">Member</span></span> |
| [<span data-ttu-id="92251-144">itemId</span><span class="sxs-lookup"><span data-stu-id="92251-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="92251-145">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-145">Member</span></span> |
| [<span data-ttu-id="92251-146">itemType</span><span class="sxs-lookup"><span data-stu-id="92251-146">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="92251-147">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-147">Member</span></span> |
| [<span data-ttu-id="92251-148">location</span><span class="sxs-lookup"><span data-stu-id="92251-148">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="92251-149">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-149">Member</span></span> |
| [<span data-ttu-id="92251-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="92251-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="92251-151">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-151">Member</span></span> |
| [<span data-ttu-id="92251-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="92251-152">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="92251-153">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-153">Member</span></span> |
| [<span data-ttu-id="92251-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="92251-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="92251-155">Member</span><span class="sxs-lookup"><span data-stu-id="92251-155">Member</span></span> |
| [<span data-ttu-id="92251-156">organizer</span><span class="sxs-lookup"><span data-stu-id="92251-156">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="92251-157">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-157">Member</span></span> |
| [<span data-ttu-id="92251-158">recurrence</span><span class="sxs-lookup"><span data-stu-id="92251-158">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="92251-159">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-159">Member</span></span> |
| [<span data-ttu-id="92251-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="92251-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="92251-161">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-161">Member</span></span> |
| [<span data-ttu-id="92251-162">sender</span><span class="sxs-lookup"><span data-stu-id="92251-162">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="92251-163">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-163">Member</span></span> |
| [<span data-ttu-id="92251-164">seriesId</span><span class="sxs-lookup"><span data-stu-id="92251-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="92251-165">Member</span><span class="sxs-lookup"><span data-stu-id="92251-165">Member</span></span> |
| [<span data-ttu-id="92251-166">start</span><span class="sxs-lookup"><span data-stu-id="92251-166">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="92251-167">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-167">Member</span></span> |
| [<span data-ttu-id="92251-168">subject</span><span class="sxs-lookup"><span data-stu-id="92251-168">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="92251-169">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-169">Member</span></span> |
| [<span data-ttu-id="92251-170">to</span><span class="sxs-lookup"><span data-stu-id="92251-170">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="92251-171">Membre</span><span class="sxs-lookup"><span data-stu-id="92251-171">Member</span></span> |
| [<span data-ttu-id="92251-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="92251-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="92251-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-173">Method</span></span> |
| [<span data-ttu-id="92251-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="92251-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="92251-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-175">Method</span></span> |
| [<span data-ttu-id="92251-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="92251-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="92251-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-177">Method</span></span> |
| [<span data-ttu-id="92251-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="92251-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="92251-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-179">Method</span></span> |
| [<span data-ttu-id="92251-180">close</span><span class="sxs-lookup"><span data-stu-id="92251-180">close</span></span>](#close) | <span data-ttu-id="92251-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-181">Method</span></span> |
| [<span data-ttu-id="92251-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="92251-182">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="92251-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-183">Method</span></span> |
| [<span data-ttu-id="92251-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="92251-184">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="92251-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-185">Method</span></span> |
| [<span data-ttu-id="92251-186">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="92251-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="92251-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-187">Method</span></span> |
| [<span data-ttu-id="92251-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="92251-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="92251-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-189">Method</span></span> |
| [<span data-ttu-id="92251-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="92251-190">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="92251-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-191">Method</span></span> |
| [<span data-ttu-id="92251-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="92251-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="92251-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-193">Method</span></span> |
| [<span data-ttu-id="92251-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="92251-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="92251-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-195">Method</span></span> |
| [<span data-ttu-id="92251-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="92251-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="92251-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-197">Method</span></span> |
| [<span data-ttu-id="92251-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="92251-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="92251-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-199">Method</span></span> |
| [<span data-ttu-id="92251-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="92251-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="92251-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-201">Method</span></span> |
| [<span data-ttu-id="92251-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="92251-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="92251-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-203">Method</span></span> |
| [<span data-ttu-id="92251-204">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="92251-204">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="92251-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-205">Method</span></span> |
| [<span data-ttu-id="92251-206">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="92251-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="92251-207">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-207">Method</span></span> |
| [<span data-ttu-id="92251-208">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="92251-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="92251-209">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-209">Method</span></span> |
| [<span data-ttu-id="92251-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="92251-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="92251-211">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-211">Method</span></span> |
| [<span data-ttu-id="92251-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="92251-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="92251-213">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-213">Method</span></span> |
| [<span data-ttu-id="92251-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="92251-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="92251-215">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-215">Method</span></span> |
| [<span data-ttu-id="92251-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="92251-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="92251-217">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-217">Method</span></span> |
| [<span data-ttu-id="92251-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="92251-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="92251-219">Méthode</span><span class="sxs-lookup"><span data-stu-id="92251-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="92251-220">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-220">Example</span></span>

<span data-ttu-id="92251-221">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="92251-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="92251-222">Membres</span><span class="sxs-lookup"><span data-stu-id="92251-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="92251-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="92251-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="92251-224">Permet d’obtenir les pièces jointes de l’élément sous forme de tableau.</span><span class="sxs-lookup"><span data-stu-id="92251-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="92251-225">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-226">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="92251-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="92251-227">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="92251-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="92251-228">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-228">Type:</span></span>

*   <span data-ttu-id="92251-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="92251-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-230">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-230">Requirements</span></span>

|<span data-ttu-id="92251-231">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-231">Requirement</span></span>|<span data-ttu-id="92251-232">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-233">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-234">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-234">1.0</span></span>|
|[<span data-ttu-id="92251-235">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-235">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-236">ReadItem</span></span>|
|[<span data-ttu-id="92251-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-237">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-238">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-239">Example</span></span>

<span data-ttu-id="92251-240">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="92251-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="92251-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="92251-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="92251-242">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="92251-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="92251-243">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-244">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-244">Type:</span></span>

*   [<span data-ttu-id="92251-245">Destinataires</span><span class="sxs-lookup"><span data-stu-id="92251-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="92251-246">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-246">Requirements</span></span>

|<span data-ttu-id="92251-247">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-247">Requirement</span></span>|<span data-ttu-id="92251-248">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-249">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-250">1.1</span><span class="sxs-lookup"><span data-stu-id="92251-250">1.1</span></span>|
|[<span data-ttu-id="92251-251">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-252">ReadItem</span></span>|
|[<span data-ttu-id="92251-253">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-254">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-255">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-255">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="92251-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="92251-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="92251-257">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="92251-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-258">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-258">Type:</span></span>

*   [<span data-ttu-id="92251-259">Corps</span><span class="sxs-lookup"><span data-stu-id="92251-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="92251-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-260">Requirements</span></span>

|<span data-ttu-id="92251-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-261">Requirement</span></span>|<span data-ttu-id="92251-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-264">1.1</span><span class="sxs-lookup"><span data-stu-id="92251-264">1.1</span></span>|
|[<span data-ttu-id="92251-265">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-266">ReadItem</span></span>|
|[<span data-ttu-id="92251-267">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-268">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-268">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="92251-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="92251-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="92251-270">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="92251-270">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="92251-271">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="92251-271">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-272">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-272">Read mode</span></span>

<span data-ttu-id="92251-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="92251-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="92251-275">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-275">Compose mode</span></span>

<span data-ttu-id="92251-276">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="92251-276">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-277">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-277">Type:</span></span>

*   <span data-ttu-id="92251-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="92251-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-279">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-279">Requirements</span></span>

|<span data-ttu-id="92251-280">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-280">Requirement</span></span>|<span data-ttu-id="92251-281">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-282">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-283">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-283">1.0</span></span>|
|[<span data-ttu-id="92251-284">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-284">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-285">ReadItem</span></span>|
|[<span data-ttu-id="92251-286">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-286">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-287">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-287">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-288">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-288">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="92251-289">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="92251-289">(nullable) conversationId :String</span></span>

<span data-ttu-id="92251-290">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="92251-290">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="92251-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="92251-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="92251-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="92251-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-295">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-295">Type:</span></span>

*   <span data-ttu-id="92251-296">Chaîne</span><span class="sxs-lookup"><span data-stu-id="92251-296">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-297">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-297">Requirements</span></span>

|<span data-ttu-id="92251-298">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-298">Requirement</span></span>|<span data-ttu-id="92251-299">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-300">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-301">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-301">1.0</span></span>|
|[<span data-ttu-id="92251-302">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-303">ReadItem</span></span>|
|[<span data-ttu-id="92251-304">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-305">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-305">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="92251-306">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="92251-306">dateTimeCreated :Date</span></span>

<span data-ttu-id="92251-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-309">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-309">Type:</span></span>

*   <span data-ttu-id="92251-310">Date</span><span class="sxs-lookup"><span data-stu-id="92251-310">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-311">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-311">Requirements</span></span>

|<span data-ttu-id="92251-312">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-312">Requirement</span></span>|<span data-ttu-id="92251-313">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-314">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-315">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-315">1.0</span></span>|
|[<span data-ttu-id="92251-316">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-316">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-317">ReadItem</span></span>|
|[<span data-ttu-id="92251-318">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-318">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-319">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-319">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-320">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-320">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="92251-321">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="92251-321">dateTimeModified :Date</span></span>

<span data-ttu-id="92251-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-324">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-324">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-325">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-325">Type:</span></span>

*   <span data-ttu-id="92251-326">Date</span><span class="sxs-lookup"><span data-stu-id="92251-326">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-327">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-327">Requirements</span></span>

|<span data-ttu-id="92251-328">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-328">Requirement</span></span>|<span data-ttu-id="92251-329">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-329">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-330">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-331">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-331">1.0</span></span>|
|[<span data-ttu-id="92251-332">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-332">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-333">ReadItem</span></span>|
|[<span data-ttu-id="92251-334">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-334">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-335">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-335">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-336">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-336">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="92251-337">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="92251-337">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="92251-338">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-338">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="92251-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="92251-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-341">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="92251-341">Read mode</span></span>

<span data-ttu-id="92251-342">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="92251-342">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="92251-343">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-343">Compose mode</span></span>

<span data-ttu-id="92251-344">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="92251-344">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="92251-345">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="92251-345">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-346">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-346">Type:</span></span>

*   <span data-ttu-id="92251-347">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="92251-347">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-348">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-348">Requirements</span></span>

|<span data-ttu-id="92251-349">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-349">Requirement</span></span>|<span data-ttu-id="92251-350">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-351">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-352">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-352">1.0</span></span>|
|[<span data-ttu-id="92251-353">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-353">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-354">ReadItem</span></span>|
|[<span data-ttu-id="92251-355">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-355">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-356">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-356">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-357">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-357">Example</span></span>

<span data-ttu-id="92251-358">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="92251-358">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="92251-359">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="92251-359">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="92251-360">Obtient ou définit les emplacements d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-360">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-361">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-361">Read mode</span></span>

<span data-ttu-id="92251-362">Le `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui vous permet d’obtenir l’ensemble des emplacements (chacune représentée par un objet [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associée au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-362">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="92251-363">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-363">Compose mode</span></span>

<span data-ttu-id="92251-364">Le `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui fournit des méthodes pour obtenir, supprimer ou ajouter des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-365">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-365">Type:</span></span>

*   [<span data-ttu-id="92251-366">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="92251-366">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="92251-367">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-367">Requirements</span></span>

|<span data-ttu-id="92251-368">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-368">Requirement</span></span>|<span data-ttu-id="92251-369">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-370">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-371">Aperçu</span><span class="sxs-lookup"><span data-stu-id="92251-371">Preview</span></span>|
|[<span data-ttu-id="92251-372">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-373">ReadItem</span></span>|
|[<span data-ttu-id="92251-374">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-374">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-375">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-375">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-376">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-376">Example</span></span>

<span data-ttu-id="92251-377">L’exemple suivant obtient les emplacements actives associées au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-377">The following example gets the current locations associated with the appointment.</span></span>

```javascript
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type == Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}

// Sample output:
// Display name: Conf Room 14
// Type: room
// Email address: cr14@contoso.com
// Display name: Paris
// Type: custom
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="92251-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="92251-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="92251-379">Permet d’obtenir l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="92251-379">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="92251-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="92251-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-382">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="92251-382">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-383">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="92251-383">Read mode</span></span>

<span data-ttu-id="92251-384">La propriété `from` renvoie un objet `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="92251-384">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="92251-385">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-385">Compose mode</span></span>

<span data-ttu-id="92251-386">La propriété `from` renvoie un objet `From` qui fournit une méthode pour obtenir la valeur from.</span><span class="sxs-lookup"><span data-stu-id="92251-386">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="92251-387">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-387">Type:</span></span>

*   <span data-ttu-id="92251-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="92251-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-389">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-389">Requirements</span></span>

|<span data-ttu-id="92251-390">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="92251-391">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-392">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-392">1.0</span></span>|<span data-ttu-id="92251-393">1.7</span><span class="sxs-lookup"><span data-stu-id="92251-393">1.7</span></span>|
|[<span data-ttu-id="92251-394">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-394">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-395">ReadItem</span></span>|<span data-ttu-id="92251-396">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92251-396">ReadWriteItem</span></span>|
|[<span data-ttu-id="92251-397">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-397">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-398">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-398">Read</span></span>|<span data-ttu-id="92251-399">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-399">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="92251-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="92251-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="92251-401">Permet d’obtenir ou de définir les en-têtes Internet d’un message.</span><span class="sxs-lookup"><span data-stu-id="92251-401">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-402">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-402">Type:</span></span>

*   [<span data-ttu-id="92251-403">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="92251-403">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="92251-404">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-404">Requirements</span></span>

|<span data-ttu-id="92251-405">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-405">Requirement</span></span>|<span data-ttu-id="92251-406">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-407">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-408">Aperçu</span><span class="sxs-lookup"><span data-stu-id="92251-408">Preview</span></span>|
|[<span data-ttu-id="92251-409">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-410">ReadItem</span></span>|
|[<span data-ttu-id="92251-411">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-412">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-412">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="92251-413">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="92251-413">internetMessageId :String</span></span>

<span data-ttu-id="92251-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-416">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-416">Type:</span></span>

*   <span data-ttu-id="92251-417">Chaîne</span><span class="sxs-lookup"><span data-stu-id="92251-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-418">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-418">Requirements</span></span>

|<span data-ttu-id="92251-419">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-419">Requirement</span></span>|<span data-ttu-id="92251-420">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-421">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-422">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-422">1.0</span></span>|
|[<span data-ttu-id="92251-423">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-424">ReadItem</span></span>|
|[<span data-ttu-id="92251-425">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-426">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-427">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-427">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="92251-428">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="92251-428">itemClass :String</span></span>

<span data-ttu-id="92251-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="92251-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="92251-433">Type</span><span class="sxs-lookup"><span data-stu-id="92251-433">Type</span></span>|<span data-ttu-id="92251-434">Description</span><span class="sxs-lookup"><span data-stu-id="92251-434">Description</span></span>|<span data-ttu-id="92251-435">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="92251-435">item class</span></span>|
|---|---|---|
|<span data-ttu-id="92251-436">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="92251-436">Appointment items</span></span>|<span data-ttu-id="92251-437">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="92251-437">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="92251-438">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="92251-438">Message items</span></span>|<span data-ttu-id="92251-439">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="92251-439">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="92251-440">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="92251-440">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-441">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-441">Type:</span></span>

*   <span data-ttu-id="92251-442">Chaîne</span><span class="sxs-lookup"><span data-stu-id="92251-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-443">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-443">Requirements</span></span>

|<span data-ttu-id="92251-444">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-444">Requirement</span></span>|<span data-ttu-id="92251-445">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-446">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-447">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-447">1.0</span></span>|
|[<span data-ttu-id="92251-448">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-449">ReadItem</span></span>|
|[<span data-ttu-id="92251-450">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-451">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-452">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-452">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="92251-453">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="92251-453">(nullable) itemId :String</span></span>

<span data-ttu-id="92251-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-456">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="92251-456">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="92251-457">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="92251-457">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="92251-458">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="92251-458">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="92251-459">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="92251-459">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="92251-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-462">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-462">Type:</span></span>

*   <span data-ttu-id="92251-463">Chaîne</span><span class="sxs-lookup"><span data-stu-id="92251-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-464">Requirements</span></span>

|<span data-ttu-id="92251-465">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-465">Requirement</span></span>|<span data-ttu-id="92251-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-468">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-468">1.0</span></span>|
|[<span data-ttu-id="92251-469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-470">ReadItem</span></span>|
|[<span data-ttu-id="92251-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-472">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-473">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-473">Example</span></span>

<span data-ttu-id="92251-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="92251-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="92251-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="92251-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="92251-477">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="92251-477">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="92251-478">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-478">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-479">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-479">Type:</span></span>

*   [<span data-ttu-id="92251-480">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="92251-480">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="92251-481">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-481">Requirements</span></span>

|<span data-ttu-id="92251-482">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-482">Requirement</span></span>|<span data-ttu-id="92251-483">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-484">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-485">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-485">1.0</span></span>|
|[<span data-ttu-id="92251-486">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-487">ReadItem</span></span>|
|[<span data-ttu-id="92251-488">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-489">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-489">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-490">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-490">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="92251-491">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="92251-491">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="92251-492">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-492">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-493">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-493">Read mode</span></span>

<span data-ttu-id="92251-494">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-494">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="92251-495">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-495">Compose mode</span></span>

<span data-ttu-id="92251-496">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-496">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-497">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-497">Type:</span></span>

*   <span data-ttu-id="92251-498">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="92251-498">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-499">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-499">Requirements</span></span>

|<span data-ttu-id="92251-500">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-500">Requirement</span></span>|<span data-ttu-id="92251-501">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-502">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-503">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-503">1.0</span></span>|
|[<span data-ttu-id="92251-504">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-505">ReadItem</span></span>|
|[<span data-ttu-id="92251-506">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-507">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-507">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-508">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-508">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="92251-509">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="92251-509">normalizedSubject :String</span></span>

<span data-ttu-id="92251-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="92251-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="92251-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-514">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-514">Type:</span></span>

*   <span data-ttu-id="92251-515">Chaîne</span><span class="sxs-lookup"><span data-stu-id="92251-515">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-516">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-516">Requirements</span></span>

|<span data-ttu-id="92251-517">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-517">Requirement</span></span>|<span data-ttu-id="92251-518">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-519">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-520">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-520">1.0</span></span>|
|[<span data-ttu-id="92251-521">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-521">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-522">ReadItem</span></span>|
|[<span data-ttu-id="92251-523">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-523">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-524">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-524">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-525">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-525">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="92251-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="92251-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="92251-527">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="92251-527">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-528">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-528">Type:</span></span>

*   [<span data-ttu-id="92251-529">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="92251-529">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="92251-530">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-530">Requirements</span></span>

|<span data-ttu-id="92251-531">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-531">Requirement</span></span>|<span data-ttu-id="92251-532">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-532">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-533">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-533">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-534">1.3</span><span class="sxs-lookup"><span data-stu-id="92251-534">1.3</span></span>|
|[<span data-ttu-id="92251-535">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-535">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-536">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-536">ReadItem</span></span>|
|[<span data-ttu-id="92251-537">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-537">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-538">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-538">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="92251-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="92251-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="92251-540">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="92251-540">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="92251-541">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="92251-541">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-542">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-542">Read mode</span></span>

<span data-ttu-id="92251-543">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="92251-543">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="92251-544">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-544">Compose mode</span></span>

<span data-ttu-id="92251-545">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="92251-545">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-546">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-546">Type:</span></span>

*   <span data-ttu-id="92251-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="92251-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-548">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-548">Requirements</span></span>

|<span data-ttu-id="92251-549">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-549">Requirement</span></span>|<span data-ttu-id="92251-550">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-551">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-552">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-552">1.0</span></span>|
|[<span data-ttu-id="92251-553">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-554">ReadItem</span></span>|
|[<span data-ttu-id="92251-555">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-556">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-556">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-557">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-557">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="92251-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="92251-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="92251-559">Permet d’obtenir l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="92251-559">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-560">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-560">Read mode</span></span>

<span data-ttu-id="92251-561">La propriété `organizer` renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="92251-561">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="92251-562">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-562">Compose mode</span></span>

<span data-ttu-id="92251-563">La propriété `organizer` renvoie un objet [Organizer](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur organizer.</span><span class="sxs-lookup"><span data-stu-id="92251-563">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-564">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-564">Type:</span></span>

*   <span data-ttu-id="92251-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="92251-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-566">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-566">Requirements</span></span>

|<span data-ttu-id="92251-567">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-567">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="92251-568">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-569">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-569">1.0</span></span>|<span data-ttu-id="92251-570">1.7</span><span class="sxs-lookup"><span data-stu-id="92251-570">1.7</span></span>|
|[<span data-ttu-id="92251-571">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-572">ReadItem</span></span>|<span data-ttu-id="92251-573">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92251-573">ReadWriteItem</span></span>|
|[<span data-ttu-id="92251-574">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-575">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-575">Read</span></span>|<span data-ttu-id="92251-576">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-576">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-577">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-577">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="92251-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="92251-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="92251-579">Permet d’obtenir ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-579">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="92251-580">Permet d’obtenir la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="92251-580">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="92251-581">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-581">Read and compose modes for appointment items.</span></span> <span data-ttu-id="92251-582">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="92251-582">Read mode for meeting request items.</span></span>

<span data-ttu-id="92251-583">La propriété `recurrence` renvoie un objet [périodicité](/javascript/api/outlook/office.recurrence) pour des demandes de réunions ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="92251-583">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="92251-584">La valeur `null` est renvoyée pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="92251-584">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="92251-585">La valeur `undefined` est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="92251-585">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="92251-586">Remarque : les demandes de réunion ont une valeur `itemClass` d’IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="92251-586">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="92251-587">Remarque : si l’objet de périodicité est `null`, cela indique que l’objet est un rendez-vous unique ou une demande de réunion de rendez-vous unique, et NON une partie d’une série.</span><span class="sxs-lookup"><span data-stu-id="92251-587">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-588">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-588">Type:</span></span>

* [<span data-ttu-id="92251-589">Recurrence</span><span class="sxs-lookup"><span data-stu-id="92251-589">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="92251-590">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-590">Requirement</span></span>|<span data-ttu-id="92251-591">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-592">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-593">1.7</span><span class="sxs-lookup"><span data-stu-id="92251-593">1.7</span></span>|
|[<span data-ttu-id="92251-594">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-595">ReadItem</span></span>|
|[<span data-ttu-id="92251-596">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-597">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-597">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="92251-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="92251-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="92251-599">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="92251-599">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="92251-600">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="92251-600">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-601">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-601">Read mode</span></span>

<span data-ttu-id="92251-602">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="92251-602">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="92251-603">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-603">Compose mode</span></span>

<span data-ttu-id="92251-604">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="92251-604">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-605">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-605">Type:</span></span>

*   <span data-ttu-id="92251-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="92251-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-607">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-607">Requirements</span></span>

|<span data-ttu-id="92251-608">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-608">Requirement</span></span>|<span data-ttu-id="92251-609">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-609">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-610">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-611">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-611">1.0</span></span>|
|[<span data-ttu-id="92251-612">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-612">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-613">ReadItem</span></span>|
|[<span data-ttu-id="92251-614">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-614">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-615">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-615">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-616">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-616">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="92251-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="92251-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="92251-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="92251-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="92251-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-622">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="92251-622">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-623">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-623">Type:</span></span>

*   [<span data-ttu-id="92251-624">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="92251-624">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="92251-625">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-625">Requirements</span></span>

|<span data-ttu-id="92251-626">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-626">Requirement</span></span>|<span data-ttu-id="92251-627">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-628">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-629">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-629">1.0</span></span>|
|[<span data-ttu-id="92251-630">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-631">ReadItem</span></span>|
|[<span data-ttu-id="92251-632">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-633">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-633">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-634">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-634">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="92251-635">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="92251-635">(nullable) seriesId :String</span></span>

<span data-ttu-id="92251-636">Permet d’obtenir l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="92251-636">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="92251-637">Dans OWA et Outlook, `seriesId` renvoie l’identificateur de services web Exchange (EWS) de l’élément (series) parent auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="92251-637">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="92251-638">Dans iOS et Android, `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="92251-638">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-639">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="92251-639">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="92251-640">La propriété `seriesId` n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="92251-640">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="92251-641">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="92251-641">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="92251-642">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="92251-642">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="92251-643">La propriété `seriesId` renvoie `null` pour les éléments qui n’ont pas d’élément parent, tels que des rendez-vous uniques, des éléments de séries ou des demandes de réunion, et renvoie `undefined` pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="92251-643">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-644">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-644">Type:</span></span>

* <span data-ttu-id="92251-645">Chaîne</span><span class="sxs-lookup"><span data-stu-id="92251-645">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-646">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-646">Requirements</span></span>

|<span data-ttu-id="92251-647">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-647">Requirement</span></span>|<span data-ttu-id="92251-648">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-649">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-650">1.7</span><span class="sxs-lookup"><span data-stu-id="92251-650">1.7</span></span>|
|[<span data-ttu-id="92251-651">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-652">ReadItem</span></span>|
|[<span data-ttu-id="92251-653">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-654">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-655">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-655">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="92251-656">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="92251-656">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="92251-657">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-657">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="92251-p130">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="92251-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-660">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-660">Read mode</span></span>

<span data-ttu-id="92251-661">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="92251-661">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="92251-662">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-662">Compose mode</span></span>

<span data-ttu-id="92251-663">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="92251-663">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="92251-664">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="92251-664">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-665">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-665">Type:</span></span>

*   <span data-ttu-id="92251-666">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="92251-666">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-667">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-667">Requirements</span></span>

|<span data-ttu-id="92251-668">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-668">Requirement</span></span>|<span data-ttu-id="92251-669">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-670">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-671">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-671">1.0</span></span>|
|[<span data-ttu-id="92251-672">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-673">ReadItem</span></span>|
|[<span data-ttu-id="92251-674">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-675">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-675">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-676">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-676">Example</span></span>

<span data-ttu-id="92251-677">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="92251-677">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="92251-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="92251-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="92251-679">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="92251-679">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="92251-680">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="92251-680">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-681">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="92251-681">Read mode</span></span>

<span data-ttu-id="92251-p131">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="92251-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="92251-684">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-684">Compose mode</span></span>

<span data-ttu-id="92251-685">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="92251-685">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="92251-686">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-686">Type:</span></span>

*   <span data-ttu-id="92251-687">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="92251-687">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-688">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-688">Requirements</span></span>

|<span data-ttu-id="92251-689">Requirement</span><span class="sxs-lookup"><span data-stu-id="92251-689">Requirement</span></span>|<span data-ttu-id="92251-690">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-691">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-692">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-692">1.0</span></span>|
|[<span data-ttu-id="92251-693">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-694">ReadItem</span></span>|
|[<span data-ttu-id="92251-695">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-696">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-696">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="92251-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="92251-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="92251-698">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="92251-698">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="92251-699">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="92251-699">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92251-700">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-700">Read mode</span></span>

<span data-ttu-id="92251-p133">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="92251-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="92251-703">Mode composition</span><span class="sxs-lookup"><span data-stu-id="92251-703">Compose mode</span></span>

<span data-ttu-id="92251-704">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="92251-704">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="92251-705">Type :</span><span class="sxs-lookup"><span data-stu-id="92251-705">Type:</span></span>

*   <span data-ttu-id="92251-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="92251-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-707">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-707">Requirements</span></span>

|<span data-ttu-id="92251-708">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-708">Requirement</span></span>|<span data-ttu-id="92251-709">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-710">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-711">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-711">1.0</span></span>|
|[<span data-ttu-id="92251-712">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-713">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-713">ReadItem</span></span>|
|[<span data-ttu-id="92251-714">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-715">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-715">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-716">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-716">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="92251-717">Méthodes</span><span class="sxs-lookup"><span data-stu-id="92251-717">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="92251-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92251-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="92251-719">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="92251-719">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="92251-720">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="92251-720">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="92251-721">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="92251-721">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-722">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-722">Parameters:</span></span>
|<span data-ttu-id="92251-723">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-723">Name</span></span>|<span data-ttu-id="92251-724">Type</span><span class="sxs-lookup"><span data-stu-id="92251-724">Type</span></span>|<span data-ttu-id="92251-725">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-725">Attributes</span></span>|<span data-ttu-id="92251-726">Description</span><span class="sxs-lookup"><span data-stu-id="92251-726">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="92251-727">String</span><span class="sxs-lookup"><span data-stu-id="92251-727">String</span></span>||<span data-ttu-id="92251-p134">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="92251-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="92251-730">String</span><span class="sxs-lookup"><span data-stu-id="92251-730">String</span></span>||<span data-ttu-id="92251-p135">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="92251-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="92251-733">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-733">Object</span></span>|<span data-ttu-id="92251-734">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-734">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-735">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-735">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-736">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-736">Object</span></span>|<span data-ttu-id="92251-737">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-737">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-738">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-738">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="92251-739">Boolean</span><span class="sxs-lookup"><span data-stu-id="92251-739">Boolean</span></span>|<span data-ttu-id="92251-740">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-740">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-741">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="92251-741">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="92251-742">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-742">function</span></span>|<span data-ttu-id="92251-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-743">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-744">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="92251-745">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="92251-745">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="92251-746">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="92251-746">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="92251-747">Erreurs</span><span class="sxs-lookup"><span data-stu-id="92251-747">Errors</span></span>

|<span data-ttu-id="92251-748">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="92251-748">Error code</span></span>|<span data-ttu-id="92251-749">Description</span><span class="sxs-lookup"><span data-stu-id="92251-749">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="92251-750">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="92251-750">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="92251-751">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="92251-751">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="92251-752">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="92251-752">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-753">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-753">Requirements</span></span>

|<span data-ttu-id="92251-754">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-754">Requirement</span></span>|<span data-ttu-id="92251-755">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-755">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-756">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-756">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-757">1.1</span><span class="sxs-lookup"><span data-stu-id="92251-757">1.1</span></span>|
|[<span data-ttu-id="92251-758">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-758">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-759">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92251-759">ReadWriteItem</span></span>|
|[<span data-ttu-id="92251-760">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-760">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-761">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-761">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="92251-762">Exemples</span><span class="sxs-lookup"><span data-stu-id="92251-762">Examples</span></span>

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

<span data-ttu-id="92251-763">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="92251-763">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="92251-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92251-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="92251-765">Ajoute un fichier provenant du codage base64 à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="92251-765">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="92251-766">La méthode `addFileAttachmentFromBase64Async` charge le fichier depuis le codage base64 et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="92251-766">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="92251-767">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="92251-767">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="92251-768">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="92251-768">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-769">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-769">Parameters:</span></span>
|<span data-ttu-id="92251-770">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-770">Name</span></span>|<span data-ttu-id="92251-771">Type</span><span class="sxs-lookup"><span data-stu-id="92251-771">Type</span></span>|<span data-ttu-id="92251-772">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-772">Attributes</span></span>|<span data-ttu-id="92251-773">Description</span><span class="sxs-lookup"><span data-stu-id="92251-773">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="92251-774">String</span><span class="sxs-lookup"><span data-stu-id="92251-774">String</span></span>||<span data-ttu-id="92251-775">Contenu codé en base64 d’une image ou d’un fichier à ajouter à un e-mail ou à un événement.</span><span class="sxs-lookup"><span data-stu-id="92251-775">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="92251-776">Chaîne</span><span class="sxs-lookup"><span data-stu-id="92251-776">String</span></span>||<span data-ttu-id="92251-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="92251-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="92251-779">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-779">Object</span></span>|<span data-ttu-id="92251-780">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-780">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-781">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-781">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-782">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-782">Object</span></span>|<span data-ttu-id="92251-783">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-783">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-784">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-784">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="92251-785">Boolean</span><span class="sxs-lookup"><span data-stu-id="92251-785">Boolean</span></span>|<span data-ttu-id="92251-786">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-786">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-787">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="92251-787">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="92251-788">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-788">function</span></span>|<span data-ttu-id="92251-789">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-789">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-790">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="92251-791">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="92251-791">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="92251-792">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="92251-792">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="92251-793">Erreurs</span><span class="sxs-lookup"><span data-stu-id="92251-793">Errors</span></span>

|<span data-ttu-id="92251-794">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="92251-794">Error code</span></span>|<span data-ttu-id="92251-795">Description</span><span class="sxs-lookup"><span data-stu-id="92251-795">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="92251-796">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="92251-796">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="92251-797">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="92251-797">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="92251-798">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="92251-798">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-799">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-799">Requirements</span></span>

|<span data-ttu-id="92251-800">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-800">Requirement</span></span>|<span data-ttu-id="92251-801">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-801">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-802">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-802">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-803">Aperçu</span><span class="sxs-lookup"><span data-stu-id="92251-803">Preview</span></span>|
|[<span data-ttu-id="92251-804">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-804">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-805">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92251-805">ReadWriteItem</span></span>|
|[<span data-ttu-id="92251-806">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-806">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-807">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-807">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="92251-808">Exemples</span><span class="sxs-lookup"><span data-stu-id="92251-808">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="92251-809">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92251-809">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="92251-810">Ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="92251-810">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="92251-811">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="92251-811">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-812">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-812">Parameters:</span></span>

| <span data-ttu-id="92251-813">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-813">Name</span></span> | <span data-ttu-id="92251-814">Type</span><span class="sxs-lookup"><span data-stu-id="92251-814">Type</span></span> | <span data-ttu-id="92251-815">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-815">Attributes</span></span> | <span data-ttu-id="92251-816">Description</span><span class="sxs-lookup"><span data-stu-id="92251-816">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="92251-817">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="92251-817">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="92251-818">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="92251-818">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="92251-819">Fonction</span><span class="sxs-lookup"><span data-stu-id="92251-819">Function</span></span> || <span data-ttu-id="92251-p138">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="92251-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="92251-823">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-823">Object</span></span> | <span data-ttu-id="92251-824">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-824">&lt;optional&gt;</span></span> | <span data-ttu-id="92251-825">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-825">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="92251-826">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-826">Object</span></span> | <span data-ttu-id="92251-827">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-827">&lt;optional&gt;</span></span> | <span data-ttu-id="92251-828">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-828">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="92251-829">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-829">function</span></span>| <span data-ttu-id="92251-830">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-830">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-831">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-831">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-832">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-832">Requirements</span></span>

|<span data-ttu-id="92251-833">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-833">Requirement</span></span>| <span data-ttu-id="92251-834">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-835">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92251-836">1.7</span><span class="sxs-lookup"><span data-stu-id="92251-836">1.7</span></span> |
|[<span data-ttu-id="92251-837">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92251-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-838">ReadItem</span></span> |
|[<span data-ttu-id="92251-839">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92251-840">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-840">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="92251-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92251-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="92251-842">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-842">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="92251-p139">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="92251-846">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="92251-846">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="92251-847">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="92251-847">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-848">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-848">Parameters:</span></span>

|<span data-ttu-id="92251-849">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-849">Name</span></span>|<span data-ttu-id="92251-850">Type</span><span class="sxs-lookup"><span data-stu-id="92251-850">Type</span></span>|<span data-ttu-id="92251-851">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-851">Attributes</span></span>|<span data-ttu-id="92251-852">Description</span><span class="sxs-lookup"><span data-stu-id="92251-852">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="92251-853">String</span><span class="sxs-lookup"><span data-stu-id="92251-853">String</span></span>||<span data-ttu-id="92251-p140">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="92251-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="92251-856">String</span><span class="sxs-lookup"><span data-stu-id="92251-856">String</span></span>||<span data-ttu-id="92251-p141">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="92251-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="92251-859">Object</span><span class="sxs-lookup"><span data-stu-id="92251-859">Object</span></span>|<span data-ttu-id="92251-860">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-860">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-861">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-862">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-862">Object</span></span>|<span data-ttu-id="92251-863">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-863">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-864">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="92251-865">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-865">function</span></span>|<span data-ttu-id="92251-866">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-866">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-867">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="92251-868">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="92251-868">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="92251-869">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="92251-869">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="92251-870">Erreurs</span><span class="sxs-lookup"><span data-stu-id="92251-870">Errors</span></span>

|<span data-ttu-id="92251-871">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="92251-871">Error code</span></span>|<span data-ttu-id="92251-872">Description</span><span class="sxs-lookup"><span data-stu-id="92251-872">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="92251-873">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="92251-873">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-874">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-874">Requirements</span></span>

|<span data-ttu-id="92251-875">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-875">Requirement</span></span>|<span data-ttu-id="92251-876">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-876">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-877">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-877">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-878">1.1</span><span class="sxs-lookup"><span data-stu-id="92251-878">1.1</span></span>|
|[<span data-ttu-id="92251-879">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-879">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-880">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92251-880">ReadWriteItem</span></span>|
|[<span data-ttu-id="92251-881">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-881">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-882">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-882">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-883">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-883">Example</span></span>

<span data-ttu-id="92251-884">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="92251-884">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="92251-885">close()</span><span class="sxs-lookup"><span data-stu-id="92251-885">close()</span></span>

<span data-ttu-id="92251-886">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="92251-886">Closes the current item that is being composed.</span></span>

<span data-ttu-id="92251-p142">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="92251-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-889">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-889">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="92251-890">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="92251-890">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-891">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-891">Requirements</span></span>

|<span data-ttu-id="92251-892">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-892">Requirement</span></span>|<span data-ttu-id="92251-893">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-893">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-894">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-894">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-895">1.3</span><span class="sxs-lookup"><span data-stu-id="92251-895">1.3</span></span>|
|[<span data-ttu-id="92251-896">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-896">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-897">Restreinte</span><span class="sxs-lookup"><span data-stu-id="92251-897">Restricted</span></span>|
|[<span data-ttu-id="92251-898">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-898">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-899">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-899">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="92251-900">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="92251-900">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="92251-901">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="92251-901">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-902">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-902">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="92251-903">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="92251-903">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="92251-904">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="92251-904">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="92251-p143">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="92251-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-908">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-908">Parameters:</span></span>

|<span data-ttu-id="92251-909">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-909">Name</span></span>|<span data-ttu-id="92251-910">Type</span><span class="sxs-lookup"><span data-stu-id="92251-910">Type</span></span>|<span data-ttu-id="92251-911">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-911">Attributes</span></span>|<span data-ttu-id="92251-912">Description</span><span class="sxs-lookup"><span data-stu-id="92251-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="92251-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="92251-913">String &#124; Object</span></span>||<span data-ttu-id="92251-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="92251-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="92251-916">**OU**</span><span class="sxs-lookup"><span data-stu-id="92251-916">**OR**</span></span><br/><span data-ttu-id="92251-p145">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="92251-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="92251-919">String</span><span class="sxs-lookup"><span data-stu-id="92251-919">String</span></span>|<span data-ttu-id="92251-920">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-920">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="92251-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="92251-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="92251-924">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-924">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-925">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="92251-926">String</span><span class="sxs-lookup"><span data-stu-id="92251-926">String</span></span>||<span data-ttu-id="92251-p147">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="92251-929">String</span><span class="sxs-lookup"><span data-stu-id="92251-929">String</span></span>||<span data-ttu-id="92251-930">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="92251-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="92251-931">String</span><span class="sxs-lookup"><span data-stu-id="92251-931">String</span></span>||<span data-ttu-id="92251-p148">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="92251-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="92251-934">Boolean</span><span class="sxs-lookup"><span data-stu-id="92251-934">Boolean</span></span>||<span data-ttu-id="92251-p149">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="92251-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="92251-937">String</span><span class="sxs-lookup"><span data-stu-id="92251-937">String</span></span>||<span data-ttu-id="92251-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="92251-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="92251-941">function</span><span class="sxs-lookup"><span data-stu-id="92251-941">function</span></span>|<span data-ttu-id="92251-942">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-942">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-943">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-944">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-944">Requirements</span></span>

|<span data-ttu-id="92251-945">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-945">Requirement</span></span>|<span data-ttu-id="92251-946">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-947">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-948">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-948">1.0</span></span>|
|[<span data-ttu-id="92251-949">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-949">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-950">ReadItem</span></span>|
|[<span data-ttu-id="92251-951">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-951">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-952">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="92251-953">Exemples</span><span class="sxs-lookup"><span data-stu-id="92251-953">Examples</span></span>

<span data-ttu-id="92251-954">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="92251-954">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="92251-955">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="92251-955">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="92251-956">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="92251-956">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="92251-957">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="92251-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="92251-958">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="92251-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="92251-959">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="92251-960">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="92251-960">displayReplyForm(formData)</span></span>

<span data-ttu-id="92251-961">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="92251-961">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-962">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-962">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="92251-963">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="92251-963">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="92251-964">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="92251-964">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="92251-p151">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="92251-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-968">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-968">Parameters:</span></span>

|<span data-ttu-id="92251-969">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-969">Name</span></span>|<span data-ttu-id="92251-970">Type</span><span class="sxs-lookup"><span data-stu-id="92251-970">Type</span></span>|<span data-ttu-id="92251-971">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-971">Attributes</span></span>|<span data-ttu-id="92251-972">Description</span><span class="sxs-lookup"><span data-stu-id="92251-972">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="92251-973">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="92251-973">String &#124; Object</span></span>||<span data-ttu-id="92251-p152">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="92251-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="92251-976">**OU**</span><span class="sxs-lookup"><span data-stu-id="92251-976">**OR**</span></span><br/><span data-ttu-id="92251-p153">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="92251-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="92251-979">String</span><span class="sxs-lookup"><span data-stu-id="92251-979">String</span></span>|<span data-ttu-id="92251-980">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-980">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="92251-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="92251-983">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-983">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="92251-984">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-984">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-985">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-985">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="92251-986">String</span><span class="sxs-lookup"><span data-stu-id="92251-986">String</span></span>||<span data-ttu-id="92251-p155">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="92251-989">String</span><span class="sxs-lookup"><span data-stu-id="92251-989">String</span></span>||<span data-ttu-id="92251-990">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="92251-990">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="92251-991">String</span><span class="sxs-lookup"><span data-stu-id="92251-991">String</span></span>||<span data-ttu-id="92251-p156">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="92251-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="92251-994">Boolean</span><span class="sxs-lookup"><span data-stu-id="92251-994">Boolean</span></span>||<span data-ttu-id="92251-p157">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="92251-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="92251-997">String</span><span class="sxs-lookup"><span data-stu-id="92251-997">String</span></span>||<span data-ttu-id="92251-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="92251-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="92251-1001">function</span><span class="sxs-lookup"><span data-stu-id="92251-1001">function</span></span>|<span data-ttu-id="92251-1002">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1002">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1003">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1003">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1004">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1004">Requirements</span></span>

|<span data-ttu-id="92251-1005">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1005">Requirement</span></span>|<span data-ttu-id="92251-1006">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1007">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1007">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-1008">1.0</span></span>|
|[<span data-ttu-id="92251-1009">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1010">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1010">ReadItem</span></span>|
|[<span data-ttu-id="92251-1011">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1012">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1012">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="92251-1013">Exemples</span><span class="sxs-lookup"><span data-stu-id="92251-1013">Examples</span></span>

<span data-ttu-id="92251-1014">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="92251-1014">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="92251-1015">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="92251-1015">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="92251-1016">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="92251-1016">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="92251-1017">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="92251-1017">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="92251-1018">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="92251-1018">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="92251-1019">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1019">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="92251-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="92251-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="92251-1021">Permet d’obtenir la pièce jointe spécifiée depuis un message ou un rendez-vous, et la renvoie en tant qu’objet `AttachmentContent`.</span><span class="sxs-lookup"><span data-stu-id="92251-1021">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="92251-1022">La méthode `getAttachmentContentAsync` permet d’obtenir la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-1022">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="92251-1023">Nous vous recommandons de suivre la bonne pratique consistant à utiliser l’identificateur pour récupérer une pièce jointe dans la même session que celle où les objets attachmentIds ont été récupérés avec l’appel `getAttachmentsAsync` ou `item.attachments`.</span><span class="sxs-lookup"><span data-stu-id="92251-1023">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="92251-1024">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="92251-1024">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="92251-1025">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer un formulaire incorporé qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="92251-1025">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1026">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1026">Parameters:</span></span>

|<span data-ttu-id="92251-1027">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1027">Name</span></span>|<span data-ttu-id="92251-1028">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1028">Type</span></span>|<span data-ttu-id="92251-1029">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1029">Attributes</span></span>|<span data-ttu-id="92251-1030">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1030">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="92251-1031">String</span><span class="sxs-lookup"><span data-stu-id="92251-1031">String</span></span>||<span data-ttu-id="92251-1032">Identificateur de la pièce jointe que vous voulez obtenir.</span><span class="sxs-lookup"><span data-stu-id="92251-1032">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="92251-1033">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1033">Object</span></span>|<span data-ttu-id="92251-1034">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1034">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1035">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-1035">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-1036">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1036">Object</span></span>|<span data-ttu-id="92251-1037">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1038">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1038">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="92251-1039">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-1039">function</span></span>|<span data-ttu-id="92251-1040">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1041">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1041">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1042">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1042">Requirements</span></span>

|<span data-ttu-id="92251-1043">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1043">Requirement</span></span>|<span data-ttu-id="92251-1044">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1044">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1045">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1045">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1046">Aperçu</span><span class="sxs-lookup"><span data-stu-id="92251-1046">Preview</span></span>|
|[<span data-ttu-id="92251-1047">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1047">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1048">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1048">ReadItem</span></span>|
|[<span data-ttu-id="92251-1049">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1049">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1050">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1050">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1051">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1051">Returns:</span></span>

<span data-ttu-id="92251-1052">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="92251-1052">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="92251-1053">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1053">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="92251-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="92251-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="92251-1055">Permet d’obtenir les pièces jointes de l’élément sous forme de tableau.</span><span class="sxs-lookup"><span data-stu-id="92251-1055">Gets the item's attachments as an array.</span></span> <span data-ttu-id="92251-1056">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="92251-1056">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1057">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1057">Parameters:</span></span>

|<span data-ttu-id="92251-1058">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1058">Name</span></span>|<span data-ttu-id="92251-1059">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1059">Type</span></span>|<span data-ttu-id="92251-1060">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1060">Attributes</span></span>|<span data-ttu-id="92251-1061">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1061">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="92251-1062">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1062">Object</span></span>|<span data-ttu-id="92251-1063">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1064">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-1065">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1065">Object</span></span>|<span data-ttu-id="92251-1066">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1067">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="92251-1068">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-1068">function</span></span>|<span data-ttu-id="92251-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1070">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1070">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1071">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1071">Requirements</span></span>

|<span data-ttu-id="92251-1072">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1072">Requirement</span></span>|<span data-ttu-id="92251-1073">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1074">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1075">Aperçu</span><span class="sxs-lookup"><span data-stu-id="92251-1075">Preview</span></span>|
|[<span data-ttu-id="92251-1076">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1077">ReadItem</span></span>|
|[<span data-ttu-id="92251-1078">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1079">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-1079">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1080">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1080">Returns:</span></span>

<span data-ttu-id="92251-1081">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="92251-1081">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="92251-1082">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1082">Example</span></span>

<span data-ttu-id="92251-1083">L’exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="92251-1083">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="92251-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="92251-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="92251-1085">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="92251-1085">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1086">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-1086">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-1087">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1087">Requirements</span></span>

|<span data-ttu-id="92251-1088">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1088">Requirement</span></span>|<span data-ttu-id="92251-1089">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1090">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1090">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1091">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-1091">1.0</span></span>|
|[<span data-ttu-id="92251-1092">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1093">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1093">ReadItem</span></span>|
|[<span data-ttu-id="92251-1094">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1095">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1095">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1096">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1096">Returns:</span></span>

<span data-ttu-id="92251-1097">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="92251-1097">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="92251-1098">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1098">Example</span></span>

<span data-ttu-id="92251-1099">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="92251-1099">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="92251-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="92251-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="92251-1101">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="92251-1101">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1102">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-1102">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1103">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1103">Parameters:</span></span>

|<span data-ttu-id="92251-1104">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1104">Name</span></span>|<span data-ttu-id="92251-1105">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1105">Type</span></span>|<span data-ttu-id="92251-1106">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1106">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="92251-1107">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="92251-1107">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="92251-1108">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="92251-1108">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1109">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1109">Requirements</span></span>

|<span data-ttu-id="92251-1110">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1110">Requirement</span></span>|<span data-ttu-id="92251-1111">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1111">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1112">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1112">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1113">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-1113">1.0</span></span>|
|[<span data-ttu-id="92251-1114">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1114">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1115">Restreinte</span><span class="sxs-lookup"><span data-stu-id="92251-1115">Restricted</span></span>|
|[<span data-ttu-id="92251-1116">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1116">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1117">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1117">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1118">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1118">Returns:</span></span>

<span data-ttu-id="92251-1119">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="92251-1119">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="92251-1120">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="92251-1120">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="92251-1121">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="92251-1121">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="92251-1122">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="92251-1122">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="92251-1123">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="92251-1123">Value of `entityType`</span></span>|<span data-ttu-id="92251-1124">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="92251-1124">Type of objects in returned array</span></span>|<span data-ttu-id="92251-1125">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="92251-1125">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="92251-1126">String</span><span class="sxs-lookup"><span data-stu-id="92251-1126">String</span></span>|<span data-ttu-id="92251-1127">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="92251-1127">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="92251-1128">Contact</span><span class="sxs-lookup"><span data-stu-id="92251-1128">Contact</span></span>|<span data-ttu-id="92251-1129">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="92251-1129">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="92251-1130">String</span><span class="sxs-lookup"><span data-stu-id="92251-1130">String</span></span>|<span data-ttu-id="92251-1131">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="92251-1131">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="92251-1132">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="92251-1132">MeetingSuggestion</span></span>|<span data-ttu-id="92251-1133">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="92251-1133">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="92251-1134">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="92251-1134">PhoneNumber</span></span>|<span data-ttu-id="92251-1135">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="92251-1135">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="92251-1136">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="92251-1136">TaskSuggestion</span></span>|<span data-ttu-id="92251-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="92251-1137">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="92251-1138">String</span><span class="sxs-lookup"><span data-stu-id="92251-1138">String</span></span>|<span data-ttu-id="92251-1139">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="92251-1139">**Restricted**</span></span>|

<span data-ttu-id="92251-1140">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="92251-1140">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="92251-1141">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1141">Example</span></span>

<span data-ttu-id="92251-1142">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="92251-1142">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="92251-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="92251-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="92251-1144">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="92251-1144">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1145">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-1145">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="92251-1146">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="92251-1146">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1147">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1147">Parameters:</span></span>

|<span data-ttu-id="92251-1148">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1148">Name</span></span>|<span data-ttu-id="92251-1149">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1149">Type</span></span>|<span data-ttu-id="92251-1150">object</span><span class="sxs-lookup"><span data-stu-id="92251-1150">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="92251-1151">String</span><span class="sxs-lookup"><span data-stu-id="92251-1151">String</span></span>|<span data-ttu-id="92251-1152">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="92251-1152">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1153">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1153">Requirements</span></span>

|<span data-ttu-id="92251-1154">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1154">Requirement</span></span>|<span data-ttu-id="92251-1155">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1155">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1156">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1157">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-1157">1.0</span></span>|
|[<span data-ttu-id="92251-1158">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1159">ReadItem</span></span>|
|[<span data-ttu-id="92251-1160">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1161">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1161">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1162">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1162">Returns:</span></span>

<span data-ttu-id="92251-p162">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="92251-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="92251-1165">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="92251-1165">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="92251-1166">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92251-1166">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="92251-1167">Récupère les données d’initialisation transmises quand le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="92251-1167">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1168">Cette méthode est uniquement prise en charge par Outlook 2016 ou version ultérieure pour Windows (versions en un clic supérieures à 16.0.8413.1000) et Outlook sur le web pour Office 365.</span><span class="sxs-lookup"><span data-stu-id="92251-1168">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1169">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1169">Parameters:</span></span>
|<span data-ttu-id="92251-1170">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1170">Name</span></span>|<span data-ttu-id="92251-1171">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1171">Type</span></span>|<span data-ttu-id="92251-1172">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1172">Attributes</span></span>|<span data-ttu-id="92251-1173">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1173">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="92251-1174">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1174">Object</span></span>|<span data-ttu-id="92251-1175">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1175">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1176">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-1176">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-1177">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1177">Object</span></span>|<span data-ttu-id="92251-1178">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1178">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1179">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1179">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="92251-1180">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-1180">function</span></span>|<span data-ttu-id="92251-1181">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1182">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1182">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="92251-1183">En cas de réussite, les données d’initialisation sont fournies dans la propriété `asyncResult.value` sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="92251-1183">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="92251-1184">S’il n’existe aucun contexte d’initialisation, l’objet `asyncResult` contient un objet `Error` dont la propriété `code` est définie sur `9020` et la propriété `name` sur `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="92251-1184">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1185">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1185">Requirements</span></span>

|<span data-ttu-id="92251-1186">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1186">Requirement</span></span>|<span data-ttu-id="92251-1187">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1188">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1189">Aperçu</span><span class="sxs-lookup"><span data-stu-id="92251-1189">Preview</span></span>|
|[<span data-ttu-id="92251-1190">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1191">ReadItem</span></span>|
|[<span data-ttu-id="92251-1192">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1193">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1193">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-1194">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1194">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="92251-1195">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="92251-1195">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="92251-1196">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="92251-1196">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1197">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-1197">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="92251-p163">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="92251-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="92251-1201">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="92251-1201">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="92251-1202">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="92251-1202">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="92251-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-1206">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1206">Requirements</span></span>

|<span data-ttu-id="92251-1207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1207">Requirement</span></span>|<span data-ttu-id="92251-1208">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1208">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1209">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1210">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-1210">1.0</span></span>|
|[<span data-ttu-id="92251-1211">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1212">ReadItem</span></span>|
|[<span data-ttu-id="92251-1213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1214">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1214">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1215">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1215">Returns:</span></span>

<span data-ttu-id="92251-p165">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="92251-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="92251-1218">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="92251-1218">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="92251-1219">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1219">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="92251-1220">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1220">Example</span></span>

<span data-ttu-id="92251-1221">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="92251-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="92251-1222">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="92251-1222">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="92251-1223">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="92251-1223">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1224">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-1224">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="92251-1225">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="92251-1225">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="92251-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="92251-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1228">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1228">Parameters:</span></span>

|<span data-ttu-id="92251-1229">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1229">Name</span></span>|<span data-ttu-id="92251-1230">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1230">Type</span></span>|<span data-ttu-id="92251-1231">object</span><span class="sxs-lookup"><span data-stu-id="92251-1231">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="92251-1232">String</span><span class="sxs-lookup"><span data-stu-id="92251-1232">String</span></span>|<span data-ttu-id="92251-1233">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="92251-1233">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1234">Requirements</span></span>

|<span data-ttu-id="92251-1235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1235">Requirement</span></span>|<span data-ttu-id="92251-1236">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1236">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1238">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-1238">1.0</span></span>|
|[<span data-ttu-id="92251-1239">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1240">ReadItem</span></span>|
|[<span data-ttu-id="92251-1241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1242">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1242">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1243">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1243">Returns:</span></span>

<span data-ttu-id="92251-1244">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="92251-1244">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="92251-1245">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="92251-1245">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="92251-1246">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="92251-1246">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="92251-1247">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1247">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="92251-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="92251-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="92251-1249">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="92251-1249">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="92251-p167">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="92251-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1252">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1252">Parameters:</span></span>

|<span data-ttu-id="92251-1253">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1253">Name</span></span>|<span data-ttu-id="92251-1254">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1254">Type</span></span>|<span data-ttu-id="92251-1255">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1255">Attributes</span></span>|<span data-ttu-id="92251-1256">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1256">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="92251-1257">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="92251-1257">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="92251-p168">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="92251-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="92251-1261">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1261">Object</span></span>|<span data-ttu-id="92251-1262">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1262">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1263">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-1263">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-1264">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1264">Object</span></span>|<span data-ttu-id="92251-1265">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1266">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1266">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="92251-1267">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-1267">function</span></span>||<span data-ttu-id="92251-1268">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1268">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="92251-1269">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="92251-1269">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="92251-1270">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="92251-1270">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1271">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1271">Requirements</span></span>

|<span data-ttu-id="92251-1272">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1272">Requirement</span></span>|<span data-ttu-id="92251-1273">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1274">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1275">1.2</span><span class="sxs-lookup"><span data-stu-id="92251-1275">1.2</span></span>|
|[<span data-ttu-id="92251-1276">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92251-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="92251-1278">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1279">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-1279">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1280">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1280">Returns:</span></span>

<span data-ttu-id="92251-1281">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="92251-1281">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="92251-1282">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="92251-1282">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="92251-1283">String</span><span class="sxs-lookup"><span data-stu-id="92251-1283">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="92251-1284">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1284">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="92251-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="92251-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="92251-p170">Permet d’obtenir les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="92251-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1288">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-1288">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-1289">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1289">Requirements</span></span>

|<span data-ttu-id="92251-1290">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1290">Requirement</span></span>|<span data-ttu-id="92251-1291">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1291">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1292">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1293">1.6</span><span class="sxs-lookup"><span data-stu-id="92251-1293">1.6</span></span>|
|[<span data-ttu-id="92251-1294">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1295">ReadItem</span></span>|
|[<span data-ttu-id="92251-1296">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1297">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1297">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1298">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1298">Returns:</span></span>

<span data-ttu-id="92251-1299">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="92251-1299">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="92251-1300">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1300">Example</span></span>

<span data-ttu-id="92251-1301">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="92251-1301">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="92251-1302">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="92251-1302">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="92251-p171">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="92251-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1305">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="92251-1305">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="92251-p172">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="92251-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="92251-1309">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="92251-1309">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="92251-1310">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="92251-1310">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="92251-p173">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="92251-1314">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1314">Requirements</span></span>

|<span data-ttu-id="92251-1315">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1315">Requirement</span></span>|<span data-ttu-id="92251-1316">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1316">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1317">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1318">1.6</span><span class="sxs-lookup"><span data-stu-id="92251-1318">1.6</span></span>|
|[<span data-ttu-id="92251-1319">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1319">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1320">ReadItem</span></span>|
|[<span data-ttu-id="92251-1321">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1321">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1322">Lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1322">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92251-1323">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="92251-1323">Returns:</span></span>

<span data-ttu-id="92251-p174">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="92251-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="92251-1326">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1326">Example</span></span>

<span data-ttu-id="92251-1327">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="92251-1327">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="92251-1328">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="92251-1328">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="92251-1329">Permet d’obtenir les propriétés du rendez-vous ou du message sélectionné dans une boîte aux lettres, un calendrier ou un dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="92251-1329">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1330">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1330">Parameters:</span></span>

|<span data-ttu-id="92251-1331">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1331">Name</span></span>|<span data-ttu-id="92251-1332">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1332">Type</span></span>|<span data-ttu-id="92251-1333">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1333">Attributes</span></span>|<span data-ttu-id="92251-1334">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1334">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="92251-1335">Object</span><span class="sxs-lookup"><span data-stu-id="92251-1335">Object</span></span>|<span data-ttu-id="92251-1336">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1336">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1337">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-1337">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-1338">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1338">Object</span></span>|<span data-ttu-id="92251-1339">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1339">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1340">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1340">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="92251-1341">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-1341">function</span></span>||<span data-ttu-id="92251-1342">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="92251-1343">Les propriétés partagées sont fournies sous la forme d’un objet [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="92251-1343">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="92251-1344">Cet objet peut être utilisé pour obtenir des propriétés partagées de l’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-1344">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1345">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1345">Requirements</span></span>

|<span data-ttu-id="92251-1346">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1346">Requirement</span></span>|<span data-ttu-id="92251-1347">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1347">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1348">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1349">Aperçu</span><span class="sxs-lookup"><span data-stu-id="92251-1349">Preview</span></span>|
|[<span data-ttu-id="92251-1350">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1351">ReadItem</span></span>|
|[<span data-ttu-id="92251-1352">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1353">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-1354">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1354">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="92251-1355">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="92251-1355">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="92251-1356">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="92251-1356">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="92251-p176">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="92251-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1360">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1360">Parameters:</span></span>

|<span data-ttu-id="92251-1361">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1361">Name</span></span>|<span data-ttu-id="92251-1362">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1362">Type</span></span>|<span data-ttu-id="92251-1363">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1363">Attributes</span></span>|<span data-ttu-id="92251-1364">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1364">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="92251-1365">function</span><span class="sxs-lookup"><span data-stu-id="92251-1365">function</span></span>||<span data-ttu-id="92251-1366">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="92251-1367">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="92251-1367">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="92251-1368">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="92251-1368">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="92251-1369">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1369">Object</span></span>|<span data-ttu-id="92251-1370">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1371">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1371">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="92251-1372">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1372">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1373">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1373">Requirements</span></span>

|<span data-ttu-id="92251-1374">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1374">Requirement</span></span>|<span data-ttu-id="92251-1375">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1375">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1376">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1377">1.0</span><span class="sxs-lookup"><span data-stu-id="92251-1377">1.0</span></span>|
|[<span data-ttu-id="92251-1378">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1379">ReadItem</span></span>|
|[<span data-ttu-id="92251-1380">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1381">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1381">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-1382">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1382">Example</span></span>

<span data-ttu-id="92251-p179">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="92251-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="92251-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92251-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="92251-1387">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="92251-1387">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="92251-1388">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-1388">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="92251-1389">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="92251-1389">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="92251-1390">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="92251-1390">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="92251-1391">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer un formulaire incorporé qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="92251-1391">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1392">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1392">Parameters:</span></span>

|<span data-ttu-id="92251-1393">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1393">Name</span></span>|<span data-ttu-id="92251-1394">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1394">Type</span></span>|<span data-ttu-id="92251-1395">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1395">Attributes</span></span>|<span data-ttu-id="92251-1396">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1396">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="92251-1397">String</span><span class="sxs-lookup"><span data-stu-id="92251-1397">String</span></span>||<span data-ttu-id="92251-1398">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="92251-1398">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="92251-1399">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1399">Object</span></span>|<span data-ttu-id="92251-1400">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1400">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1401">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-1401">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-1402">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1402">Object</span></span>|<span data-ttu-id="92251-1403">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1403">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1404">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1404">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="92251-1405">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-1405">function</span></span>|<span data-ttu-id="92251-1406">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1407">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1407">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="92251-1408">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="92251-1408">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="92251-1409">Erreurs</span><span class="sxs-lookup"><span data-stu-id="92251-1409">Errors</span></span>

|<span data-ttu-id="92251-1410">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="92251-1410">Error code</span></span>|<span data-ttu-id="92251-1411">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1411">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="92251-1412">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="92251-1412">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1413">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1413">Requirements</span></span>

|<span data-ttu-id="92251-1414">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1414">Requirement</span></span>|<span data-ttu-id="92251-1415">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1415">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1416">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1417">1.1</span><span class="sxs-lookup"><span data-stu-id="92251-1417">1.1</span></span>|
|[<span data-ttu-id="92251-1418">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1418">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1419">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92251-1419">ReadWriteItem</span></span>|
|[<span data-ttu-id="92251-1420">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1420">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1421">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-1421">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-1422">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1422">Example</span></span>

<span data-ttu-id="92251-1423">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="92251-1423">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="92251-1424">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92251-1424">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="92251-1425">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="92251-1425">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="92251-1426">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="92251-1426">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1427">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1427">Parameters:</span></span>

| <span data-ttu-id="92251-1428">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1428">Name</span></span> | <span data-ttu-id="92251-1429">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1429">Type</span></span> | <span data-ttu-id="92251-1430">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1430">Attributes</span></span> | <span data-ttu-id="92251-1431">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1431">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="92251-1432">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="92251-1432">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="92251-1433">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="92251-1433">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="92251-1434">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1434">Object</span></span> | <span data-ttu-id="92251-1435">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1435">&lt;optional&gt;</span></span> | <span data-ttu-id="92251-1436">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-1436">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="92251-1437">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1437">Object</span></span> | <span data-ttu-id="92251-1438">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1438">&lt;optional&gt;</span></span> | <span data-ttu-id="92251-1439">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1439">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="92251-1440">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-1440">function</span></span>| <span data-ttu-id="92251-1441">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1441">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1442">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1442">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1443">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1443">Requirements</span></span>

|<span data-ttu-id="92251-1444">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1444">Requirement</span></span>| <span data-ttu-id="92251-1445">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1445">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1446">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92251-1447">1.7</span><span class="sxs-lookup"><span data-stu-id="92251-1447">1.7</span></span> |
|[<span data-ttu-id="92251-1448">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92251-1449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92251-1449">ReadItem</span></span> |
|[<span data-ttu-id="92251-1450">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92251-1451">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92251-1451">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="92251-1452">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="92251-1452">saveAsync([options], callback)</span></span>

<span data-ttu-id="92251-1453">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="92251-1453">Asynchronously saves an item.</span></span>

<span data-ttu-id="92251-p181">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="92251-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1457">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="92251-1457">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="92251-1458">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="92251-1458">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="92251-p183">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="92251-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="92251-1462">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="92251-1462">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="92251-1463">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="92251-1463">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="92251-1464">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="92251-1464">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="92251-1465">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="92251-1465">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1466">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1466">Parameters:</span></span>

|<span data-ttu-id="92251-1467">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1467">Name</span></span>|<span data-ttu-id="92251-1468">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1468">Type</span></span>|<span data-ttu-id="92251-1469">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1469">Attributes</span></span>|<span data-ttu-id="92251-1470">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1470">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="92251-1471">Object</span><span class="sxs-lookup"><span data-stu-id="92251-1471">Object</span></span>|<span data-ttu-id="92251-1472">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1472">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1473">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-1473">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-1474">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1474">Object</span></span>|<span data-ttu-id="92251-1475">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1475">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1476">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1476">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="92251-1477">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-1477">function</span></span>||<span data-ttu-id="92251-1478">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1478">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="92251-1479">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="92251-1479">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1480">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1480">Requirements</span></span>

|<span data-ttu-id="92251-1481">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1481">Requirement</span></span>|<span data-ttu-id="92251-1482">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1482">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1483">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1484">1.3</span><span class="sxs-lookup"><span data-stu-id="92251-1484">1.3</span></span>|
|[<span data-ttu-id="92251-1485">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1485">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1486">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92251-1486">ReadWriteItem</span></span>|
|[<span data-ttu-id="92251-1487">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1487">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1488">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-1488">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="92251-1489">範例</span><span class="sxs-lookup"><span data-stu-id="92251-1489">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="92251-p185">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="92251-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="92251-1492">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="92251-1492">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="92251-1493">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="92251-1493">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="92251-p186">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="92251-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92251-1497">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="92251-1497">Parameters:</span></span>

|<span data-ttu-id="92251-1498">Nom</span><span class="sxs-lookup"><span data-stu-id="92251-1498">Name</span></span>|<span data-ttu-id="92251-1499">Type</span><span class="sxs-lookup"><span data-stu-id="92251-1499">Type</span></span>|<span data-ttu-id="92251-1500">Attributs</span><span class="sxs-lookup"><span data-stu-id="92251-1500">Attributes</span></span>|<span data-ttu-id="92251-1501">Description</span><span class="sxs-lookup"><span data-stu-id="92251-1501">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="92251-1502">String</span><span class="sxs-lookup"><span data-stu-id="92251-1502">String</span></span>||<span data-ttu-id="92251-p187">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="92251-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="92251-1506">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1506">Object</span></span>|<span data-ttu-id="92251-1507">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1507">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1508">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92251-1508">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="92251-1509">Objet</span><span class="sxs-lookup"><span data-stu-id="92251-1509">Object</span></span>|<span data-ttu-id="92251-1510">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1510">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-1511">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="92251-1511">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="92251-1512">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="92251-1512">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="92251-1513">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92251-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="92251-p188">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="92251-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="92251-p189">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="92251-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="92251-1518">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="92251-1518">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="92251-1519">fonction</span><span class="sxs-lookup"><span data-stu-id="92251-1519">function</span></span>||<span data-ttu-id="92251-1520">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="92251-1520">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92251-1521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92251-1521">Requirements</span></span>

|<span data-ttu-id="92251-1522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92251-1522">Requirement</span></span>|<span data-ttu-id="92251-1523">Valeur</span><span class="sxs-lookup"><span data-stu-id="92251-1523">Value</span></span>|
|---|---|
|[<span data-ttu-id="92251-1524">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92251-1524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="92251-1525">1.2</span><span class="sxs-lookup"><span data-stu-id="92251-1525">1.2</span></span>|
|[<span data-ttu-id="92251-1526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="92251-1526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="92251-1527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92251-1527">ReadWriteItem</span></span>|
|[<span data-ttu-id="92251-1528">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92251-1528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="92251-1529">Composition</span><span class="sxs-lookup"><span data-stu-id="92251-1529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92251-1530">Exemple</span><span class="sxs-lookup"><span data-stu-id="92251-1530">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
