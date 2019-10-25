---
title: Office. Context. Mailbox. Item-Preview ensemble de conditions requises
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: 7a72e6fbbec6dbf9cee07d85237baef93ca7ecd8
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682661"
---
# <a name="item"></a><span data-ttu-id="a658c-102">élément</span><span class="sxs-lookup"><span data-stu-id="a658c-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="a658c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="a658c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="a658c-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a658c-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-106">Requirements</span></span>

|<span data-ttu-id="a658c-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-107">Requirement</span></span>|<span data-ttu-id="a658c-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-110">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-110">1.0</span></span>|
|[<span data-ttu-id="a658c-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="a658c-112">Restricted</span></span>|
|[<span data-ttu-id="a658c-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a658c-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="a658c-115">Members and methods</span></span>

| <span data-ttu-id="a658c-116">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-116">Member</span></span> | <span data-ttu-id="a658c-117">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a658c-118">attachments</span><span class="sxs-lookup"><span data-stu-id="a658c-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="a658c-119">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-119">Member</span></span> |
| [<span data-ttu-id="a658c-120">bcc</span><span class="sxs-lookup"><span data-stu-id="a658c-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="a658c-121">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-121">Member</span></span> |
| [<span data-ttu-id="a658c-122">body</span><span class="sxs-lookup"><span data-stu-id="a658c-122">body</span></span>](#body-body) | <span data-ttu-id="a658c-123">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-123">Member</span></span> |
| [<span data-ttu-id="a658c-124">categories</span><span class="sxs-lookup"><span data-stu-id="a658c-124">categories</span></span>](#categories-categories) | <span data-ttu-id="a658c-125">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-125">Member</span></span> |
| [<span data-ttu-id="a658c-126">cc</span><span class="sxs-lookup"><span data-stu-id="a658c-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a658c-127">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-127">Member</span></span> |
| [<span data-ttu-id="a658c-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="a658c-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a658c-129">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-129">Member</span></span> |
| [<span data-ttu-id="a658c-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a658c-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a658c-131">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-131">Member</span></span> |
| [<span data-ttu-id="a658c-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a658c-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a658c-133">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-133">Member</span></span> |
| [<span data-ttu-id="a658c-134">end</span><span class="sxs-lookup"><span data-stu-id="a658c-134">end</span></span>](#end-datetime) | <span data-ttu-id="a658c-135">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-135">Member</span></span> |
| [<span data-ttu-id="a658c-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="a658c-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="a658c-137">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-137">Member</span></span> |
| [<span data-ttu-id="a658c-138">from</span><span class="sxs-lookup"><span data-stu-id="a658c-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="a658c-139">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-139">Member</span></span> |
| [<span data-ttu-id="a658c-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="a658c-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="a658c-141">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-141">Member</span></span> |
| [<span data-ttu-id="a658c-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a658c-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a658c-143">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-143">Member</span></span> |
| [<span data-ttu-id="a658c-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="a658c-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a658c-145">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-145">Member</span></span> |
| [<span data-ttu-id="a658c-146">itemId</span><span class="sxs-lookup"><span data-stu-id="a658c-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a658c-147">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-147">Member</span></span> |
| [<span data-ttu-id="a658c-148">itemType</span><span class="sxs-lookup"><span data-stu-id="a658c-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="a658c-149">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-149">Member</span></span> |
| [<span data-ttu-id="a658c-150">location</span><span class="sxs-lookup"><span data-stu-id="a658c-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="a658c-151">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-151">Member</span></span> |
| [<span data-ttu-id="a658c-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a658c-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a658c-153">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-153">Member</span></span> |
| [<span data-ttu-id="a658c-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="a658c-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="a658c-155">Member</span><span class="sxs-lookup"><span data-stu-id="a658c-155">Member</span></span> |
| [<span data-ttu-id="a658c-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a658c-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a658c-157">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-157">Member</span></span> |
| [<span data-ttu-id="a658c-158">organizer</span><span class="sxs-lookup"><span data-stu-id="a658c-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="a658c-159">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-159">Member</span></span> |
| [<span data-ttu-id="a658c-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="a658c-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="a658c-161">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-161">Member</span></span> |
| [<span data-ttu-id="a658c-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a658c-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a658c-163">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-163">Member</span></span> |
| [<span data-ttu-id="a658c-164">sender</span><span class="sxs-lookup"><span data-stu-id="a658c-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="a658c-165">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-165">Member</span></span> |
| [<span data-ttu-id="a658c-166">seriesId</span><span class="sxs-lookup"><span data-stu-id="a658c-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="a658c-167">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-167">Member</span></span> |
| [<span data-ttu-id="a658c-168">start</span><span class="sxs-lookup"><span data-stu-id="a658c-168">start</span></span>](#start-datetime) | <span data-ttu-id="a658c-169">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-169">Member</span></span> |
| [<span data-ttu-id="a658c-170">subject</span><span class="sxs-lookup"><span data-stu-id="a658c-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="a658c-171">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-171">Member</span></span> |
| [<span data-ttu-id="a658c-172">to</span><span class="sxs-lookup"><span data-stu-id="a658c-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a658c-173">Membre</span><span class="sxs-lookup"><span data-stu-id="a658c-173">Member</span></span> |
| [<span data-ttu-id="a658c-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a658c-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-175">Method</span></span> |
| [<span data-ttu-id="a658c-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="a658c-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="a658c-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-177">Method</span></span> |
| [<span data-ttu-id="a658c-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a658c-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-179">Method</span></span> |
| [<span data-ttu-id="a658c-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a658c-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-181">Method</span></span> |
| [<span data-ttu-id="a658c-182">close</span><span class="sxs-lookup"><span data-stu-id="a658c-182">close</span></span>](#close) | <span data-ttu-id="a658c-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-183">Method</span></span> |
| [<span data-ttu-id="a658c-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a658c-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="a658c-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-185">Method</span></span> |
| [<span data-ttu-id="a658c-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a658c-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="a658c-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-187">Method</span></span> |
| [<span data-ttu-id="a658c-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="a658c-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-189">Method</span></span> |
| [<span data-ttu-id="a658c-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="a658c-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-191">Method</span></span> |
| [<span data-ttu-id="a658c-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="a658c-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-193">Method</span></span> |
| [<span data-ttu-id="a658c-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="a658c-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="a658c-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-195">Method</span></span> |
| [<span data-ttu-id="a658c-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a658c-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="a658c-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-197">Method</span></span> |
| [<span data-ttu-id="a658c-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a658c-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="a658c-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-199">Method</span></span> |
| [<span data-ttu-id="a658c-200">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-200">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="a658c-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-201">Method</span></span> |
| [<span data-ttu-id="a658c-202">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-202">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="a658c-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-203">Method</span></span> |
| [<span data-ttu-id="a658c-204">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a658c-204">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a658c-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-205">Method</span></span> |
| [<span data-ttu-id="a658c-206">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a658c-206">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a658c-207">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-207">Method</span></span> |
| [<span data-ttu-id="a658c-208">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-208">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a658c-209">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-209">Method</span></span> |
| [<span data-ttu-id="a658c-210">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="a658c-210">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="a658c-211">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-211">Method</span></span> |
| [<span data-ttu-id="a658c-212">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a658c-212">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="a658c-213">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-213">Method</span></span> |
| [<span data-ttu-id="a658c-214">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-214">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="a658c-215">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-215">Method</span></span> |
| [<span data-ttu-id="a658c-216">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-216">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a658c-217">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-217">Method</span></span> |
| [<span data-ttu-id="a658c-218">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-218">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a658c-219">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-219">Method</span></span> |
| [<span data-ttu-id="a658c-220">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-220">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="a658c-221">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-221">Method</span></span> |
| [<span data-ttu-id="a658c-222">saveAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-222">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="a658c-223">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-223">Method</span></span> |
| [<span data-ttu-id="a658c-224">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a658c-224">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a658c-225">Méthode</span><span class="sxs-lookup"><span data-stu-id="a658c-225">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a658c-226">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-226">Example</span></span>

<span data-ttu-id="a658c-227">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="a658c-227">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="a658c-228">Members</span><span class="sxs-lookup"><span data-stu-id="a658c-228">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="a658c-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a658c-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="a658c-230">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="a658c-230">Gets the item's attachments as an array.</span></span> <span data-ttu-id="a658c-231">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-231">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-232">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="a658c-232">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a658c-233">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a658c-233">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-234">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-234">Type</span></span>

*   <span data-ttu-id="a658c-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a658c-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-236">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-236">Requirements</span></span>

|<span data-ttu-id="a658c-237">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-237">Requirement</span></span>|<span data-ttu-id="a658c-238">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-239">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-240">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-240">1.0</span></span>|
|[<span data-ttu-id="a658c-241">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-242">ReadItem</span></span>|
|[<span data-ttu-id="a658c-243">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-244">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-244">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-245">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-245">Example</span></span>

<span data-ttu-id="a658c-246">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a658c-246">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a658c-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a658c-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a658c-248">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="a658c-248">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a658c-249">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-249">Compose mode only.</span></span>

<span data-ttu-id="a658c-250">Par défaut, la collection est limitée à un maximum de 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a658c-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a658c-251">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="a658c-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a658c-252">Obtenir 500 membres maximum.</span><span class="sxs-lookup"><span data-stu-id="a658c-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="a658c-253">Définissez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="a658c-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-254">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-254">Type</span></span>

*   [<span data-ttu-id="a658c-255">Destinataires</span><span class="sxs-lookup"><span data-stu-id="a658c-255">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a658c-256">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-256">Requirements</span></span>

|<span data-ttu-id="a658c-257">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-257">Requirement</span></span>|<span data-ttu-id="a658c-258">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-259">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-260">1.1</span><span class="sxs-lookup"><span data-stu-id="a658c-260">1.1</span></span>|
|[<span data-ttu-id="a658c-261">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-262">ReadItem</span></span>|
|[<span data-ttu-id="a658c-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-264">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-264">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-265">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-265">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="a658c-266">body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="a658c-266">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="a658c-267">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-267">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-268">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-268">Type</span></span>

*   [<span data-ttu-id="a658c-269">Body</span><span class="sxs-lookup"><span data-stu-id="a658c-269">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="a658c-270">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-270">Requirements</span></span>

|<span data-ttu-id="a658c-271">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-271">Requirement</span></span>|<span data-ttu-id="a658c-272">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-273">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-274">1.1</span><span class="sxs-lookup"><span data-stu-id="a658c-274">1.1</span></span>|
|[<span data-ttu-id="a658c-275">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-276">ReadItem</span></span>|
|[<span data-ttu-id="a658c-277">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-278">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-279">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-279">Example</span></span>

<span data-ttu-id="a658c-280">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="a658c-280">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="a658c-281">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-281">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="a658c-282">Catégories : [catégories](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="a658c-282">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="a658c-283">Obtient un objet qui fournit des méthodes pour la gestion des catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-283">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-284">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-284">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-285">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-285">Type</span></span>

*   [<span data-ttu-id="a658c-286">Categories</span><span class="sxs-lookup"><span data-stu-id="a658c-286">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="a658c-287">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-287">Requirements</span></span>

|<span data-ttu-id="a658c-288">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-288">Requirement</span></span>|<span data-ttu-id="a658c-289">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-290">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-291">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-291">Preview</span></span>|
|[<span data-ttu-id="a658c-292">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-293">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-293">ReadItem</span></span>|
|[<span data-ttu-id="a658c-294">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-295">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-295">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-296">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-296">Example</span></span>

<span data-ttu-id="a658c-297">Cet exemple obtient les catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-297">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a658c-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a658c-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a658c-299">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="a658c-299">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a658c-300">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a658c-300">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-301">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-301">Read mode</span></span>

<span data-ttu-id="a658c-302">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="a658c-302">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="a658c-303">Par défaut, la collection est limitée à un maximum de 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a658c-303">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a658c-304">Toutefois, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="a658c-304">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-305">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-305">Compose mode</span></span>

<span data-ttu-id="a658c-306">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="a658c-306">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="a658c-307">Par défaut, la collection est limitée à un maximum de 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a658c-307">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a658c-308">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="a658c-308">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a658c-309">Obtenir 500 membres maximum.</span><span class="sxs-lookup"><span data-stu-id="a658c-309">Get 500 members maximum.</span></span>
- <span data-ttu-id="a658c-310">Définissez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="a658c-310">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a658c-311">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-311">Type</span></span>

*   <span data-ttu-id="a658c-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a658c-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-313">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-313">Requirements</span></span>

|<span data-ttu-id="a658c-314">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-314">Requirement</span></span>|<span data-ttu-id="a658c-315">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-316">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-317">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-317">1.0</span></span>|
|[<span data-ttu-id="a658c-318">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-319">ReadItem</span></span>|
|[<span data-ttu-id="a658c-320">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-321">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-321">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="a658c-322">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="a658c-322">(nullable) conversationId: String</span></span>

<span data-ttu-id="a658c-323">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="a658c-323">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a658c-p109">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="a658c-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a658c-p110">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="a658c-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-328">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-328">Type</span></span>

*   <span data-ttu-id="a658c-329">String</span><span class="sxs-lookup"><span data-stu-id="a658c-329">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-330">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-330">Requirements</span></span>

|<span data-ttu-id="a658c-331">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-331">Requirement</span></span>|<span data-ttu-id="a658c-332">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-333">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-334">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-334">1.0</span></span>|
|[<span data-ttu-id="a658c-335">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-336">ReadItem</span></span>|
|[<span data-ttu-id="a658c-337">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-338">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-338">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-339">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-339">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="a658c-340">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="a658c-340">dateTimeCreated: Date</span></span>

<span data-ttu-id="a658c-p111">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-343">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-343">Type</span></span>

*   <span data-ttu-id="a658c-344">Date</span><span class="sxs-lookup"><span data-stu-id="a658c-344">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-345">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-345">Requirements</span></span>

|<span data-ttu-id="a658c-346">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-346">Requirement</span></span>|<span data-ttu-id="a658c-347">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-348">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-349">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-349">1.0</span></span>|
|[<span data-ttu-id="a658c-350">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-350">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-351">ReadItem</span></span>|
|[<span data-ttu-id="a658c-352">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-352">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-353">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-353">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-354">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-354">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="a658c-355">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="a658c-355">dateTimeModified: Date</span></span>

<span data-ttu-id="a658c-p112">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-358">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-358">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-359">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-359">Type</span></span>

*   <span data-ttu-id="a658c-360">Date</span><span class="sxs-lookup"><span data-stu-id="a658c-360">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-361">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-361">Requirements</span></span>

|<span data-ttu-id="a658c-362">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-362">Requirement</span></span>|<span data-ttu-id="a658c-363">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-363">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-364">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-364">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-365">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-365">1.0</span></span>|
|[<span data-ttu-id="a658c-366">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-366">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-367">ReadItem</span></span>|
|[<span data-ttu-id="a658c-368">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-368">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-369">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-369">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-370">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-370">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a658c-371">end: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a658c-371">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a658c-372">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-372">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a658c-p113">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="a658c-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-375">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-375">Read mode</span></span>

<span data-ttu-id="a658c-376">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="a658c-376">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-377">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-377">Compose mode</span></span>

<span data-ttu-id="a658c-378">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a658c-378">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a658c-379">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="a658c-379">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="a658c-380">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a658c-380">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="a658c-381">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-381">Type</span></span>

*   <span data-ttu-id="a658c-382">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a658c-382">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-383">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-383">Requirements</span></span>

|<span data-ttu-id="a658c-384">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-384">Requirement</span></span>|<span data-ttu-id="a658c-385">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-386">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-387">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-387">1.0</span></span>|
|[<span data-ttu-id="a658c-388">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-389">ReadItem</span></span>|
|[<span data-ttu-id="a658c-390">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-391">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-391">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="a658c-392">enhancedLocation : [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="a658c-392">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="a658c-393">Obtient ou définit les emplacements d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-393">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-394">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-394">Read mode</span></span>

<span data-ttu-id="a658c-395">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui vous permet d’obtenir l’ensemble des emplacements (chacun représenté par un objet [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associé au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a658c-396">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-396">Compose mode</span></span>

<span data-ttu-id="a658c-397">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui fournit des méthodes pour obtenir, supprimer ou ajouter des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-397">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-398">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-398">Type</span></span>

*   [<span data-ttu-id="a658c-399">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="a658c-399">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="a658c-400">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-400">Requirements</span></span>

|<span data-ttu-id="a658c-401">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-401">Requirement</span></span>|<span data-ttu-id="a658c-402">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-403">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-404">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-404">Preview</span></span>|
|[<span data-ttu-id="a658c-405">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-405">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-406">ReadItem</span></span>|
|[<span data-ttu-id="a658c-407">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-407">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-408">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-408">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-409">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-409">Example</span></span>

<span data-ttu-id="a658c-410">L’exemple suivant obtient les emplacements actuels associés au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-410">The following example gets the current locations associated with the appointment.</span></span>

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="a658c-411">from : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a658c-411">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="a658c-412">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="a658c-412">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="a658c-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="a658c-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-415">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a658c-415">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-416">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-416">Read mode</span></span>

<span data-ttu-id="a658c-417">La `from` propriété renvoie un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="a658c-417">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-418">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-418">Compose mode</span></span>

<span data-ttu-id="a658c-419">La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.</span><span class="sxs-lookup"><span data-stu-id="a658c-419">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a658c-420">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-420">Type</span></span>

*   <span data-ttu-id="a658c-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [à partir de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a658c-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-422">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-422">Requirements</span></span>

|<span data-ttu-id="a658c-423">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-423">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a658c-424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-425">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-425">1.0</span></span>|<span data-ttu-id="a658c-426">1.7</span><span class="sxs-lookup"><span data-stu-id="a658c-426">1.7</span></span>|
|[<span data-ttu-id="a658c-427">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-427">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-428">ReadItem</span></span>|<span data-ttu-id="a658c-429">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a658c-429">ReadWriteItem</span></span>|
|[<span data-ttu-id="a658c-430">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-431">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-431">Read</span></span>|<span data-ttu-id="a658c-432">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-432">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="a658c-433">internetHeaders : [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="a658c-433">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="a658c-434">Obtient ou définit les en-têtes Internet personnalisés d’un message.</span><span class="sxs-lookup"><span data-stu-id="a658c-434">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="a658c-435">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-435">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-436">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-436">Type</span></span>

*   [<span data-ttu-id="a658c-437">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="a658c-437">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="a658c-438">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-438">Requirements</span></span>

|<span data-ttu-id="a658c-439">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-439">Requirement</span></span>|<span data-ttu-id="a658c-440">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-441">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-442">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-442">Preview</span></span>|
|[<span data-ttu-id="a658c-443">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-444">ReadItem</span></span>|
|[<span data-ttu-id="a658c-445">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-446">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-446">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-447">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-447">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="a658c-448">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="a658c-448">internetMessageId: String</span></span>

<span data-ttu-id="a658c-p116">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-451">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-451">Type</span></span>

*   <span data-ttu-id="a658c-452">String</span><span class="sxs-lookup"><span data-stu-id="a658c-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-453">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-453">Requirements</span></span>

|<span data-ttu-id="a658c-454">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-454">Requirement</span></span>|<span data-ttu-id="a658c-455">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-456">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-457">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-457">1.0</span></span>|
|[<span data-ttu-id="a658c-458">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-459">ReadItem</span></span>|
|[<span data-ttu-id="a658c-460">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-461">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-462">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-462">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="a658c-463">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="a658c-463">itemClass: String</span></span>

<span data-ttu-id="a658c-p117">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a658c-p118">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="a658c-468">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-468">Type</span></span>|<span data-ttu-id="a658c-469">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-469">Description</span></span>|<span data-ttu-id="a658c-470">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="a658c-470">item class</span></span>|
|---|---|---|
|<span data-ttu-id="a658c-471">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="a658c-471">Appointment items</span></span>|<span data-ttu-id="a658c-472">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="a658c-472">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="a658c-473">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="a658c-473">Message items</span></span>|<span data-ttu-id="a658c-474">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="a658c-474">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="a658c-475">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="a658c-475">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-476">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-476">Type</span></span>

*   <span data-ttu-id="a658c-477">String</span><span class="sxs-lookup"><span data-stu-id="a658c-477">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-478">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-478">Requirements</span></span>

|<span data-ttu-id="a658c-479">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-479">Requirement</span></span>|<span data-ttu-id="a658c-480">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-481">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-482">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-482">1.0</span></span>|
|[<span data-ttu-id="a658c-483">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-484">ReadItem</span></span>|
|[<span data-ttu-id="a658c-485">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-486">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-487">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-487">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a658c-488">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="a658c-488">(nullable) itemId: String</span></span>

<span data-ttu-id="a658c-p119">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-491">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="a658c-491">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a658c-492">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="a658c-492">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a658c-493">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a658c-493">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a658c-494">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a658c-494">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="a658c-p121">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-497">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-497">Type</span></span>

*   <span data-ttu-id="a658c-498">String</span><span class="sxs-lookup"><span data-stu-id="a658c-498">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-499">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-499">Requirements</span></span>

|<span data-ttu-id="a658c-500">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-500">Requirement</span></span>|<span data-ttu-id="a658c-501">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-502">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-503">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-503">1.0</span></span>|
|[<span data-ttu-id="a658c-504">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-504">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-505">ReadItem</span></span>|
|[<span data-ttu-id="a658c-506">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-506">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-507">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-508">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-508">Example</span></span>

<span data-ttu-id="a658c-p122">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a658c-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="a658c-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a658c-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a658c-512">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="a658c-512">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a658c-513">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-513">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-514">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-514">Type</span></span>

*   [<span data-ttu-id="a658c-515">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a658c-515">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a658c-516">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-516">Requirements</span></span>

|<span data-ttu-id="a658c-517">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-517">Requirement</span></span>|<span data-ttu-id="a658c-518">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-519">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-520">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-520">1.0</span></span>|
|[<span data-ttu-id="a658c-521">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-522">ReadItem</span></span>|
|[<span data-ttu-id="a658c-523">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-524">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-524">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-525">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-525">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="a658c-526">location: String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a658c-526">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="a658c-527">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-527">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-528">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-528">Read mode</span></span>

<span data-ttu-id="a658c-529">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-529">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-530">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-530">Compose mode</span></span>

<span data-ttu-id="a658c-531">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-531">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a658c-532">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-532">Type</span></span>

*   <span data-ttu-id="a658c-533">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a658c-533">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-534">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-534">Requirements</span></span>

|<span data-ttu-id="a658c-535">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-535">Requirement</span></span>|<span data-ttu-id="a658c-536">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-537">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-538">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-538">1.0</span></span>|
|[<span data-ttu-id="a658c-539">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-540">ReadItem</span></span>|
|[<span data-ttu-id="a658c-541">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-542">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-542">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a658c-543">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="a658c-543">normalizedSubject: String</span></span>

<span data-ttu-id="a658c-p123">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a658c-p124">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="a658c-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-548">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-548">Type</span></span>

*   <span data-ttu-id="a658c-549">String</span><span class="sxs-lookup"><span data-stu-id="a658c-549">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-550">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-550">Requirements</span></span>

|<span data-ttu-id="a658c-551">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-551">Requirement</span></span>|<span data-ttu-id="a658c-552">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-553">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-554">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-554">1.0</span></span>|
|[<span data-ttu-id="a658c-555">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-556">ReadItem</span></span>|
|[<span data-ttu-id="a658c-557">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-558">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-558">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-559">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-559">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="a658c-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="a658c-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="a658c-561">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-561">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-562">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-562">Type</span></span>

*   [<span data-ttu-id="a658c-563">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="a658c-563">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="a658c-564">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-564">Requirements</span></span>

|<span data-ttu-id="a658c-565">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-565">Requirement</span></span>|<span data-ttu-id="a658c-566">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-567">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-568">1.3</span><span class="sxs-lookup"><span data-stu-id="a658c-568">1.3</span></span>|
|[<span data-ttu-id="a658c-569">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-570">ReadItem</span></span>|
|[<span data-ttu-id="a658c-571">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-572">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-572">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-573">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-573">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a658c-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a658c-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a658c-575">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="a658c-575">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a658c-576">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a658c-576">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-577">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-577">Read mode</span></span>

<span data-ttu-id="a658c-578">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-578">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="a658c-579">Par défaut, la collection est limitée à un maximum de 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a658c-579">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a658c-580">Toutefois, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="a658c-580">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-581">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-581">Compose mode</span></span>

<span data-ttu-id="a658c-582">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-582">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="a658c-583">Par défaut, la collection est limitée à un maximum de 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a658c-583">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a658c-584">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="a658c-584">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a658c-585">Obtenir 500 membres maximum.</span><span class="sxs-lookup"><span data-stu-id="a658c-585">Get 500 members maximum.</span></span>
- <span data-ttu-id="a658c-586">Définissez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="a658c-586">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a658c-587">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-587">Type</span></span>

*   <span data-ttu-id="a658c-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a658c-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-589">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-589">Requirements</span></span>

|<span data-ttu-id="a658c-590">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-590">Requirement</span></span>|<span data-ttu-id="a658c-591">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-592">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-593">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-593">1.0</span></span>|
|[<span data-ttu-id="a658c-594">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-595">ReadItem</span></span>|
|[<span data-ttu-id="a658c-596">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-597">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-597">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="a658c-598">Organisateur : [](/javascript/api/outlook/office.emailaddressdetails)|[organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a658c-598">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="a658c-599">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="a658c-599">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-600">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-600">Read mode</span></span>

<span data-ttu-id="a658c-601">La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-601">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-602">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-602">Compose mode</span></span>

<span data-ttu-id="a658c-603">La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="a658c-603">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="a658c-604">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-604">Type</span></span>

*   <span data-ttu-id="a658c-605">[](/javascript/api/outlook/office.emailaddressdetails) | [Organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a658c-605">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-606">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-606">Requirements</span></span>

|<span data-ttu-id="a658c-607">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-607">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a658c-608">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-609">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-609">1.0</span></span>|<span data-ttu-id="a658c-610">1.7</span><span class="sxs-lookup"><span data-stu-id="a658c-610">1.7</span></span>|
|[<span data-ttu-id="a658c-611">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-611">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-612">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-612">ReadItem</span></span>|<span data-ttu-id="a658c-613">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a658c-613">ReadWriteItem</span></span>|
|[<span data-ttu-id="a658c-614">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-615">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-615">Read</span></span>|<span data-ttu-id="a658c-616">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-616">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="a658c-617">(Nullable) récurrence : [périodicité](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="a658c-617">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="a658c-618">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-618">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="a658c-619">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-619">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="a658c-620">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-620">Read and compose modes for appointment items.</span></span> <span data-ttu-id="a658c-621">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-621">Read mode for meeting request items.</span></span>

<span data-ttu-id="a658c-622">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="a658c-622">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="a658c-623">`null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="a658c-623">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="a658c-624">`undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-624">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="a658c-625">Remarque : les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="a658c-625">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="a658c-626">Remarque : si l’objet de périodicité `null`est, cela indique que l’objet est un rendez-vous unique ou une demande de réunion d’un seul rendez-vous et non d’une série.</span><span class="sxs-lookup"><span data-stu-id="a658c-626">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-627">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-627">Read mode</span></span>

<span data-ttu-id="a658c-628">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-628">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="a658c-629">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-629">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-630">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-630">Compose mode</span></span>

<span data-ttu-id="a658c-631">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-631">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="a658c-632">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-632">This is available for appointments.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="a658c-633">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-633">Type</span></span>

* [<span data-ttu-id="a658c-634">Instances</span><span class="sxs-lookup"><span data-stu-id="a658c-634">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="a658c-635">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-635">Requirement</span></span>|<span data-ttu-id="a658c-636">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-637">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-638">1.7</span><span class="sxs-lookup"><span data-stu-id="a658c-638">1.7</span></span>|
|[<span data-ttu-id="a658c-639">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-640">ReadItem</span></span>|
|[<span data-ttu-id="a658c-641">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-642">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a658c-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a658c-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a658c-644">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="a658c-644">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a658c-645">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a658c-645">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-646">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-646">Read mode</span></span>

<span data-ttu-id="a658c-647">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-647">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="a658c-648">Par défaut, la collection est limitée à un maximum de 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a658c-648">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a658c-649">Toutefois, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="a658c-649">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-650">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-650">Compose mode</span></span>

<span data-ttu-id="a658c-651">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-651">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="a658c-652">Par défaut, la collection est limitée à un maximum de 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a658c-652">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a658c-653">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="a658c-653">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a658c-654">Obtenir 500 membres maximum.</span><span class="sxs-lookup"><span data-stu-id="a658c-654">Get 500 members maximum.</span></span>
- <span data-ttu-id="a658c-655">Définissez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="a658c-655">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="a658c-656">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-656">Type</span></span>

*   <span data-ttu-id="a658c-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a658c-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-658">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-658">Requirements</span></span>

|<span data-ttu-id="a658c-659">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-659">Requirement</span></span>|<span data-ttu-id="a658c-660">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-661">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-662">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-662">1.0</span></span>|
|[<span data-ttu-id="a658c-663">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-664">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-664">ReadItem</span></span>|
|[<span data-ttu-id="a658c-665">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-666">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-666">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="a658c-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a658c-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="a658c-p135">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a658c-p136">Les propriétés [`from`](#from-emailaddressdetailsfrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="a658c-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-672">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a658c-672">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-673">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-673">Type</span></span>

*   [<span data-ttu-id="a658c-674">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a658c-674">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a658c-675">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-675">Requirements</span></span>

|<span data-ttu-id="a658c-676">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-676">Requirement</span></span>|<span data-ttu-id="a658c-677">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-677">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-678">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-678">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-679">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-679">1.0</span></span>|
|[<span data-ttu-id="a658c-680">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-680">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-681">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-681">ReadItem</span></span>|
|[<span data-ttu-id="a658c-682">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-682">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-683">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-683">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-684">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-684">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="a658c-685">(Nullable) seriesId : chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-685">(nullable) seriesId: String</span></span>

<span data-ttu-id="a658c-686">Obtient l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="a658c-686">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="a658c-687">Dans Outlook sur le Web et les clients de bureau `seriesId` , le renvoie l’ID des services Web Exchange (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="a658c-687">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="a658c-688">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="a658c-688">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-689">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="a658c-689">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a658c-690">La `seriesId` propriété n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="a658c-690">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="a658c-691">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a658c-691">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a658c-692">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="a658c-692">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="a658c-693">La `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-693">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="a658c-694">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-694">Type</span></span>

* <span data-ttu-id="a658c-695">String</span><span class="sxs-lookup"><span data-stu-id="a658c-695">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-696">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-696">Requirements</span></span>

|<span data-ttu-id="a658c-697">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-697">Requirement</span></span>|<span data-ttu-id="a658c-698">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-699">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-700">1.7</span><span class="sxs-lookup"><span data-stu-id="a658c-700">1.7</span></span>|
|[<span data-ttu-id="a658c-701">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-702">ReadItem</span></span>|
|[<span data-ttu-id="a658c-703">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-704">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-704">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-705">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-705">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a658c-706">start: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a658c-706">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a658c-707">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-707">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a658c-p139">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="a658c-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-710">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-710">Read mode</span></span>

<span data-ttu-id="a658c-711">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="a658c-711">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-712">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-712">Compose mode</span></span>

<span data-ttu-id="a658c-713">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a658c-713">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a658c-714">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="a658c-714">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="a658c-715">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a658c-715">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="a658c-716">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-716">Type</span></span>

*   <span data-ttu-id="a658c-717">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a658c-717">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-718">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-718">Requirements</span></span>

|<span data-ttu-id="a658c-719">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-719">Requirement</span></span>|<span data-ttu-id="a658c-720">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-720">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-721">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-721">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-722">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-722">1.0</span></span>|
|[<span data-ttu-id="a658c-723">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-723">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-724">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-724">ReadItem</span></span>|
|[<span data-ttu-id="a658c-725">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-725">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-726">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-726">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="a658c-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a658c-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="a658c-728">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-728">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a658c-729">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="a658c-729">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-730">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-730">Read mode</span></span>

<span data-ttu-id="a658c-p140">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a658c-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="a658c-733">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="a658c-733">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-734">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-734">Compose mode</span></span>
<span data-ttu-id="a658c-735">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="a658c-735">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="a658c-736">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-736">Type</span></span>

*   <span data-ttu-id="a658c-737">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a658c-737">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-738">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-738">Requirements</span></span>

|<span data-ttu-id="a658c-739">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-739">Requirement</span></span>|<span data-ttu-id="a658c-740">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-741">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-742">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-742">1.0</span></span>|
|[<span data-ttu-id="a658c-743">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-743">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-744">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-744">ReadItem</span></span>|
|[<span data-ttu-id="a658c-745">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-745">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-746">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-746">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a658c-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a658c-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a658c-748">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="a658c-748">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a658c-749">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a658c-749">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a658c-750">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-750">Read mode</span></span>

<span data-ttu-id="a658c-751">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="a658c-751">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="a658c-752">Par défaut, la collection est limitée à un maximum de 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a658c-752">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a658c-753">Toutefois, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="a658c-753">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="a658c-754">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a658c-754">Compose mode</span></span>

<span data-ttu-id="a658c-755">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="a658c-755">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="a658c-756">Par défaut, la collection est limitée à un maximum de 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a658c-756">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a658c-757">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="a658c-757">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a658c-758">Obtenir 500 membres maximum.</span><span class="sxs-lookup"><span data-stu-id="a658c-758">Get 500 members maximum.</span></span>
- <span data-ttu-id="a658c-759">Définissez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="a658c-759">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a658c-760">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-760">Type</span></span>

*   <span data-ttu-id="a658c-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a658c-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-762">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-762">Requirements</span></span>

|<span data-ttu-id="a658c-763">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-763">Requirement</span></span>|<span data-ttu-id="a658c-764">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-765">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-766">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-766">1.0</span></span>|
|[<span data-ttu-id="a658c-767">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-767">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-768">ReadItem</span></span>|
|[<span data-ttu-id="a658c-769">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-769">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-770">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-770">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="a658c-771">Méthodes</span><span class="sxs-lookup"><span data-stu-id="a658c-771">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a658c-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a658c-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a658c-773">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="a658c-773">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a658c-774">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="a658c-774">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a658c-775">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a658c-775">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-776">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-776">Parameters</span></span>
|<span data-ttu-id="a658c-777">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-777">Name</span></span>|<span data-ttu-id="a658c-778">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-778">Type</span></span>|<span data-ttu-id="a658c-779">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-779">Attributes</span></span>|<span data-ttu-id="a658c-780">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-780">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="a658c-781">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-781">String</span></span>||<span data-ttu-id="a658c-p144">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="a658c-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a658c-784">String</span><span class="sxs-lookup"><span data-stu-id="a658c-784">String</span></span>||<span data-ttu-id="a658c-p145">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a658c-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a658c-787">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-787">Object</span></span>|<span data-ttu-id="a658c-788">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-788">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-789">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-789">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-790">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-790">Object</span></span>|<span data-ttu-id="a658c-791">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-791">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-792">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-792">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a658c-793">Boolean</span><span class="sxs-lookup"><span data-stu-id="a658c-793">Boolean</span></span>|<span data-ttu-id="a658c-794">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-794">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-795">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a658c-795">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a658c-796">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-796">function</span></span>|<span data-ttu-id="a658c-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-797">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-798">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-798">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a658c-799">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a658c-799">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a658c-800">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="a658c-800">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a658c-801">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a658c-801">Errors</span></span>

|<span data-ttu-id="a658c-802">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a658c-802">Error code</span></span>|<span data-ttu-id="a658c-803">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-803">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a658c-804">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="a658c-804">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a658c-805">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="a658c-805">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a658c-806">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a658c-806">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-807">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-807">Requirements</span></span>

|<span data-ttu-id="a658c-808">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-808">Requirement</span></span>|<span data-ttu-id="a658c-809">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-809">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-810">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-810">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-811">1.1</span><span class="sxs-lookup"><span data-stu-id="a658c-811">1.1</span></span>|
|[<span data-ttu-id="a658c-812">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-812">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-813">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a658c-813">ReadWriteItem</span></span>|
|[<span data-ttu-id="a658c-814">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-814">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-815">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-815">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a658c-816">Exemples</span><span class="sxs-lookup"><span data-stu-id="a658c-816">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="a658c-817">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="a658c-817">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync(
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
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="a658c-818">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a658c-818">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a658c-819">Ajoute un fichier à partir du codage Base64 à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="a658c-819">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a658c-820">La `addFileAttachmentFromBase64Async` méthode charge le fichier à partir du codage Base64 et l’associe à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="a658c-820">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="a658c-821">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="a658c-821">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="a658c-822">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a658c-822">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-823">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-823">Parameters</span></span>

|<span data-ttu-id="a658c-824">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-824">Name</span></span>|<span data-ttu-id="a658c-825">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-825">Type</span></span>|<span data-ttu-id="a658c-826">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-826">Attributes</span></span>|<span data-ttu-id="a658c-827">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-827">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="a658c-828">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-828">String</span></span>||<span data-ttu-id="a658c-829">Contenu encodé en base64 d’une image ou d’un fichier à ajouter à un message électronique ou à un événement.</span><span class="sxs-lookup"><span data-stu-id="a658c-829">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="a658c-830">String</span><span class="sxs-lookup"><span data-stu-id="a658c-830">String</span></span>||<span data-ttu-id="a658c-p147">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a658c-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a658c-833">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-833">Object</span></span>|<span data-ttu-id="a658c-834">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-834">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-835">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-835">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-836">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-836">Object</span></span>|<span data-ttu-id="a658c-837">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-837">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-838">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-838">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a658c-839">Boolean</span><span class="sxs-lookup"><span data-stu-id="a658c-839">Boolean</span></span>|<span data-ttu-id="a658c-840">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-840">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-841">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a658c-841">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a658c-842">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-842">function</span></span>|<span data-ttu-id="a658c-843">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-843">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-844">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-844">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a658c-845">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a658c-845">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a658c-846">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="a658c-846">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a658c-847">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a658c-847">Errors</span></span>

|<span data-ttu-id="a658c-848">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a658c-848">Error code</span></span>|<span data-ttu-id="a658c-849">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-849">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a658c-850">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="a658c-850">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a658c-851">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="a658c-851">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a658c-852">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a658c-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-853">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-853">Requirements</span></span>

|<span data-ttu-id="a658c-854">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-854">Requirement</span></span>|<span data-ttu-id="a658c-855">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-856">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-857">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-857">Preview</span></span>|
|[<span data-ttu-id="a658c-858">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-858">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a658c-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="a658c-860">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-860">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-861">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-861">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a658c-862">Exemples</span><span class="sxs-lookup"><span data-stu-id="a658c-862">Examples</span></span>

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
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a658c-863">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a658c-863">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a658c-864">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="a658c-864">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="a658c-865">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="a658c-865">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-866">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-866">Parameters</span></span>

| <span data-ttu-id="a658c-867">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-867">Name</span></span> | <span data-ttu-id="a658c-868">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-868">Type</span></span> | <span data-ttu-id="a658c-869">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-869">Attributes</span></span> | <span data-ttu-id="a658c-870">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-870">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a658c-871">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a658c-871">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a658c-872">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="a658c-872">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a658c-873">Fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-873">Function</span></span> || <span data-ttu-id="a658c-p148">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a658c-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a658c-877">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-877">Object</span></span> | <span data-ttu-id="a658c-878">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-878">&lt;optional&gt;</span></span> | <span data-ttu-id="a658c-879">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-879">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a658c-880">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-880">Object</span></span> | <span data-ttu-id="a658c-881">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-881">&lt;optional&gt;</span></span> | <span data-ttu-id="a658c-882">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-882">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a658c-883">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-883">function</span></span>| <span data-ttu-id="a658c-884">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-884">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-885">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-885">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-886">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-886">Requirements</span></span>

|<span data-ttu-id="a658c-887">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-887">Requirement</span></span>| <span data-ttu-id="a658c-888">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-888">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-889">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-889">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a658c-890">1.7</span><span class="sxs-lookup"><span data-stu-id="a658c-890">1.7</span></span> |
|[<span data-ttu-id="a658c-891">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-891">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a658c-892">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-892">ReadItem</span></span> |
|[<span data-ttu-id="a658c-893">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-893">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a658c-894">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-894">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="a658c-895">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-895">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a658c-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a658c-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a658c-897">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-897">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a658c-p149">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a658c-901">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a658c-901">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a658c-902">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="a658c-902">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-903">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-903">Parameters</span></span>

|<span data-ttu-id="a658c-904">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-904">Name</span></span>|<span data-ttu-id="a658c-905">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-905">Type</span></span>|<span data-ttu-id="a658c-906">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-906">Attributes</span></span>|<span data-ttu-id="a658c-907">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-907">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="a658c-908">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-908">String</span></span>||<span data-ttu-id="a658c-p150">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a658c-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a658c-911">String</span><span class="sxs-lookup"><span data-stu-id="a658c-911">String</span></span>||<span data-ttu-id="a658c-912">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="a658c-912">The subject of the item to be attached.</span></span> <span data-ttu-id="a658c-913">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a658c-913">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a658c-914">Object</span><span class="sxs-lookup"><span data-stu-id="a658c-914">Object</span></span>|<span data-ttu-id="a658c-915">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-915">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-916">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-916">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-917">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-917">Object</span></span>|<span data-ttu-id="a658c-918">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-918">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-919">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-919">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-920">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-920">function</span></span>|<span data-ttu-id="a658c-921">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-921">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-922">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a658c-923">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a658c-923">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a658c-924">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="a658c-924">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a658c-925">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a658c-925">Errors</span></span>

|<span data-ttu-id="a658c-926">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a658c-926">Error code</span></span>|<span data-ttu-id="a658c-927">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-927">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a658c-928">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a658c-928">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-929">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-929">Requirements</span></span>

|<span data-ttu-id="a658c-930">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-930">Requirement</span></span>|<span data-ttu-id="a658c-931">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-931">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-932">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-932">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-933">1.1</span><span class="sxs-lookup"><span data-stu-id="a658c-933">1.1</span></span>|
|[<span data-ttu-id="a658c-934">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-934">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-935">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a658c-935">ReadWriteItem</span></span>|
|[<span data-ttu-id="a658c-936">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-936">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-937">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-937">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-938">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-938">Example</span></span>

<span data-ttu-id="a658c-939">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a658c-939">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

<br>

---
---

#### <a name="close"></a><span data-ttu-id="a658c-940">close()</span><span class="sxs-lookup"><span data-stu-id="a658c-940">close()</span></span>

<span data-ttu-id="a658c-941">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="a658c-941">Closes the current item that is being composed.</span></span>

<span data-ttu-id="a658c-p152">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="a658c-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-944">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-944">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="a658c-945">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="a658c-945">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-946">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-946">Requirements</span></span>

|<span data-ttu-id="a658c-947">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-947">Requirement</span></span>|<span data-ttu-id="a658c-948">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-948">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-949">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-949">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-950">1.3</span><span class="sxs-lookup"><span data-stu-id="a658c-950">1.3</span></span>|
|[<span data-ttu-id="a658c-951">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-951">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-952">Restreinte</span><span class="sxs-lookup"><span data-stu-id="a658c-952">Restricted</span></span>|
|[<span data-ttu-id="a658c-953">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-953">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-954">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-954">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="a658c-955">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a658c-955">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="a658c-956">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a658c-956">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-957">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a658c-958">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="a658c-958">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a658c-959">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="a658c-959">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a658c-p153">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="a658c-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-963">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-963">Parameters</span></span>

|<span data-ttu-id="a658c-964">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-964">Name</span></span>|<span data-ttu-id="a658c-965">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-965">Type</span></span>|<span data-ttu-id="a658c-966">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-966">Attributes</span></span>|<span data-ttu-id="a658c-967">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-967">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a658c-968">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a658c-968">String &#124; Object</span></span>||<span data-ttu-id="a658c-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a658c-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a658c-971">**OU**</span><span class="sxs-lookup"><span data-stu-id="a658c-971">**OR**</span></span><br/><span data-ttu-id="a658c-p155">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="a658c-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a658c-974">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-974">String</span></span>|<span data-ttu-id="a658c-975">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-975">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-p156">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a658c-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a658c-978">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-978">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a658c-979">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-979">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-980">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-980">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a658c-981">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-981">String</span></span>||<span data-ttu-id="a658c-p157">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a658c-984">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-984">String</span></span>||<span data-ttu-id="a658c-985">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a658c-985">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a658c-986">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-986">String</span></span>||<span data-ttu-id="a658c-p158">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="a658c-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a658c-989">Booléen</span><span class="sxs-lookup"><span data-stu-id="a658c-989">Boolean</span></span>||<span data-ttu-id="a658c-p159">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a658c-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a658c-992">String</span><span class="sxs-lookup"><span data-stu-id="a658c-992">String</span></span>||<span data-ttu-id="a658c-p160">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a658c-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a658c-996">function</span><span class="sxs-lookup"><span data-stu-id="a658c-996">function</span></span>|<span data-ttu-id="a658c-997">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-997">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-998">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-998">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-999">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-999">Requirements</span></span>

|<span data-ttu-id="a658c-1000">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1000">Requirement</span></span>|<span data-ttu-id="a658c-1001">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1002">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-1003">1.0</span></span>|
|[<span data-ttu-id="a658c-1004">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1005">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1006">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1007">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1007">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a658c-1008">Exemples</span><span class="sxs-lookup"><span data-stu-id="a658c-1008">Examples</span></span>

<span data-ttu-id="a658c-1009">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1009">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a658c-1010">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="a658c-1010">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a658c-1011">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="a658c-1011">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a658c-1012">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="a658c-1012">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a658c-1013">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-1013">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a658c-1014">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1014">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="a658c-1015">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a658c-1015">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="a658c-1016">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a658c-1016">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1017">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-1017">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a658c-1018">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1018">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a658c-1019">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="a658c-1019">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a658c-p161">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="a658c-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1023">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1023">Parameters</span></span>

|<span data-ttu-id="a658c-1024">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1024">Name</span></span>|<span data-ttu-id="a658c-1025">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1025">Type</span></span>|<span data-ttu-id="a658c-1026">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1026">Attributes</span></span>|<span data-ttu-id="a658c-1027">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1027">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a658c-1028">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a658c-1028">String &#124; Object</span></span>||<span data-ttu-id="a658c-p162">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a658c-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a658c-1031">**OU**</span><span class="sxs-lookup"><span data-stu-id="a658c-1031">**OR**</span></span><br/><span data-ttu-id="a658c-p163">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="a658c-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a658c-1034">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-1034">String</span></span>|<span data-ttu-id="a658c-1035">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-p164">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a658c-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a658c-1038">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1038">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a658c-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1040">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-1040">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a658c-1041">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-1041">String</span></span>||<span data-ttu-id="a658c-p165">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a658c-1044">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-1044">String</span></span>||<span data-ttu-id="a658c-1045">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a658c-1045">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a658c-1046">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-1046">String</span></span>||<span data-ttu-id="a658c-p166">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="a658c-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a658c-1049">Booléen</span><span class="sxs-lookup"><span data-stu-id="a658c-1049">Boolean</span></span>||<span data-ttu-id="a658c-p167">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a658c-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a658c-1052">String</span><span class="sxs-lookup"><span data-stu-id="a658c-1052">String</span></span>||<span data-ttu-id="a658c-p168">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a658c-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a658c-1056">function</span><span class="sxs-lookup"><span data-stu-id="a658c-1056">function</span></span>|<span data-ttu-id="a658c-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1058">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1059">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1059">Requirements</span></span>

|<span data-ttu-id="a658c-1060">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1060">Requirement</span></span>|<span data-ttu-id="a658c-1061">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1061">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1062">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1062">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1063">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-1063">1.0</span></span>|
|[<span data-ttu-id="a658c-1064">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1064">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1065">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1065">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1066">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1066">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1067">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1067">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a658c-1068">Exemples</span><span class="sxs-lookup"><span data-stu-id="a658c-1068">Examples</span></span>

<span data-ttu-id="a658c-1069">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1069">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a658c-1070">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="a658c-1070">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a658c-1071">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="a658c-1071">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a658c-1072">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="a658c-1072">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a658c-1073">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-1073">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a658c-1074">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1074">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="a658c-1075">getAllInternetHeadersAsync ([options], [Rappel])</span><span class="sxs-lookup"><span data-stu-id="a658c-1075">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="a658c-1076">Obtient tous les en-têtes Internet pour le message sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="a658c-1076">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="a658c-1077">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-1077">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1078">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1078">Parameters</span></span>

|<span data-ttu-id="a658c-1079">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1079">Name</span></span>|<span data-ttu-id="a658c-1080">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1080">Type</span></span>|<span data-ttu-id="a658c-1081">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1081">Attributes</span></span>|<span data-ttu-id="a658c-1082">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1082">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a658c-1083">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1083">Object</span></span>|<span data-ttu-id="a658c-1084">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1084">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1085">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1085">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1086">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1086">Object</span></span>|<span data-ttu-id="a658c-1087">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1088">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1088">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-1089">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1089">function</span></span>|<span data-ttu-id="a658c-1090">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1091">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1091">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="a658c-1092">En cas de réussite, les données des en-têtes Internet sont fournies dans la propriété asyncResult. Value sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="a658c-1092">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="a658c-1093">Reportez-vous à la [norme RFC 2183](https://tools.ietf.org/html/rfc2183) pour les informations de mise en forme de la valeur de chaîne renvoyée.</span><span class="sxs-lookup"><span data-stu-id="a658c-1093">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="a658c-1094">En cas d’échec de l’appel, la propriété asyncResult. Error contient un code d’erreur correspondant à la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="a658c-1094">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1095">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1095">Requirements</span></span>

|<span data-ttu-id="a658c-1096">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1096">Requirement</span></span>|<span data-ttu-id="a658c-1097">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1098">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1099">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-1099">Preview</span></span>|
|[<span data-ttu-id="a658c-1100">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1101">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1102">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1103">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1103">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1104">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1104">Returns:</span></span>

<span data-ttu-id="a658c-1105">Les données des en-têtes Internet sous forme de chaîne formatée conformément à la [norme RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="a658c-1105">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="a658c-1106">Type : String</span><span class="sxs-lookup"><span data-stu-id="a658c-1106">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a658c-1107">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1107">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="a658c-1108">getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="a658c-1108">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="a658c-1109">Obtient la pièce jointe spécifiée à partir d’un message ou d’un `AttachmentContent` rendez-vous et la renvoie en tant qu’objet.</span><span class="sxs-lookup"><span data-stu-id="a658c-1109">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="a658c-1110">La `getAttachmentContentAsync` méthode obtient la pièce jointe avec l’identificateur spécifié à partir de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-1110">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="a658c-1111">Il est recommandé d’utiliser l’identificateur pour récupérer une pièce jointe dans la même session que l’attachmentIds a été récupérée avec l' `getAttachmentsAsync` appel ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="a658c-1111">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="a658c-1112">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a658c-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="a658c-1113">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="a658c-1113">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1114">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1114">Parameters</span></span>

|<span data-ttu-id="a658c-1115">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1115">Name</span></span>|<span data-ttu-id="a658c-1116">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1116">Type</span></span>|<span data-ttu-id="a658c-1117">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1117">Attributes</span></span>|<span data-ttu-id="a658c-1118">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a658c-1119">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-1119">String</span></span>||<span data-ttu-id="a658c-1120">Identificateur de la pièce jointe que vous souhaitez obtenir.</span><span class="sxs-lookup"><span data-stu-id="a658c-1120">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="a658c-1121">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1121">Object</span></span>|<span data-ttu-id="a658c-1122">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1123">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1124">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1124">Object</span></span>|<span data-ttu-id="a658c-1125">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1126">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-1127">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1127">function</span></span>|<span data-ttu-id="a658c-1128">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1129">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1130">Requirements</span></span>

|<span data-ttu-id="a658c-1131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1131">Requirement</span></span>|<span data-ttu-id="a658c-1132">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1134">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-1134">Preview</span></span>|
|[<span data-ttu-id="a658c-1135">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1136">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1137">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1138">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1138">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1139">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1139">Returns:</span></span>

<span data-ttu-id="a658c-1140">Type : [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="a658c-1140">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="a658c-1141">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1141">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="a658c-1142">getAttachmentsAsync ([options], [Rappel]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a658c-1142">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="a658c-1143">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="a658c-1143">Gets the item's attachments as an array.</span></span> <span data-ttu-id="a658c-1144">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-1144">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1145">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1145">Parameters</span></span>

|<span data-ttu-id="a658c-1146">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1146">Name</span></span>|<span data-ttu-id="a658c-1147">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1147">Type</span></span>|<span data-ttu-id="a658c-1148">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1148">Attributes</span></span>|<span data-ttu-id="a658c-1149">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1149">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a658c-1150">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1150">Object</span></span>|<span data-ttu-id="a658c-1151">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1151">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1152">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1152">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1153">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1153">Object</span></span>|<span data-ttu-id="a658c-1154">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1155">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1155">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-1156">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1156">function</span></span>|<span data-ttu-id="a658c-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1158">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1158">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1159">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1159">Requirements</span></span>

|<span data-ttu-id="a658c-1160">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1160">Requirement</span></span>|<span data-ttu-id="a658c-1161">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1162">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1163">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-1163">Preview</span></span>|
|[<span data-ttu-id="a658c-1164">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1165">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1166">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1167">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-1167">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1168">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1168">Returns:</span></span>

<span data-ttu-id="a658c-1169">Type : Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a658c-1169">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="a658c-1170">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1170">Example</span></span>

<span data-ttu-id="a658c-1171">L’exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a658c-1171">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a658c-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a658c-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a658c-1173">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a658c-1173">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1174">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-1174">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-1175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1175">Requirements</span></span>

|<span data-ttu-id="a658c-1176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1176">Requirement</span></span>|<span data-ttu-id="a658c-1177">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1179">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-1179">1.0</span></span>|
|[<span data-ttu-id="a658c-1180">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1181">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1182">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1183">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1183">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1184">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1184">Returns:</span></span>

<span data-ttu-id="a658c-1185">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a658c-1185">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a658c-1186">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1186">Example</span></span>

<span data-ttu-id="a658c-1187">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a658c-1187">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a658c-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a658c-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a658c-1189">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a658c-1189">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1190">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-1190">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1191">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1191">Parameters</span></span>

|<span data-ttu-id="a658c-1192">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1192">Name</span></span>|<span data-ttu-id="a658c-1193">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1193">Type</span></span>|<span data-ttu-id="a658c-1194">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1194">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="a658c-1195">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a658c-1195">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="a658c-1196">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="a658c-1196">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1197">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1197">Requirements</span></span>

|<span data-ttu-id="a658c-1198">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1198">Requirement</span></span>|<span data-ttu-id="a658c-1199">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1199">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1200">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1201">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-1201">1.0</span></span>|
|[<span data-ttu-id="a658c-1202">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1203">Restreinte</span><span class="sxs-lookup"><span data-stu-id="a658c-1203">Restricted</span></span>|
|[<span data-ttu-id="a658c-1204">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1205">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1205">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1206">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1206">Returns:</span></span>

<span data-ttu-id="a658c-1207">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="a658c-1207">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a658c-1208">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="a658c-1208">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a658c-1209">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1209">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a658c-1210">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="a658c-1210">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="a658c-1211">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="a658c-1211">Value of `entityType`</span></span>|<span data-ttu-id="a658c-1212">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="a658c-1212">Type of objects in returned array</span></span>|<span data-ttu-id="a658c-1213">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="a658c-1213">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="a658c-1214">String</span><span class="sxs-lookup"><span data-stu-id="a658c-1214">String</span></span>|<span data-ttu-id="a658c-1215">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a658c-1215">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="a658c-1216">Contact</span><span class="sxs-lookup"><span data-stu-id="a658c-1216">Contact</span></span>|<span data-ttu-id="a658c-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a658c-1217">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="a658c-1218">String</span><span class="sxs-lookup"><span data-stu-id="a658c-1218">String</span></span>|<span data-ttu-id="a658c-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a658c-1219">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="a658c-1220">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a658c-1220">MeetingSuggestion</span></span>|<span data-ttu-id="a658c-1221">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a658c-1221">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="a658c-1222">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a658c-1222">PhoneNumber</span></span>|<span data-ttu-id="a658c-1223">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a658c-1223">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="a658c-1224">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a658c-1224">TaskSuggestion</span></span>|<span data-ttu-id="a658c-1225">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a658c-1225">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="a658c-1226">String</span><span class="sxs-lookup"><span data-stu-id="a658c-1226">String</span></span>|<span data-ttu-id="a658c-1227">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a658c-1227">**Restricted**</span></span>|

<span data-ttu-id="a658c-1228">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a658c-1228">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="a658c-1229">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1229">Example</span></span>

<span data-ttu-id="a658c-1230">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a658c-1230">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a658c-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a658c-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a658c-1232">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a658c-1232">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1233">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-1233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a658c-1234">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="a658c-1234">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1235">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1235">Parameters</span></span>

|<span data-ttu-id="a658c-1236">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1236">Name</span></span>|<span data-ttu-id="a658c-1237">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1237">Type</span></span>|<span data-ttu-id="a658c-1238">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1238">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a658c-1239">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a658c-1239">String</span></span>|<span data-ttu-id="a658c-1240">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="a658c-1240">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1241">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1241">Requirements</span></span>

|<span data-ttu-id="a658c-1242">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1242">Requirement</span></span>|<span data-ttu-id="a658c-1243">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1243">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1244">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1244">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1245">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-1245">1.0</span></span>|
|[<span data-ttu-id="a658c-1246">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1246">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1247">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1247">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1248">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1248">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1249">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1249">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1250">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1250">Returns:</span></span>

<span data-ttu-id="a658c-p174">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="a658c-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a658c-1253">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a658c-1253">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="a658c-1254">getInitializationContextAsync ([options], [Rappel])</span><span class="sxs-lookup"><span data-stu-id="a658c-1254">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="a658c-1255">Obtient les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="a658c-1255">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1256">Cette méthode est uniquement prise en charge par Outlook 2016 ou une version ultérieure sur Windows (versions « démarrer en un clic » ultérieures à 16.0.8413.1000) et Outlook sur le Web pour Office 365.</span><span class="sxs-lookup"><span data-stu-id="a658c-1256">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1257">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1257">Parameters</span></span>

|<span data-ttu-id="a658c-1258">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1258">Name</span></span>|<span data-ttu-id="a658c-1259">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1259">Type</span></span>|<span data-ttu-id="a658c-1260">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1260">Attributes</span></span>|<span data-ttu-id="a658c-1261">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a658c-1262">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1262">Object</span></span>|<span data-ttu-id="a658c-1263">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1264">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1265">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1265">Object</span></span>|<span data-ttu-id="a658c-1266">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1267">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-1268">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1268">function</span></span>|<span data-ttu-id="a658c-1269">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1269">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1270">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a658c-1271">En cas de réussite, les données d’initialisation sont fournies `asyncResult.value` dans la propriété sous la forme d’une chaîne.</span><span class="sxs-lookup"><span data-stu-id="a658c-1271">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="a658c-1272">S’il n’existe pas de contexte d’initialisation `asyncResult` , l’objet contient `Error` un objet dont `code` la propriété est `9020` définie sur `name` et sa propriété `GenericResponseError`est définie sur.</span><span class="sxs-lookup"><span data-stu-id="a658c-1272">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1273">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1273">Requirements</span></span>

|<span data-ttu-id="a658c-1274">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1274">Requirement</span></span>|<span data-ttu-id="a658c-1275">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1275">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1276">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1277">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-1277">Preview</span></span>|
|[<span data-ttu-id="a658c-1278">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1279">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1280">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1281">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1281">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-1282">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1282">Example</span></span>

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="a658c-1283">getItemIdAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="a658c-1283">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="a658c-1284">Obtient de manière asynchrone l’ID d’un élément enregistré.</span><span class="sxs-lookup"><span data-stu-id="a658c-1284">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="a658c-1285">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="a658c-1285">Compose mode only.</span></span>

<span data-ttu-id="a658c-1286">Lorsqu’elle est appelée, cette méthode renvoie l’ID de l’élément par le biais de la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1286">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1287">Si votre complément appelle `getItemIdAsync` sur un élément en mode composition (par exemple, pour obtenir un à utiliser avec `itemId` EWS ou l’API REST), sachez que lorsque Outlook est en mode mis en cache, l’élément peut prendre un certain temps avant la synchronisation de l’élément avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="a658c-1287">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="a658c-1288">Tant que l’élément n’est pas synchronisé `itemId` , le n’est pas reconnu et son utilisation renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="a658c-1288">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1289">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1289">Parameters</span></span>

|<span data-ttu-id="a658c-1290">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1290">Name</span></span>|<span data-ttu-id="a658c-1291">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1291">Type</span></span>|<span data-ttu-id="a658c-1292">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1292">Attributes</span></span>|<span data-ttu-id="a658c-1293">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1293">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a658c-1294">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1294">Object</span></span>|<span data-ttu-id="a658c-1295">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1295">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1296">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1296">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1297">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1297">Object</span></span>|<span data-ttu-id="a658c-1298">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1299">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1299">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-1300">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1300">function</span></span>||<span data-ttu-id="a658c-1301">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1301">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a658c-1302">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1302">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a658c-1303">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a658c-1303">Errors</span></span>

|<span data-ttu-id="a658c-1304">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a658c-1304">Error code</span></span>|<span data-ttu-id="a658c-1305">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1305">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="a658c-1306">L’ID ne peut pas être récupéré tant que l’élément n’est pas enregistré.</span><span class="sxs-lookup"><span data-stu-id="a658c-1306">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1307">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1307">Requirements</span></span>

|<span data-ttu-id="a658c-1308">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1308">Requirement</span></span>|<span data-ttu-id="a658c-1309">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1309">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1310">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1310">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1311">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-1311">Preview</span></span>|
|[<span data-ttu-id="a658c-1312">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1312">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1313">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1314">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1314">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1315">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-1315">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a658c-1316">Exemples</span><span class="sxs-lookup"><span data-stu-id="a658c-1316">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="a658c-1317">L’exemple suivant montre la structure du `result` paramètre transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1317">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="a658c-1318">La `value` propriété contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-1318">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="a658c-1319">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a658c-1319">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a658c-1320">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a658c-1320">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1321">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-1321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a658c-p178">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="a658c-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a658c-1325">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="a658c-1325">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a658c-1326">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1326">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a658c-p179">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-1330">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1330">Requirements</span></span>

|<span data-ttu-id="a658c-1331">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1331">Requirement</span></span>|<span data-ttu-id="a658c-1332">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1332">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1333">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1334">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-1334">1.0</span></span>|
|[<span data-ttu-id="a658c-1335">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1336">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1337">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1338">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1338">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1339">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1339">Returns:</span></span>

<span data-ttu-id="a658c-p180">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="a658c-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a658c-1342">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a658c-1342">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a658c-1343">Object</span><span class="sxs-lookup"><span data-stu-id="a658c-1343">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a658c-1344">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1344">Example</span></span>

<span data-ttu-id="a658c-1345">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="a658c-1345">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a658c-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="a658c-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a658c-1347">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a658c-1347">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1348">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-1348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a658c-1349">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="a658c-1349">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a658c-p181">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="a658c-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1352">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1352">Parameters</span></span>

|<span data-ttu-id="a658c-1353">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1353">Name</span></span>|<span data-ttu-id="a658c-1354">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1354">Type</span></span>|<span data-ttu-id="a658c-1355">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1355">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a658c-1356">String</span><span class="sxs-lookup"><span data-stu-id="a658c-1356">String</span></span>|<span data-ttu-id="a658c-1357">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="a658c-1357">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1358">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1358">Requirements</span></span>

|<span data-ttu-id="a658c-1359">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1359">Requirement</span></span>|<span data-ttu-id="a658c-1360">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1360">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1361">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1362">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-1362">1.0</span></span>|
|[<span data-ttu-id="a658c-1363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1364">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1365">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1366">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1366">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1367">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1367">Returns:</span></span>

<span data-ttu-id="a658c-1368">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a658c-1368">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="a658c-1369">Type : Array.< String ></span><span class="sxs-lookup"><span data-stu-id="a658c-1369">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="a658c-1370">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1370">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a658c-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a658c-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a658c-1372">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="a658c-1372">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a658c-p182">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="a658c-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1375">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1375">Parameters</span></span>

|<span data-ttu-id="a658c-1376">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1376">Name</span></span>|<span data-ttu-id="a658c-1377">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1377">Type</span></span>|<span data-ttu-id="a658c-1378">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1378">Attributes</span></span>|<span data-ttu-id="a658c-1379">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1379">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="a658c-1380">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a658c-1380">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a658c-p183">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="a658c-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="a658c-1384">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1384">Object</span></span>|<span data-ttu-id="a658c-1385">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1386">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1386">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1387">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1387">Object</span></span>|<span data-ttu-id="a658c-1388">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1388">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1389">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1389">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-1390">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1390">function</span></span>||<span data-ttu-id="a658c-1391">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1391">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a658c-1392">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1392">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a658c-1393">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1393">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1394">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1394">Requirements</span></span>

|<span data-ttu-id="a658c-1395">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1395">Requirement</span></span>|<span data-ttu-id="a658c-1396">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1397">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="a658c-1398">1.2</span></span>|
|[<span data-ttu-id="a658c-1399">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1400">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1401">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1402">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-1402">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1403">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1403">Returns:</span></span>

<span data-ttu-id="a658c-1404">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1404">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="a658c-1405">Type : String</span><span class="sxs-lookup"><span data-stu-id="a658c-1405">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a658c-1406">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1406">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a658c-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a658c-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a658c-1408">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a658c-1408">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="a658c-1409">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a658c-1409">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1410">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-1410">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-1411">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1411">Requirements</span></span>

|<span data-ttu-id="a658c-1412">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1412">Requirement</span></span>|<span data-ttu-id="a658c-1413">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1413">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1414">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1414">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1415">1.6</span><span class="sxs-lookup"><span data-stu-id="a658c-1415">1.6</span></span>|
|[<span data-ttu-id="a658c-1416">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1417">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1418">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1418">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1419">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1419">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1420">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1420">Returns:</span></span>

<span data-ttu-id="a658c-1421">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a658c-1421">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a658c-1422">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1422">Example</span></span>

<span data-ttu-id="a658c-1423">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a658c-1423">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="a658c-1424">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a658c-1424">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="a658c-p186">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a658c-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1427">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a658c-1427">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a658c-p187">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="a658c-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a658c-1431">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="a658c-1431">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a658c-1432">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1432">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a658c-p188">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a658c-1436">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1436">Requirements</span></span>

|<span data-ttu-id="a658c-1437">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1437">Requirement</span></span>|<span data-ttu-id="a658c-1438">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1438">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1439">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1440">1.6</span><span class="sxs-lookup"><span data-stu-id="a658c-1440">1.6</span></span>|
|[<span data-ttu-id="a658c-1441">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1442">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1443">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1444">Lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1444">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a658c-1445">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a658c-1445">Returns:</span></span>

<span data-ttu-id="a658c-p189">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="a658c-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="a658c-1448">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1448">Example</span></span>

<span data-ttu-id="a658c-1449">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="a658c-1449">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="a658c-1450">getSharedPropertiesAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="a658c-1450">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="a658c-1451">Obtient les propriétés du rendez-vous ou du message sélectionné dans un dossier partagé, un calendrier ou une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="a658c-1451">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1452">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1452">Parameters</span></span>

|<span data-ttu-id="a658c-1453">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1453">Name</span></span>|<span data-ttu-id="a658c-1454">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1454">Type</span></span>|<span data-ttu-id="a658c-1455">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1455">Attributes</span></span>|<span data-ttu-id="a658c-1456">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1456">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a658c-1457">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1457">Object</span></span>|<span data-ttu-id="a658c-1458">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1458">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1459">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1459">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1460">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1460">Object</span></span>|<span data-ttu-id="a658c-1461">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1461">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1462">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1462">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-1463">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1463">function</span></span>||<span data-ttu-id="a658c-1464">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1464">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a658c-1465">Les propriétés partagées sont fournies sous [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) la forme d' `asyncResult.value` un objet dans la propriété.</span><span class="sxs-lookup"><span data-stu-id="a658c-1465">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a658c-1466">Cet objet peut être utilisé pour obtenir les propriétés partagées de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-1466">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1467">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1467">Requirements</span></span>

|<span data-ttu-id="a658c-1468">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1468">Requirement</span></span>|<span data-ttu-id="a658c-1469">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1470">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1471">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a658c-1471">Preview</span></span>|
|[<span data-ttu-id="a658c-1472">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1473">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1474">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1475">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-1476">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1476">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a658c-1477">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a658c-1477">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a658c-1478">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a658c-1478">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a658c-p191">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="a658c-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1482">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1482">Parameters</span></span>

|<span data-ttu-id="a658c-1483">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1483">Name</span></span>|<span data-ttu-id="a658c-1484">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1484">Type</span></span>|<span data-ttu-id="a658c-1485">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1485">Attributes</span></span>|<span data-ttu-id="a658c-1486">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1486">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="a658c-1487">function</span><span class="sxs-lookup"><span data-stu-id="a658c-1487">function</span></span>||<span data-ttu-id="a658c-1488">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1488">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a658c-1489">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1489">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a658c-1490">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="a658c-1490">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="a658c-1491">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1491">Object</span></span>|<span data-ttu-id="a658c-1492">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1493">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1493">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a658c-1494">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1494">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1495">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1495">Requirements</span></span>

|<span data-ttu-id="a658c-1496">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1496">Requirement</span></span>|<span data-ttu-id="a658c-1497">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1497">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1498">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1499">1.0</span><span class="sxs-lookup"><span data-stu-id="a658c-1499">1.0</span></span>|
|[<span data-ttu-id="a658c-1500">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1500">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1501">ReadItem</span></span>|
|[<span data-ttu-id="a658c-1502">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1502">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1503">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1503">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-1504">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1504">Example</span></span>

<span data-ttu-id="a658c-p194">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="a658c-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a658c-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a658c-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a658c-1509">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a658c-1509">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a658c-1510">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-1510">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="a658c-1511">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="a658c-1511">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="a658c-1512">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a658c-1512">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="a658c-1513">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="a658c-1513">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1514">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1514">Parameters</span></span>

|<span data-ttu-id="a658c-1515">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1515">Name</span></span>|<span data-ttu-id="a658c-1516">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1516">Type</span></span>|<span data-ttu-id="a658c-1517">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1517">Attributes</span></span>|<span data-ttu-id="a658c-1518">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1518">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a658c-1519">String</span><span class="sxs-lookup"><span data-stu-id="a658c-1519">String</span></span>||<span data-ttu-id="a658c-1520">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="a658c-1520">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="a658c-1521">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1521">Object</span></span>|<span data-ttu-id="a658c-1522">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1522">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1523">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1523">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1524">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1524">Object</span></span>|<span data-ttu-id="a658c-1525">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1525">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1526">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1526">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-1527">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1527">function</span></span>|<span data-ttu-id="a658c-1528">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1528">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1529">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1529">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a658c-1530">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="a658c-1530">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a658c-1531">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a658c-1531">Errors</span></span>

|<span data-ttu-id="a658c-1532">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a658c-1532">Error code</span></span>|<span data-ttu-id="a658c-1533">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1533">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="a658c-1534">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="a658c-1534">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1535">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1535">Requirements</span></span>

|<span data-ttu-id="a658c-1536">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1536">Requirement</span></span>|<span data-ttu-id="a658c-1537">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1537">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1538">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1539">1.1</span><span class="sxs-lookup"><span data-stu-id="a658c-1539">1.1</span></span>|
|[<span data-ttu-id="a658c-1540">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1541">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1541">ReadWriteItem</span></span>|
|[<span data-ttu-id="a658c-1542">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1543">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-1543">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-1544">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1544">Example</span></span>

<span data-ttu-id="a658c-1545">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="a658c-1545">The following code removes an attachment with an identifier of '0'.</span></span>

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

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="a658c-1546">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a658c-1546">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="a658c-1547">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="a658c-1547">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="a658c-1548">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="a658c-1548">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1549">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1549">Parameters</span></span>

| <span data-ttu-id="a658c-1550">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1550">Name</span></span> | <span data-ttu-id="a658c-1551">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1551">Type</span></span> | <span data-ttu-id="a658c-1552">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1552">Attributes</span></span> | <span data-ttu-id="a658c-1553">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1553">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a658c-1554">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a658c-1554">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a658c-1555">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="a658c-1555">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="a658c-1556">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1556">Object</span></span> | <span data-ttu-id="a658c-1557">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1557">&lt;optional&gt;</span></span> | <span data-ttu-id="a658c-1558">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1558">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a658c-1559">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1559">Object</span></span> | <span data-ttu-id="a658c-1560">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1560">&lt;optional&gt;</span></span> | <span data-ttu-id="a658c-1561">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1561">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a658c-1562">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1562">function</span></span>| <span data-ttu-id="a658c-1563">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1563">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1564">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1564">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1565">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1565">Requirements</span></span>

|<span data-ttu-id="a658c-1566">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1566">Requirement</span></span>| <span data-ttu-id="a658c-1567">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1567">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1568">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a658c-1569">1.7</span><span class="sxs-lookup"><span data-stu-id="a658c-1569">1.7</span></span> |
|[<span data-ttu-id="a658c-1570">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1570">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a658c-1571">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1571">ReadItem</span></span> |
|[<span data-ttu-id="a658c-1572">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1572">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a658c-1573">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a658c-1573">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="a658c-1574">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a658c-1574">saveAsync([options], callback)</span></span>

<span data-ttu-id="a658c-1575">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a658c-1575">Asynchronously saves an item.</span></span>

<span data-ttu-id="a658c-1576">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1576">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="a658c-1577">Dans Outlook sur le web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="a658c-1577">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="a658c-1578">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="a658c-1578">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1579">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="a658c-1579">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="a658c-1580">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="a658c-1580">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="a658c-p198">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="a658c-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="a658c-1584">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="a658c-1584">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="a658c-1585">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="a658c-1585">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="a658c-1586">La méthode `saveAsync` échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="a658c-1586">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="a658c-1587">Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="a658c-1587">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="a658c-1588">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="a658c-1588">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1589">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1589">Parameters</span></span>

|<span data-ttu-id="a658c-1590">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1590">Name</span></span>|<span data-ttu-id="a658c-1591">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1591">Type</span></span>|<span data-ttu-id="a658c-1592">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1592">Attributes</span></span>|<span data-ttu-id="a658c-1593">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1593">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a658c-1594">Object</span><span class="sxs-lookup"><span data-stu-id="a658c-1594">Object</span></span>|<span data-ttu-id="a658c-1595">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1595">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1596">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1596">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1597">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1597">Object</span></span>|<span data-ttu-id="a658c-1598">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1599">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1599">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a658c-1600">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1600">function</span></span>||<span data-ttu-id="a658c-1601">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a658c-1602">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a658c-1602">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1603">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1603">Requirements</span></span>

|<span data-ttu-id="a658c-1604">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1604">Requirement</span></span>|<span data-ttu-id="a658c-1605">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1605">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1606">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1607">1.3</span><span class="sxs-lookup"><span data-stu-id="a658c-1607">1.3</span></span>|
|[<span data-ttu-id="a658c-1608">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1609">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1609">ReadWriteItem</span></span>|
|[<span data-ttu-id="a658c-1610">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1611">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-1611">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a658c-1612">範例</span><span class="sxs-lookup"><span data-stu-id="a658c-1612">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="a658c-p200">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a658c-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a658c-1615">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a658c-1615">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a658c-1616">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a658c-1616">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a658c-p201">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="a658c-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a658c-1620">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a658c-1620">Parameters</span></span>

|<span data-ttu-id="a658c-1621">Nom</span><span class="sxs-lookup"><span data-stu-id="a658c-1621">Name</span></span>|<span data-ttu-id="a658c-1622">Type</span><span class="sxs-lookup"><span data-stu-id="a658c-1622">Type</span></span>|<span data-ttu-id="a658c-1623">Attributs</span><span class="sxs-lookup"><span data-stu-id="a658c-1623">Attributes</span></span>|<span data-ttu-id="a658c-1624">Description</span><span class="sxs-lookup"><span data-stu-id="a658c-1624">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="a658c-1625">String</span><span class="sxs-lookup"><span data-stu-id="a658c-1625">String</span></span>||<span data-ttu-id="a658c-p202">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="a658c-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="a658c-1629">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1629">Object</span></span>|<span data-ttu-id="a658c-1630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1630">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1631">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a658c-1631">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a658c-1632">Objet</span><span class="sxs-lookup"><span data-stu-id="a658c-1632">Object</span></span>|<span data-ttu-id="a658c-1633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1633">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1634">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a658c-1634">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="a658c-1635">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a658c-1635">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="a658c-1636">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a658c-1636">&lt;optional&gt;</span></span>|<span data-ttu-id="a658c-1637">Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau.</span><span class="sxs-lookup"><span data-stu-id="a658c-1637">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="a658c-1638">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="a658c-1638">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a658c-1639">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="a658c-1639">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="a658c-1640">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="a658c-1640">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a658c-1641">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="a658c-1641">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="a658c-1642">fonction</span><span class="sxs-lookup"><span data-stu-id="a658c-1642">function</span></span>||<span data-ttu-id="a658c-1643">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a658c-1643">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a658c-1644">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a658c-1644">Requirements</span></span>

|<span data-ttu-id="a658c-1645">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a658c-1645">Requirement</span></span>|<span data-ttu-id="a658c-1646">Valeur</span><span class="sxs-lookup"><span data-stu-id="a658c-1646">Value</span></span>|
|---|---|
|[<span data-ttu-id="a658c-1647">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a658c-1647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a658c-1648">1.2</span><span class="sxs-lookup"><span data-stu-id="a658c-1648">1.2</span></span>|
|[<span data-ttu-id="a658c-1649">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a658c-1649">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a658c-1650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a658c-1650">ReadWriteItem</span></span>|
|[<span data-ttu-id="a658c-1651">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a658c-1651">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a658c-1652">Composition</span><span class="sxs-lookup"><span data-stu-id="a658c-1652">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a658c-1653">Exemple</span><span class="sxs-lookup"><span data-stu-id="a658c-1653">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
