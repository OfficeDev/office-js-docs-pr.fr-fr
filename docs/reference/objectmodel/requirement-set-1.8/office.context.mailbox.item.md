---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,8
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 065ea3c74580555c0df1af7b495127a25493b612
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001571"
---
# <a name="item"></a><span data-ttu-id="0e2fc-102">élément</span><span class="sxs-lookup"><span data-stu-id="0e2fc-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="0e2fc-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="0e2fc-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="0e2fc-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-106">Requirements</span></span>

|<span data-ttu-id="0e2fc-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-107">Requirement</span></span>|<span data-ttu-id="0e2fc-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-110">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-110">1.0</span></span>|
|[<span data-ttu-id="0e2fc-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="0e2fc-112">Restricted</span></span>|
|[<span data-ttu-id="0e2fc-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0e2fc-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="0e2fc-115">Members and methods</span></span>

| <span data-ttu-id="0e2fc-116">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-116">Member</span></span> | <span data-ttu-id="0e2fc-117">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0e2fc-118">attachments</span><span class="sxs-lookup"><span data-stu-id="0e2fc-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="0e2fc-119">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-119">Member</span></span> |
| [<span data-ttu-id="0e2fc-120">bcc</span><span class="sxs-lookup"><span data-stu-id="0e2fc-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="0e2fc-121">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-121">Member</span></span> |
| [<span data-ttu-id="0e2fc-122">body</span><span class="sxs-lookup"><span data-stu-id="0e2fc-122">body</span></span>](#body-body) | <span data-ttu-id="0e2fc-123">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-123">Member</span></span> |
| [<span data-ttu-id="0e2fc-124">catégories</span><span class="sxs-lookup"><span data-stu-id="0e2fc-124">categories</span></span>](#categories-categories) | <span data-ttu-id="0e2fc-125">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-125">Member</span></span> |
| [<span data-ttu-id="0e2fc-126">cc</span><span class="sxs-lookup"><span data-stu-id="0e2fc-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0e2fc-127">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-127">Member</span></span> |
| [<span data-ttu-id="0e2fc-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="0e2fc-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="0e2fc-129">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-129">Member</span></span> |
| [<span data-ttu-id="0e2fc-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="0e2fc-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="0e2fc-131">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-131">Member</span></span> |
| [<span data-ttu-id="0e2fc-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="0e2fc-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="0e2fc-133">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-133">Member</span></span> |
| [<span data-ttu-id="0e2fc-134">end</span><span class="sxs-lookup"><span data-stu-id="0e2fc-134">end</span></span>](#end-datetime) | <span data-ttu-id="0e2fc-135">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-135">Member</span></span> |
| [<span data-ttu-id="0e2fc-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="0e2fc-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="0e2fc-137">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-137">Member</span></span> |
| [<span data-ttu-id="0e2fc-138">from</span><span class="sxs-lookup"><span data-stu-id="0e2fc-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="0e2fc-139">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-139">Member</span></span> |
| [<span data-ttu-id="0e2fc-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="0e2fc-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="0e2fc-141">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-141">Member</span></span> |
| [<span data-ttu-id="0e2fc-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="0e2fc-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="0e2fc-143">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-143">Member</span></span> |
| [<span data-ttu-id="0e2fc-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="0e2fc-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="0e2fc-145">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-145">Member</span></span> |
| [<span data-ttu-id="0e2fc-146">itemId</span><span class="sxs-lookup"><span data-stu-id="0e2fc-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="0e2fc-147">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-147">Member</span></span> |
| [<span data-ttu-id="0e2fc-148">itemType</span><span class="sxs-lookup"><span data-stu-id="0e2fc-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="0e2fc-149">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-149">Member</span></span> |
| [<span data-ttu-id="0e2fc-150">location</span><span class="sxs-lookup"><span data-stu-id="0e2fc-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="0e2fc-151">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-151">Member</span></span> |
| [<span data-ttu-id="0e2fc-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="0e2fc-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="0e2fc-153">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-153">Member</span></span> |
| [<span data-ttu-id="0e2fc-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="0e2fc-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="0e2fc-155">Member</span><span class="sxs-lookup"><span data-stu-id="0e2fc-155">Member</span></span> |
| [<span data-ttu-id="0e2fc-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="0e2fc-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0e2fc-157">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-157">Member</span></span> |
| [<span data-ttu-id="0e2fc-158">organizer</span><span class="sxs-lookup"><span data-stu-id="0e2fc-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="0e2fc-159">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-159">Member</span></span> |
| [<span data-ttu-id="0e2fc-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="0e2fc-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="0e2fc-161">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-161">Member</span></span> |
| [<span data-ttu-id="0e2fc-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="0e2fc-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0e2fc-163">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-163">Member</span></span> |
| [<span data-ttu-id="0e2fc-164">sender</span><span class="sxs-lookup"><span data-stu-id="0e2fc-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="0e2fc-165">Member</span><span class="sxs-lookup"><span data-stu-id="0e2fc-165">Member</span></span> |
| [<span data-ttu-id="0e2fc-166">seriesId</span><span class="sxs-lookup"><span data-stu-id="0e2fc-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="0e2fc-167">Member</span><span class="sxs-lookup"><span data-stu-id="0e2fc-167">Member</span></span> |
| [<span data-ttu-id="0e2fc-168">start</span><span class="sxs-lookup"><span data-stu-id="0e2fc-168">start</span></span>](#start-datetime) | <span data-ttu-id="0e2fc-169">Member</span><span class="sxs-lookup"><span data-stu-id="0e2fc-169">Member</span></span> |
| [<span data-ttu-id="0e2fc-170">subject</span><span class="sxs-lookup"><span data-stu-id="0e2fc-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="0e2fc-171">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-171">Member</span></span> |
| [<span data-ttu-id="0e2fc-172">to</span><span class="sxs-lookup"><span data-stu-id="0e2fc-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0e2fc-173">Membre</span><span class="sxs-lookup"><span data-stu-id="0e2fc-173">Member</span></span> |
| [<span data-ttu-id="0e2fc-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="0e2fc-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-175">Method</span></span> |
| [<span data-ttu-id="0e2fc-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="0e2fc-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="0e2fc-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-177">Method</span></span> |
| [<span data-ttu-id="0e2fc-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="0e2fc-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-179">Method</span></span> |
| [<span data-ttu-id="0e2fc-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="0e2fc-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-181">Method</span></span> |
| [<span data-ttu-id="0e2fc-182">close</span><span class="sxs-lookup"><span data-stu-id="0e2fc-182">close</span></span>](#close) | <span data-ttu-id="0e2fc-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-183">Method</span></span> |
| [<span data-ttu-id="0e2fc-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="0e2fc-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="0e2fc-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-185">Method</span></span> |
| [<span data-ttu-id="0e2fc-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="0e2fc-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="0e2fc-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-187">Method</span></span> |
| [<span data-ttu-id="0e2fc-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="0e2fc-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-189">Method</span></span> |
| [<span data-ttu-id="0e2fc-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="0e2fc-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-191">Method</span></span> |
| [<span data-ttu-id="0e2fc-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="0e2fc-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-193">Method</span></span> |
| [<span data-ttu-id="0e2fc-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="0e2fc-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="0e2fc-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-195">Method</span></span> |
| [<span data-ttu-id="0e2fc-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="0e2fc-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0e2fc-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-197">Method</span></span> |
| [<span data-ttu-id="0e2fc-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="0e2fc-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0e2fc-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-199">Method</span></span> |
| [<span data-ttu-id="0e2fc-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="0e2fc-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-201">Method</span></span> |
| [<span data-ttu-id="0e2fc-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0e2fc-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="0e2fc-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-203">Method</span></span> |
| [<span data-ttu-id="0e2fc-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="0e2fc-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="0e2fc-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-205">Method</span></span> |
| [<span data-ttu-id="0e2fc-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="0e2fc-207">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-207">Method</span></span> |
| [<span data-ttu-id="0e2fc-208">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="0e2fc-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="0e2fc-209">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-209">Method</span></span> |
| [<span data-ttu-id="0e2fc-210">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0e2fc-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="0e2fc-211">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-211">Method</span></span> |
| [<span data-ttu-id="0e2fc-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="0e2fc-213">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-213">Method</span></span> |
| [<span data-ttu-id="0e2fc-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="0e2fc-215">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-215">Method</span></span> |
| [<span data-ttu-id="0e2fc-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="0e2fc-217">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-217">Method</span></span> |
| [<span data-ttu-id="0e2fc-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="0e2fc-219">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-219">Method</span></span> |
| [<span data-ttu-id="0e2fc-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="0e2fc-221">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-221">Method</span></span> |
| [<span data-ttu-id="0e2fc-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0e2fc-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="0e2fc-223">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e2fc-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="0e2fc-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-224">Example</span></span>

<span data-ttu-id="0e2fc-225">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="0e2fc-226">Members</span><span class="sxs-lookup"><span data-stu-id="0e2fc-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="0e2fc-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="0e2fc-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="0e2fc-228">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="0e2fc-229">Mode Lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-230">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="0e2fc-231">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-232">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-232">Type</span></span>

*   <span data-ttu-id="0e2fc-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="0e2fc-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-234">Requirements</span></span>

|<span data-ttu-id="0e2fc-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-235">Requirement</span></span>|<span data-ttu-id="0e2fc-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-238">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-238">1.0</span></span>|
|[<span data-ttu-id="0e2fc-239">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-240">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-242">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-243">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-243">Example</span></span>

<span data-ttu-id="0e2fc-244">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="0e2fc-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-246">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="0e2fc-247">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-247">Compose mode only.</span></span>

<span data-ttu-id="0e2fc-248">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-248">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0e2fc-249">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-249">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0e2fc-250">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-250">Get 500 members maximum.</span></span>
- <span data-ttu-id="0e2fc-251">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-251">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-252">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-252">Type</span></span>

*   [<span data-ttu-id="0e2fc-253">Destinataires</span><span class="sxs-lookup"><span data-stu-id="0e2fc-253">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="0e2fc-254">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-254">Requirements</span></span>

|<span data-ttu-id="0e2fc-255">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-255">Requirement</span></span>|<span data-ttu-id="0e2fc-256">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-256">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-257">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-257">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-258">1.1</span><span class="sxs-lookup"><span data-stu-id="0e2fc-258">1.1</span></span>|
|[<span data-ttu-id="0e2fc-259">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-259">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-260">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-260">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-261">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-261">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-262">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-262">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-263">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-263">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-18"></a><span data-ttu-id="0e2fc-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-265">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-265">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-266">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-266">Type</span></span>

*   [<span data-ttu-id="0e2fc-267">Body</span><span class="sxs-lookup"><span data-stu-id="0e2fc-267">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="0e2fc-268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-268">Requirements</span></span>

|<span data-ttu-id="0e2fc-269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-269">Requirement</span></span>|<span data-ttu-id="0e2fc-270">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-272">1.1</span><span class="sxs-lookup"><span data-stu-id="0e2fc-272">1.1</span></span>|
|[<span data-ttu-id="0e2fc-273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-274">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-276">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-277">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-277">Example</span></span>

<span data-ttu-id="0e2fc-278">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-278">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="0e2fc-279">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-279">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategoriesviewoutlook-js-18"></a><span data-ttu-id="0e2fc-280">Catégories : [catégories](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-280">categories: [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-281">Obtient un objet qui fournit des méthodes pour la gestion des catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-281">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-282">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-282">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-283">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-283">Type</span></span>

*   [<span data-ttu-id="0e2fc-284">Categories</span><span class="sxs-lookup"><span data-stu-id="0e2fc-284">Categories</span></span>](/javascript/api/outlook/office.categories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="0e2fc-285">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-285">Requirements</span></span>

|<span data-ttu-id="0e2fc-286">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-286">Requirement</span></span>|<span data-ttu-id="0e2fc-287">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-288">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-289">1.8</span><span class="sxs-lookup"><span data-stu-id="0e2fc-289">1.8</span></span>|
|[<span data-ttu-id="0e2fc-290">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-291">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-293">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-294">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-294">Example</span></span>

<span data-ttu-id="0e2fc-295">Cet exemple obtient les catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-295">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="0e2fc-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-297">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-297">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="0e2fc-298">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-298">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-299">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-299">Read mode</span></span>

<span data-ttu-id="0e2fc-300">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-300">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="0e2fc-301">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-301">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0e2fc-302">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-302">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-303">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-303">Compose mode</span></span>

<span data-ttu-id="0e2fc-304">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-304">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="0e2fc-305">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-305">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0e2fc-306">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-306">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0e2fc-307">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-307">Get 500 members maximum.</span></span>
- <span data-ttu-id="0e2fc-308">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-308">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e2fc-309">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-309">Type</span></span>

*   <span data-ttu-id="0e2fc-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-311">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-311">Requirements</span></span>

|<span data-ttu-id="0e2fc-312">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-312">Requirement</span></span>|<span data-ttu-id="0e2fc-313">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-314">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-315">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-315">1.0</span></span>|
|[<span data-ttu-id="0e2fc-316">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-316">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-317">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-318">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-318">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-319">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-319">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="0e2fc-320">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-320">(nullable) conversationId: String</span></span>

<span data-ttu-id="0e2fc-321">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-321">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="0e2fc-p109">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="0e2fc-p110">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-326">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-326">Type</span></span>

*   <span data-ttu-id="0e2fc-327">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-327">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-328">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-328">Requirements</span></span>

|<span data-ttu-id="0e2fc-329">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-329">Requirement</span></span>|<span data-ttu-id="0e2fc-330">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-330">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-331">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-331">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-332">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-332">1.0</span></span>|
|[<span data-ttu-id="0e2fc-333">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-333">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-334">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-334">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-335">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-335">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-336">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-336">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-337">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-337">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="0e2fc-338">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="0e2fc-338">dateTimeCreated: Date</span></span>

<span data-ttu-id="0e2fc-p111">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-341">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-341">Type</span></span>

*   <span data-ttu-id="0e2fc-342">Date</span><span class="sxs-lookup"><span data-stu-id="0e2fc-342">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-343">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-343">Requirements</span></span>

|<span data-ttu-id="0e2fc-344">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-344">Requirement</span></span>|<span data-ttu-id="0e2fc-345">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-346">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-347">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-347">1.0</span></span>|
|[<span data-ttu-id="0e2fc-348">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-349">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-350">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-351">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-352">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-352">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="0e2fc-353">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="0e2fc-353">dateTimeModified: Date</span></span>

<span data-ttu-id="0e2fc-p112">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-356">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-356">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-357">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-357">Type</span></span>

*   <span data-ttu-id="0e2fc-358">Date</span><span class="sxs-lookup"><span data-stu-id="0e2fc-358">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-359">Requirements</span></span>

|<span data-ttu-id="0e2fc-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-360">Requirement</span></span>|<span data-ttu-id="0e2fc-361">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-362">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-363">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-363">1.0</span></span>|
|[<span data-ttu-id="0e2fc-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-365">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-367">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-368">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-368">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="0e2fc-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-370">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-370">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="0e2fc-p113">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-373">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-373">Read mode</span></span>

<span data-ttu-id="0e2fc-374">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-374">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-375">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-375">Compose mode</span></span>

<span data-ttu-id="0e2fc-376">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-376">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="0e2fc-377">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-377">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0e2fc-378">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-378">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0e2fc-379">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-379">Type</span></span>

*   <span data-ttu-id="0e2fc-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-381">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-381">Requirements</span></span>

|<span data-ttu-id="0e2fc-382">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-382">Requirement</span></span>|<span data-ttu-id="0e2fc-383">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-383">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-384">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-384">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-385">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-385">1.0</span></span>|
|[<span data-ttu-id="0e2fc-386">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-386">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-387">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-387">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-388">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-388">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-389">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-389">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocationviewoutlook-js-18"></a><span data-ttu-id="0e2fc-390">enhancedLocation : [enhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-390">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-391">Obtient ou définit les emplacements d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-391">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-392">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-392">Read mode</span></span>

<span data-ttu-id="0e2fc-393">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) qui vous permet d’obtenir l’ensemble des emplacements (chacun représenté par un objet [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) ) associé au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-393">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-394">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-394">Compose mode</span></span>

<span data-ttu-id="0e2fc-395">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) qui fournit des méthodes pour obtenir, supprimer ou ajouter des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-396">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-396">Type</span></span>

*   [<span data-ttu-id="0e2fc-397">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="0e2fc-397">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="0e2fc-398">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-398">Requirements</span></span>

|<span data-ttu-id="0e2fc-399">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-399">Requirement</span></span>|<span data-ttu-id="0e2fc-400">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-400">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-401">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-401">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-402">1.8</span><span class="sxs-lookup"><span data-stu-id="0e2fc-402">1.8</span></span>|
|[<span data-ttu-id="0e2fc-403">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-403">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-404">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-404">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-405">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-405">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-406">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-406">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-407">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-407">Example</span></span>

<span data-ttu-id="0e2fc-408">L’exemple suivant obtient les emplacements actuels associés au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-408">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18fromjavascriptapioutlookofficefromviewoutlook-js-18"></a><span data-ttu-id="0e2fc-409">from : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[from](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-409">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-410">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-410">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="0e2fc-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-413">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-413">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-414">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-414">Read mode</span></span>

<span data-ttu-id="0e2fc-415">La `from` propriété renvoie un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-415">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-416">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-416">Compose mode</span></span>

<span data-ttu-id="0e2fc-417">La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-417">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e2fc-418">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-418">Type</span></span>

*   <span data-ttu-id="0e2fc-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [à partir de](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-420">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-420">Requirements</span></span>

|<span data-ttu-id="0e2fc-421">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-421">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="0e2fc-422">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-423">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-423">1.0</span></span>|<span data-ttu-id="0e2fc-424">1.7</span><span class="sxs-lookup"><span data-stu-id="0e2fc-424">1.7</span></span>|
|[<span data-ttu-id="0e2fc-425">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-426">ReadItem</span></span>|<span data-ttu-id="0e2fc-427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-427">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e2fc-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-429">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-429">Read</span></span>|<span data-ttu-id="0e2fc-430">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-430">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheadersviewoutlook-js-18"></a><span data-ttu-id="0e2fc-431">internetHeaders : [internetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-431">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-432">Obtient ou définit les en-têtes Internet personnalisés d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-432">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="0e2fc-433">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-433">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-434">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-434">Type</span></span>

*   [<span data-ttu-id="0e2fc-435">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="0e2fc-435">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="0e2fc-436">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-436">Requirements</span></span>

|<span data-ttu-id="0e2fc-437">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-437">Requirement</span></span>|<span data-ttu-id="0e2fc-438">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-439">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-440">1.8</span><span class="sxs-lookup"><span data-stu-id="0e2fc-440">1.8</span></span>|
|[<span data-ttu-id="0e2fc-441">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-442">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-443">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-444">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-444">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-445">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-445">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="0e2fc-446">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-446">internetMessageId: String</span></span>

<span data-ttu-id="0e2fc-p116">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-449">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-449">Type</span></span>

*   <span data-ttu-id="0e2fc-450">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-450">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-451">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-451">Requirements</span></span>

|<span data-ttu-id="0e2fc-452">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-452">Requirement</span></span>|<span data-ttu-id="0e2fc-453">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-453">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-454">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-454">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-455">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-455">1.0</span></span>|
|[<span data-ttu-id="0e2fc-456">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-456">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-457">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-457">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-458">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-458">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-459">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-459">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-460">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-460">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="0e2fc-461">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-461">itemClass: String</span></span>

<span data-ttu-id="0e2fc-p117">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="0e2fc-p118">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="0e2fc-466">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-466">Type</span></span>|<span data-ttu-id="0e2fc-467">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-467">Description</span></span>|<span data-ttu-id="0e2fc-468">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="0e2fc-468">item class</span></span>|
|---|---|---|
|<span data-ttu-id="0e2fc-469">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="0e2fc-469">Appointment items</span></span>|<span data-ttu-id="0e2fc-470">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-470">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="0e2fc-471">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="0e2fc-471">Message items</span></span>|<span data-ttu-id="0e2fc-472">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-472">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="0e2fc-473">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-473">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-474">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-474">Type</span></span>

*   <span data-ttu-id="0e2fc-475">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-475">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-476">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-476">Requirements</span></span>

|<span data-ttu-id="0e2fc-477">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-477">Requirement</span></span>|<span data-ttu-id="0e2fc-478">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-478">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-479">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-479">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-480">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-480">1.0</span></span>|
|[<span data-ttu-id="0e2fc-481">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-481">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-482">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-482">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-483">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-483">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-484">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-484">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-485">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-485">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="0e2fc-486">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-486">(nullable) itemId: String</span></span>

<span data-ttu-id="0e2fc-487">Obtient l' [identificateur d’élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) pour l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-487">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="0e2fc-488">Mode Lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-488">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-489">L’identificateur renvoyé par la `itemId` propriété est identique à l’identificateur d' [élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-489">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="0e2fc-490">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-490">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="0e2fc-491">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-491">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0e2fc-492">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-492">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="0e2fc-p121">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-495">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-495">Type</span></span>

*   <span data-ttu-id="0e2fc-496">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-496">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-497">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-497">Requirements</span></span>

|<span data-ttu-id="0e2fc-498">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-498">Requirement</span></span>|<span data-ttu-id="0e2fc-499">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-499">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-500">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-500">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-501">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-501">1.0</span></span>|
|[<span data-ttu-id="0e2fc-502">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-502">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-503">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-503">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-504">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-504">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-505">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-505">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-506">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-506">Example</span></span>

<span data-ttu-id="0e2fc-p122">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-18"></a><span data-ttu-id="0e2fc-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-510">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-510">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="0e2fc-511">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-511">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-512">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-512">Type</span></span>

*   [<span data-ttu-id="0e2fc-513">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="0e2fc-513">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="0e2fc-514">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-514">Requirements</span></span>

|<span data-ttu-id="0e2fc-515">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-515">Requirement</span></span>|<span data-ttu-id="0e2fc-516">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-516">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-517">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-517">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-518">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-518">1.0</span></span>|
|[<span data-ttu-id="0e2fc-519">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-519">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-520">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-520">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-521">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-521">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-522">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-522">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-523">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-523">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-18"></a><span data-ttu-id="0e2fc-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-525">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-525">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-526">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-526">Read mode</span></span>

<span data-ttu-id="0e2fc-527">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-527">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-528">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-528">Compose mode</span></span>

<span data-ttu-id="0e2fc-529">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-529">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e2fc-530">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-530">Type</span></span>

*   <span data-ttu-id="0e2fc-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-532">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-532">Requirements</span></span>

|<span data-ttu-id="0e2fc-533">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-533">Requirement</span></span>|<span data-ttu-id="0e2fc-534">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-535">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-536">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-536">1.0</span></span>|
|[<span data-ttu-id="0e2fc-537">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-537">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-538">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-539">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-539">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-540">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-540">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="0e2fc-541">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-541">normalizedSubject: String</span></span>

<span data-ttu-id="0e2fc-p123">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="0e2fc-p124">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-546">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-546">Type</span></span>

*   <span data-ttu-id="0e2fc-547">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-547">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-548">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-548">Requirements</span></span>

|<span data-ttu-id="0e2fc-549">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-549">Requirement</span></span>|<span data-ttu-id="0e2fc-550">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-551">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-552">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-552">1.0</span></span>|
|[<span data-ttu-id="0e2fc-553">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-553">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-554">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-555">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-555">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-556">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-556">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-557">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-557">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-18"></a><span data-ttu-id="0e2fc-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-559">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-559">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-560">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-560">Type</span></span>

*   [<span data-ttu-id="0e2fc-561">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="0e2fc-561">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="0e2fc-562">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-562">Requirements</span></span>

|<span data-ttu-id="0e2fc-563">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-563">Requirement</span></span>|<span data-ttu-id="0e2fc-564">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-565">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-566">1.3</span><span class="sxs-lookup"><span data-stu-id="0e2fc-566">1.3</span></span>|
|[<span data-ttu-id="0e2fc-567">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-567">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-568">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-569">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-569">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-570">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-570">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-571">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-571">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="0e2fc-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-573">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-573">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="0e2fc-574">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-574">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-575">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-575">Read mode</span></span>

<span data-ttu-id="0e2fc-576">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-576">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="0e2fc-577">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-577">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0e2fc-578">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-578">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-579">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-579">Compose mode</span></span>

<span data-ttu-id="0e2fc-580">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-580">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="0e2fc-581">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-581">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0e2fc-582">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-582">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0e2fc-583">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-583">Get 500 members maximum.</span></span>
- <span data-ttu-id="0e2fc-584">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-584">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e2fc-585">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-585">Type</span></span>

*   <span data-ttu-id="0e2fc-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-587">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-587">Requirements</span></span>

|<span data-ttu-id="0e2fc-588">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-588">Requirement</span></span>|<span data-ttu-id="0e2fc-589">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-589">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-590">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-591">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-591">1.0</span></span>|
|[<span data-ttu-id="0e2fc-592">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-592">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-593">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-593">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-594">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-595">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-595">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18organizerjavascriptapioutlookofficeorganizerviewoutlook-js-18"></a><span data-ttu-id="0e2fc-596">Organisateur : [](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[organisateur](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0e2fc-596">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-597">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-597">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-598">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-598">Read mode</span></span>

<span data-ttu-id="0e2fc-599">La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-599">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-600">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-600">Compose mode</span></span>

<span data-ttu-id="0e2fc-601">La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-601">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="0e2fc-602">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-602">Type</span></span>

*   <span data-ttu-id="0e2fc-603">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Organisateur](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0e2fc-603">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-604">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-604">Requirements</span></span>

|<span data-ttu-id="0e2fc-605">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-605">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="0e2fc-606">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-607">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-607">1.0</span></span>|<span data-ttu-id="0e2fc-608">1.7</span><span class="sxs-lookup"><span data-stu-id="0e2fc-608">1.7</span></span>|
|[<span data-ttu-id="0e2fc-609">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-610">ReadItem</span></span>|<span data-ttu-id="0e2fc-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e2fc-612">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-613">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-613">Read</span></span>|<span data-ttu-id="0e2fc-614">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-614">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-18"></a><span data-ttu-id="0e2fc-615">(Nullable) récurrence : [périodicité](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-615">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-616">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-616">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="0e2fc-617">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-617">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="0e2fc-618">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-618">Read and compose modes for appointment items.</span></span> <span data-ttu-id="0e2fc-619">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-619">Read mode for meeting request items.</span></span>

<span data-ttu-id="0e2fc-620">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-620">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="0e2fc-621">`null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-621">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="0e2fc-622">`undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-622">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="0e2fc-623">Remarque : les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-623">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="0e2fc-624">Remarque : si l’objet de périodicité `null`est, cela indique que l’objet est un rendez-vous unique ou une demande de réunion d’un seul rendez-vous et non d’une série.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-624">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-625">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-625">Read mode</span></span>

<span data-ttu-id="0e2fc-626">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-626">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that represents the appointment recurrence.</span></span> <span data-ttu-id="0e2fc-627">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-627">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-628">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-628">Compose mode</span></span>

<span data-ttu-id="0e2fc-629">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-629">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="0e2fc-630">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-630">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0e2fc-631">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-631">Type</span></span>

* [<span data-ttu-id="0e2fc-632">Instances</span><span class="sxs-lookup"><span data-stu-id="0e2fc-632">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)

|<span data-ttu-id="0e2fc-633">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-633">Requirement</span></span>|<span data-ttu-id="0e2fc-634">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-635">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-636">1.7</span><span class="sxs-lookup"><span data-stu-id="0e2fc-636">1.7</span></span>|
|[<span data-ttu-id="0e2fc-637">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-638">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-639">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-640">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-640">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="0e2fc-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-642">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-642">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="0e2fc-643">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-643">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-644">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-644">Read mode</span></span>

<span data-ttu-id="0e2fc-645">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-645">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="0e2fc-646">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-646">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0e2fc-647">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-647">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-648">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-648">Compose mode</span></span>

<span data-ttu-id="0e2fc-649">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-649">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="0e2fc-650">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-650">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0e2fc-651">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-651">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0e2fc-652">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-652">Get 500 members maximum.</span></span>
- <span data-ttu-id="0e2fc-653">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-653">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="0e2fc-654">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-654">Type</span></span>

*   <span data-ttu-id="0e2fc-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-656">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-656">Requirements</span></span>

|<span data-ttu-id="0e2fc-657">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-657">Requirement</span></span>|<span data-ttu-id="0e2fc-658">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-658">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-659">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-659">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-660">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-660">1.0</span></span>|
|[<span data-ttu-id="0e2fc-661">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-661">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-662">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-662">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-663">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-663">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-664">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-664">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18"></a><span data-ttu-id="0e2fc-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-p135">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="0e2fc-p136">Les propriétés [`from`](#from-emailaddressdetailsfrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-670">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-670">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-671">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-671">Type</span></span>

*   [<span data-ttu-id="0e2fc-672">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0e2fc-672">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="0e2fc-673">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-673">Requirements</span></span>

|<span data-ttu-id="0e2fc-674">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-674">Requirement</span></span>|<span data-ttu-id="0e2fc-675">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-676">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-677">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-677">1.0</span></span>|
|[<span data-ttu-id="0e2fc-678">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-679">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-679">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-680">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-681">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-681">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-682">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-682">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="0e2fc-683">(Nullable) seriesId : chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-683">(nullable) seriesId: String</span></span>

<span data-ttu-id="0e2fc-684">Obtient l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-684">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="0e2fc-685">Dans Outlook sur le Web et les clients de bureau `seriesId` , le renvoie l’ID des services Web Exchange (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-685">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="0e2fc-686">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-686">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-687">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-687">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0e2fc-688">La `seriesId` propriété n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-688">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="0e2fc-689">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-689">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0e2fc-690">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-690">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="0e2fc-691">La `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-691">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="0e2fc-692">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-692">Type</span></span>

* <span data-ttu-id="0e2fc-693">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-693">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-694">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-694">Requirements</span></span>

|<span data-ttu-id="0e2fc-695">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-695">Requirement</span></span>|<span data-ttu-id="0e2fc-696">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-697">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-698">1.7</span><span class="sxs-lookup"><span data-stu-id="0e2fc-698">1.7</span></span>|
|[<span data-ttu-id="0e2fc-699">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-700">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-701">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-702">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-702">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-703">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-703">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="0e2fc-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-705">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-705">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="0e2fc-p139">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-708">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-708">Read mode</span></span>

<span data-ttu-id="0e2fc-709">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-709">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-710">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-710">Compose mode</span></span>

<span data-ttu-id="0e2fc-711">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-711">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="0e2fc-712">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-712">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0e2fc-713">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-713">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0e2fc-714">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-714">Type</span></span>

*   <span data-ttu-id="0e2fc-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-716">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-716">Requirements</span></span>

|<span data-ttu-id="0e2fc-717">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-717">Requirement</span></span>|<span data-ttu-id="0e2fc-718">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-719">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-719">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-720">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-720">1.0</span></span>|
|[<span data-ttu-id="0e2fc-721">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-721">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-722">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-723">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-723">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-724">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-724">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-18"></a><span data-ttu-id="0e2fc-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-726">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-726">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="0e2fc-727">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-727">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-728">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-728">Read mode</span></span>

<span data-ttu-id="0e2fc-p140">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="0e2fc-731">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-731">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-732">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-732">Compose mode</span></span>
<span data-ttu-id="0e2fc-733">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-733">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="0e2fc-734">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-734">Type</span></span>

*   <span data-ttu-id="0e2fc-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-736">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-736">Requirements</span></span>

|<span data-ttu-id="0e2fc-737">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-737">Requirement</span></span>|<span data-ttu-id="0e2fc-738">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-738">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-739">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-739">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-740">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-740">1.0</span></span>|
|[<span data-ttu-id="0e2fc-741">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-741">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-742">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-742">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-743">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-743">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-744">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-744">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="0e2fc-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-746">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-746">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="0e2fc-747">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-747">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e2fc-748">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-748">Read mode</span></span>

<span data-ttu-id="0e2fc-749">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-749">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="0e2fc-750">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-750">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0e2fc-751">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-751">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e2fc-752">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-752">Compose mode</span></span>

<span data-ttu-id="0e2fc-753">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-753">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="0e2fc-754">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-754">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0e2fc-755">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-755">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0e2fc-756">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-756">Get 500 members maximum.</span></span>
- <span data-ttu-id="0e2fc-757">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-757">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e2fc-758">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-758">Type</span></span>

*   <span data-ttu-id="0e2fc-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-760">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-760">Requirements</span></span>

|<span data-ttu-id="0e2fc-761">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-761">Requirement</span></span>|<span data-ttu-id="0e2fc-762">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-762">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-763">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-763">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-764">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-764">1.0</span></span>|
|[<span data-ttu-id="0e2fc-765">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-765">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-766">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-766">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-767">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-767">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-768">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-768">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="0e2fc-769">Méthodes</span><span class="sxs-lookup"><span data-stu-id="0e2fc-769">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="0e2fc-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0e2fc-771">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-771">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0e2fc-772">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-772">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="0e2fc-773">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-774">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-774">Parameters</span></span>
|<span data-ttu-id="0e2fc-775">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-775">Name</span></span>|<span data-ttu-id="0e2fc-776">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-776">Type</span></span>|<span data-ttu-id="0e2fc-777">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-777">Attributes</span></span>|<span data-ttu-id="0e2fc-778">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-778">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="0e2fc-779">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-779">String</span></span>||<span data-ttu-id="0e2fc-p144">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="0e2fc-782">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-782">String</span></span>||<span data-ttu-id="0e2fc-p145">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0e2fc-785">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-785">Object</span></span>|<span data-ttu-id="0e2fc-786">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-786">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-787">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-787">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-788">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-788">Object</span></span>|<span data-ttu-id="0e2fc-789">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-789">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-790">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-790">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="0e2fc-791">Boolean</span><span class="sxs-lookup"><span data-stu-id="0e2fc-791">Boolean</span></span>|<span data-ttu-id="0e2fc-792">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-792">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-793">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-793">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-794">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-794">function</span></span>|<span data-ttu-id="0e2fc-795">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-795">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-796">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-796">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e2fc-797">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-797">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0e2fc-798">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-798">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e2fc-799">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-799">Errors</span></span>

|<span data-ttu-id="0e2fc-800">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-800">Error code</span></span>|<span data-ttu-id="0e2fc-801">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-801">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="0e2fc-802">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-802">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="0e2fc-803">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-803">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0e2fc-804">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-804">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-805">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-805">Requirements</span></span>

|<span data-ttu-id="0e2fc-806">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-806">Requirement</span></span>|<span data-ttu-id="0e2fc-807">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-808">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-809">1.1</span><span class="sxs-lookup"><span data-stu-id="0e2fc-809">1.1</span></span>|
|[<span data-ttu-id="0e2fc-810">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-811">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-811">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e2fc-812">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-813">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-813">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e2fc-814">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e2fc-814">Examples</span></span>

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

<span data-ttu-id="0e2fc-815">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-815">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="0e2fc-816">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-816">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0e2fc-817">Ajoute un fichier à partir du codage Base64 à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-817">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0e2fc-818">La `addFileAttachmentFromBase64Async` méthode charge le fichier à partir du codage Base64 et l’associe à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-818">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="0e2fc-819">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-819">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="0e2fc-820">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-820">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-821">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-821">Parameters</span></span>

|<span data-ttu-id="0e2fc-822">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-822">Name</span></span>|<span data-ttu-id="0e2fc-823">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-823">Type</span></span>|<span data-ttu-id="0e2fc-824">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-824">Attributes</span></span>|<span data-ttu-id="0e2fc-825">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-825">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="0e2fc-826">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-826">String</span></span>||<span data-ttu-id="0e2fc-827">Contenu encodé en base64 d’une image ou d’un fichier à ajouter à un message électronique ou à un événement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-827">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="0e2fc-828">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-828">String</span></span>||<span data-ttu-id="0e2fc-p147">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0e2fc-831">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-831">Object</span></span>|<span data-ttu-id="0e2fc-832">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-832">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-833">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-833">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-834">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-834">Object</span></span>|<span data-ttu-id="0e2fc-835">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-835">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-836">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-836">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="0e2fc-837">Boolean</span><span class="sxs-lookup"><span data-stu-id="0e2fc-837">Boolean</span></span>|<span data-ttu-id="0e2fc-838">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-838">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-839">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-839">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-840">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-840">function</span></span>|<span data-ttu-id="0e2fc-841">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-841">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-842">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-842">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e2fc-843">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-843">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0e2fc-844">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-844">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e2fc-845">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-845">Errors</span></span>

|<span data-ttu-id="0e2fc-846">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-846">Error code</span></span>|<span data-ttu-id="0e2fc-847">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-847">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="0e2fc-848">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-848">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="0e2fc-849">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-849">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0e2fc-850">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-850">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-851">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-851">Requirements</span></span>

|<span data-ttu-id="0e2fc-852">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-852">Requirement</span></span>|<span data-ttu-id="0e2fc-853">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-853">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-854">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-854">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-855">1.8</span><span class="sxs-lookup"><span data-stu-id="0e2fc-855">1.8</span></span>|
|[<span data-ttu-id="0e2fc-856">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-856">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-857">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-857">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e2fc-858">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-858">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-859">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-859">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e2fc-860">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e2fc-860">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="0e2fc-861">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-861">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="0e2fc-862">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-862">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="0e2fc-863">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-863">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-864">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-864">Parameters</span></span>

| <span data-ttu-id="0e2fc-865">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-865">Name</span></span> | <span data-ttu-id="0e2fc-866">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-866">Type</span></span> | <span data-ttu-id="0e2fc-867">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-867">Attributes</span></span> | <span data-ttu-id="0e2fc-868">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-868">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0e2fc-869">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0e2fc-869">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0e2fc-870">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-870">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="0e2fc-871">Fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-871">Function</span></span> || <span data-ttu-id="0e2fc-p148">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="0e2fc-875">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-875">Object</span></span> | <span data-ttu-id="0e2fc-876">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-876">&lt;optional&gt;</span></span> | <span data-ttu-id="0e2fc-877">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-877">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0e2fc-878">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-878">Object</span></span> | <span data-ttu-id="0e2fc-879">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-879">&lt;optional&gt;</span></span> | <span data-ttu-id="0e2fc-880">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-880">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0e2fc-881">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-881">function</span></span>| <span data-ttu-id="0e2fc-882">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-882">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-883">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-884">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-884">Requirements</span></span>

|<span data-ttu-id="0e2fc-885">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-885">Requirement</span></span>| <span data-ttu-id="0e2fc-886">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-887">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e2fc-888">1.7</span><span class="sxs-lookup"><span data-stu-id="0e2fc-888">1.7</span></span> |
|[<span data-ttu-id="0e2fc-889">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e2fc-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-890">ReadItem</span></span> |
|[<span data-ttu-id="0e2fc-891">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0e2fc-892">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-892">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="0e2fc-893">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-893">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="0e2fc-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0e2fc-895">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-895">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="0e2fc-p149">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="0e2fc-899">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-899">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="0e2fc-900">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-900">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-901">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-901">Parameters</span></span>

|<span data-ttu-id="0e2fc-902">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-902">Name</span></span>|<span data-ttu-id="0e2fc-903">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-903">Type</span></span>|<span data-ttu-id="0e2fc-904">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-904">Attributes</span></span>|<span data-ttu-id="0e2fc-905">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-905">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="0e2fc-906">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-906">String</span></span>||<span data-ttu-id="0e2fc-p150">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="0e2fc-909">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-909">String</span></span>||<span data-ttu-id="0e2fc-910">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-910">The subject of the item to be attached.</span></span> <span data-ttu-id="0e2fc-911">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-911">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0e2fc-912">Object</span><span class="sxs-lookup"><span data-stu-id="0e2fc-912">Object</span></span>|<span data-ttu-id="0e2fc-913">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-913">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-914">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-914">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-915">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-915">Object</span></span>|<span data-ttu-id="0e2fc-916">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-916">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-917">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-917">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-918">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-918">function</span></span>|<span data-ttu-id="0e2fc-919">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-919">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-920">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-920">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e2fc-921">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-921">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0e2fc-922">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-922">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e2fc-923">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-923">Errors</span></span>

|<span data-ttu-id="0e2fc-924">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-924">Error code</span></span>|<span data-ttu-id="0e2fc-925">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-925">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0e2fc-926">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-926">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-927">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-927">Requirements</span></span>

|<span data-ttu-id="0e2fc-928">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-928">Requirement</span></span>|<span data-ttu-id="0e2fc-929">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-930">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-931">1.1</span><span class="sxs-lookup"><span data-stu-id="0e2fc-931">1.1</span></span>|
|[<span data-ttu-id="0e2fc-932">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-932">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-933">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-933">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e2fc-934">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-934">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-935">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-935">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-936">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-936">Example</span></span>

<span data-ttu-id="0e2fc-937">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-937">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="0e2fc-938">close()</span><span class="sxs-lookup"><span data-stu-id="0e2fc-938">close()</span></span>

<span data-ttu-id="0e2fc-939">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-939">Closes the current item that is being composed.</span></span>

<span data-ttu-id="0e2fc-p152">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-942">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-942">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="0e2fc-943">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-943">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-944">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-944">Requirements</span></span>

|<span data-ttu-id="0e2fc-945">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-945">Requirement</span></span>|<span data-ttu-id="0e2fc-946">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-947">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-948">1.3</span><span class="sxs-lookup"><span data-stu-id="0e2fc-948">1.3</span></span>|
|[<span data-ttu-id="0e2fc-949">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-950">Restreinte</span><span class="sxs-lookup"><span data-stu-id="0e2fc-950">Restricted</span></span>|
|[<span data-ttu-id="0e2fc-951">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-952">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-952">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="0e2fc-953">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-953">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="0e2fc-954">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-954">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-955">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-955">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e2fc-956">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-956">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0e2fc-957">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-957">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="0e2fc-p153">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-961">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-961">Parameters</span></span>

|<span data-ttu-id="0e2fc-962">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-962">Name</span></span>|<span data-ttu-id="0e2fc-963">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-963">Type</span></span>|<span data-ttu-id="0e2fc-964">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-964">Attributes</span></span>|<span data-ttu-id="0e2fc-965">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-965">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="0e2fc-966">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0e2fc-966">String &#124; Object</span></span>||<span data-ttu-id="0e2fc-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0e2fc-969">**OU**</span><span class="sxs-lookup"><span data-stu-id="0e2fc-969">**OR**</span></span><br/><span data-ttu-id="0e2fc-p155">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="0e2fc-972">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-972">String</span></span>|<span data-ttu-id="0e2fc-973">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-973">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-p156">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="0e2fc-976">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-976">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="0e2fc-977">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-977">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-978">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-978">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="0e2fc-979">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-979">String</span></span>||<span data-ttu-id="0e2fc-p157">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="0e2fc-982">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-982">String</span></span>||<span data-ttu-id="0e2fc-983">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-983">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="0e2fc-984">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-984">String</span></span>||<span data-ttu-id="0e2fc-p158">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="0e2fc-987">Booléen</span><span class="sxs-lookup"><span data-stu-id="0e2fc-987">Boolean</span></span>||<span data-ttu-id="0e2fc-p159">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="0e2fc-990">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-990">String</span></span>||<span data-ttu-id="0e2fc-p160">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-994">function</span><span class="sxs-lookup"><span data-stu-id="0e2fc-994">function</span></span>|<span data-ttu-id="0e2fc-995">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-995">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-996">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-996">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-997">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-997">Requirements</span></span>

|<span data-ttu-id="0e2fc-998">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-998">Requirement</span></span>|<span data-ttu-id="0e2fc-999">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1000">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1001">1.0</span></span>|
|[<span data-ttu-id="0e2fc-1002">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1003">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1004">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1005">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1005">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e2fc-1006">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1006">Examples</span></span>

<span data-ttu-id="0e2fc-1007">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1007">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="0e2fc-1008">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1008">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="0e2fc-1009">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1009">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0e2fc-1010">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1010">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0e2fc-1011">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1011">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0e2fc-1012">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1012">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="0e2fc-1013">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1013">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="0e2fc-1014">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1014">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1015">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1015">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e2fc-1016">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1016">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0e2fc-1017">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1017">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="0e2fc-p161">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1021">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1021">Parameters</span></span>

|<span data-ttu-id="0e2fc-1022">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1022">Name</span></span>|<span data-ttu-id="0e2fc-1023">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1023">Type</span></span>|<span data-ttu-id="0e2fc-1024">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1024">Attributes</span></span>|<span data-ttu-id="0e2fc-1025">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1025">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="0e2fc-1026">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1026">String &#124; Object</span></span>||<span data-ttu-id="0e2fc-p162">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0e2fc-1029">**OU**</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1029">**OR**</span></span><br/><span data-ttu-id="0e2fc-p163">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="0e2fc-1032">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1032">String</span></span>|<span data-ttu-id="0e2fc-1033">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1033">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-p164">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="0e2fc-1036">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1036">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="0e2fc-1037">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1038">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1038">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="0e2fc-1039">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1039">String</span></span>||<span data-ttu-id="0e2fc-p165">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="0e2fc-1042">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1042">String</span></span>||<span data-ttu-id="0e2fc-1043">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1043">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="0e2fc-1044">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1044">String</span></span>||<span data-ttu-id="0e2fc-p166">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="0e2fc-1047">Booléen</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1047">Boolean</span></span>||<span data-ttu-id="0e2fc-p167">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="0e2fc-1050">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1050">String</span></span>||<span data-ttu-id="0e2fc-p168">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1054">function</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1054">function</span></span>|<span data-ttu-id="0e2fc-1055">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1055">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1056">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1056">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1057">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1057">Requirements</span></span>

|<span data-ttu-id="0e2fc-1058">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1058">Requirement</span></span>|<span data-ttu-id="0e2fc-1059">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1059">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1060">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1060">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1061">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1061">1.0</span></span>|
|[<span data-ttu-id="0e2fc-1062">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1062">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1063">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1063">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1064">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1064">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1065">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1065">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e2fc-1066">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1066">Examples</span></span>

<span data-ttu-id="0e2fc-1067">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1067">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="0e2fc-1068">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1068">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="0e2fc-1069">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1069">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0e2fc-1070">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1070">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0e2fc-1071">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1071">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0e2fc-1072">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1072">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="0e2fc-1073">getAllInternetHeadersAsync ([options], [Rappel])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1073">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="0e2fc-1074">Obtient tous les en-têtes Internet pour le message sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1074">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="0e2fc-1075">Mode Lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1075">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1076">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1076">Parameters</span></span>

|<span data-ttu-id="0e2fc-1077">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1077">Name</span></span>|<span data-ttu-id="0e2fc-1078">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1078">Type</span></span>|<span data-ttu-id="0e2fc-1079">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1079">Attributes</span></span>|<span data-ttu-id="0e2fc-1080">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1080">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e2fc-1081">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1081">Object</span></span>|<span data-ttu-id="0e2fc-1082">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1083">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1083">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-1084">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1084">Object</span></span>|<span data-ttu-id="0e2fc-1085">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1085">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1086">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1086">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1087">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1087">function</span></span>|<span data-ttu-id="0e2fc-1088">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1088">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1089">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="0e2fc-1090">En cas de réussite, les données des en-têtes Internet sont fournies dans la propriété asyncResult. Value sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1090">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="0e2fc-1091">Reportez-vous à la [norme RFC 2183](https://tools.ietf.org/html/rfc2183) pour les informations de mise en forme de la valeur de chaîne renvoyée.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1091">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="0e2fc-1092">En cas d’échec de l’appel, la propriété asyncResult. Error contient un code d’erreur correspondant à la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1092">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1093">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1093">Requirements</span></span>

|<span data-ttu-id="0e2fc-1094">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1094">Requirement</span></span>|<span data-ttu-id="0e2fc-1095">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1096">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1097">1.8</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1097">1.8</span></span>|
|[<span data-ttu-id="0e2fc-1098">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1099">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1100">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1101">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1102">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1102">Returns:</span></span>

<span data-ttu-id="0e2fc-1103">Les données des en-têtes Internet sous forme de chaîne formatée conformément à la [norme RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1103">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="0e2fc-1104">Type : String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1104">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0e2fc-1105">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1105">Example</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontentviewoutlook-js-18"></a><span data-ttu-id="0e2fc-1106">getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1106">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

<span data-ttu-id="0e2fc-1107">Obtient la pièce jointe spécifiée à partir d’un message ou d’un `AttachmentContent` rendez-vous et la renvoie en tant qu’objet.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1107">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="0e2fc-1108">La `getAttachmentContentAsync` méthode obtient la pièce jointe avec l’identificateur spécifié à partir de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1108">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0e2fc-1109">Il est recommandé d’utiliser l’identificateur pour récupérer une pièce jointe dans la même session que l’attachmentIds a été récupérée avec l' `getAttachmentsAsync` appel ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1109">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="0e2fc-1110">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1110">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0e2fc-1111">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1111">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1112">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1112">Parameters</span></span>

|<span data-ttu-id="0e2fc-1113">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1113">Name</span></span>|<span data-ttu-id="0e2fc-1114">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1114">Type</span></span>|<span data-ttu-id="0e2fc-1115">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1115">Attributes</span></span>|<span data-ttu-id="0e2fc-1116">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1116">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="0e2fc-1117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1117">String</span></span>||<span data-ttu-id="0e2fc-1118">Identificateur de la pièce jointe que vous souhaitez obtenir.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1118">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="0e2fc-1119">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1119">Object</span></span>|<span data-ttu-id="0e2fc-1120">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1121">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-1122">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1122">Object</span></span>|<span data-ttu-id="0e2fc-1123">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1124">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1125">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1125">function</span></span>|<span data-ttu-id="0e2fc-1126">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1126">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1127">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1128">Requirements</span></span>

|<span data-ttu-id="0e2fc-1129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1129">Requirement</span></span>|<span data-ttu-id="0e2fc-1130">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1132">1.8</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1132">1.8</span></span>|
|[<span data-ttu-id="0e2fc-1133">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1133">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1134">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1135">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1136">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1136">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1137">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1137">Returns:</span></span>

<span data-ttu-id="0e2fc-1138">Type : [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1138">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="0e2fc-1139">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1139">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="0e2fc-1140">getAttachmentsAsync ([options], [Rappel]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="0e2fc-1140">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="0e2fc-1141">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1141">Gets the item's attachments as an array.</span></span> <span data-ttu-id="0e2fc-1142">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1142">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1143">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1143">Parameters</span></span>

|<span data-ttu-id="0e2fc-1144">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1144">Name</span></span>|<span data-ttu-id="0e2fc-1145">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1145">Type</span></span>|<span data-ttu-id="0e2fc-1146">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1146">Attributes</span></span>|<span data-ttu-id="0e2fc-1147">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1147">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e2fc-1148">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1148">Object</span></span>|<span data-ttu-id="0e2fc-1149">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1149">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1150">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1150">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-1151">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1151">Object</span></span>|<span data-ttu-id="0e2fc-1152">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1152">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1153">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1153">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1154">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1154">function</span></span>|<span data-ttu-id="0e2fc-1155">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1155">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1156">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1157">Requirements</span></span>

|<span data-ttu-id="0e2fc-1158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1158">Requirement</span></span>|<span data-ttu-id="0e2fc-1159">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1159">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1161">1.8</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1161">1.8</span></span>|
|[<span data-ttu-id="0e2fc-1162">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1162">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1163">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1163">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1164">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1165">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1165">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1166">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1166">Returns:</span></span>

<span data-ttu-id="0e2fc-1167">Type : Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="0e2fc-1167">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="example"></a><span data-ttu-id="0e2fc-1168">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1168">Example</span></span>

<span data-ttu-id="0e2fc-1169">L’exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1169">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="0e2fc-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="0e2fc-1171">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1171">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1172">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1172">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1173">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1173">Requirements</span></span>

|<span data-ttu-id="0e2fc-1174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1174">Requirement</span></span>|<span data-ttu-id="0e2fc-1175">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1175">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1176">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1177">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1177">1.0</span></span>|
|[<span data-ttu-id="0e2fc-1178">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1178">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1179">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1181">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1181">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1182">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1182">Returns:</span></span>

<span data-ttu-id="0e2fc-1183">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1183">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="0e2fc-1184">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1184">Example</span></span>

<span data-ttu-id="0e2fc-1185">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1185">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="0e2fc-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="0e2fc-1187">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1187">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1188">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1188">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1189">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1189">Parameters</span></span>

|<span data-ttu-id="0e2fc-1190">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1190">Name</span></span>|<span data-ttu-id="0e2fc-1191">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1191">Type</span></span>|<span data-ttu-id="0e2fc-1192">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1192">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="0e2fc-1193">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1193">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.8)|<span data-ttu-id="0e2fc-1194">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1194">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1195">Requirements</span></span>

|<span data-ttu-id="0e2fc-1196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1196">Requirement</span></span>|<span data-ttu-id="0e2fc-1197">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1197">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1199">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1199">1.0</span></span>|
|[<span data-ttu-id="0e2fc-1200">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1201">Restreinte</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1201">Restricted</span></span>|
|[<span data-ttu-id="0e2fc-1202">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1203">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1203">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1204">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1204">Returns:</span></span>

<span data-ttu-id="0e2fc-1205">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1205">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="0e2fc-1206">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1206">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="0e2fc-1207">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1207">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="0e2fc-1208">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1208">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="0e2fc-1209">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1209">Value of `entityType`</span></span>|<span data-ttu-id="0e2fc-1210">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1210">Type of objects in returned array</span></span>|<span data-ttu-id="0e2fc-1211">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1211">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="0e2fc-1212">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1212">String</span></span>|<span data-ttu-id="0e2fc-1213">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1213">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="0e2fc-1214">Contact</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1214">Contact</span></span>|<span data-ttu-id="0e2fc-1215">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1215">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="0e2fc-1216">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1216">String</span></span>|<span data-ttu-id="0e2fc-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1217">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="0e2fc-1218">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1218">MeetingSuggestion</span></span>|<span data-ttu-id="0e2fc-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1219">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="0e2fc-1220">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1220">PhoneNumber</span></span>|<span data-ttu-id="0e2fc-1221">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1221">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="0e2fc-1222">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1222">TaskSuggestion</span></span>|<span data-ttu-id="0e2fc-1223">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1223">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="0e2fc-1224">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1224">String</span></span>|<span data-ttu-id="0e2fc-1225">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1225">**Restricted**</span></span>|

<span data-ttu-id="0e2fc-1226">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="0e2fc-1226">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

##### <a name="example"></a><span data-ttu-id="0e2fc-1227">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1227">Example</span></span>

<span data-ttu-id="0e2fc-1228">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1228">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="0e2fc-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="0e2fc-1230">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1230">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1231">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1231">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e2fc-1232">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1232">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1233">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1233">Parameters</span></span>

|<span data-ttu-id="0e2fc-1234">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1234">Name</span></span>|<span data-ttu-id="0e2fc-1235">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1235">Type</span></span>|<span data-ttu-id="0e2fc-1236">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1236">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="0e2fc-1237">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1237">String</span></span>|<span data-ttu-id="0e2fc-1238">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1238">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1239">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1239">Requirements</span></span>

|<span data-ttu-id="0e2fc-1240">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1240">Requirement</span></span>|<span data-ttu-id="0e2fc-1241">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1241">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1242">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1243">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1243">1.0</span></span>|
|[<span data-ttu-id="0e2fc-1244">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1245">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1245">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1246">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1247">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1247">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1248">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1248">Returns:</span></span>

<span data-ttu-id="0e2fc-p174">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="0e2fc-1251">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="0e2fc-1251">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="0e2fc-1252">getItemIdAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1252">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="0e2fc-1253">Obtient de manière asynchrone l’ID d’un élément enregistré.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1253">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="0e2fc-1254">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1254">Compose mode only.</span></span>

<span data-ttu-id="0e2fc-1255">Lorsqu’elle est appelée, cette méthode renvoie l’ID de l’élément par le biais de la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1255">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1256">Si votre complément appelle `getItemIdAsync` sur un élément en mode composition (par exemple, pour obtenir un à utiliser avec `itemId` EWS ou l’API REST), sachez que lorsque Outlook est en mode mis en cache, l’élément peut prendre un certain temps avant la synchronisation de l’élément avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1256">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="0e2fc-1257">Tant que l’élément n’est pas synchronisé `itemId` , le n’est pas reconnu et son utilisation renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1257">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1258">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1258">Parameters</span></span>

|<span data-ttu-id="0e2fc-1259">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1259">Name</span></span>|<span data-ttu-id="0e2fc-1260">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1260">Type</span></span>|<span data-ttu-id="0e2fc-1261">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1261">Attributes</span></span>|<span data-ttu-id="0e2fc-1262">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1262">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e2fc-1263">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1263">Object</span></span>|<span data-ttu-id="0e2fc-1264">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1264">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1265">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1265">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-1266">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1266">Object</span></span>|<span data-ttu-id="0e2fc-1267">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1268">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1268">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1269">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1269">function</span></span>||<span data-ttu-id="0e2fc-1270">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e2fc-1271">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1271">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e2fc-1272">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1272">Errors</span></span>

|<span data-ttu-id="0e2fc-1273">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1273">Error code</span></span>|<span data-ttu-id="0e2fc-1274">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1274">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="0e2fc-1275">L’ID ne peut pas être récupéré tant que l’élément n’est pas enregistré.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1275">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1276">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1276">Requirements</span></span>

|<span data-ttu-id="0e2fc-1277">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1277">Requirement</span></span>|<span data-ttu-id="0e2fc-1278">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1279">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1280">1.8</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1280">1.8</span></span>|
|[<span data-ttu-id="0e2fc-1281">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1282">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1283">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1284">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e2fc-1285">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1285">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="0e2fc-1286">L’exemple suivant montre la structure du `result` paramètre transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1286">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="0e2fc-1287">La `value` propriété contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1287">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="0e2fc-1288">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1288">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="0e2fc-1289">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1289">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1290">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1290">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e2fc-p178">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0e2fc-1294">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1294">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0e2fc-1295">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1295">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0e2fc-p179">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1299">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1299">Requirements</span></span>

|<span data-ttu-id="0e2fc-1300">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1300">Requirement</span></span>|<span data-ttu-id="0e2fc-1301">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1302">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1303">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1303">1.0</span></span>|
|[<span data-ttu-id="0e2fc-1304">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1305">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1306">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1307">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1307">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1308">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1308">Returns:</span></span>

<span data-ttu-id="0e2fc-p180">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="0e2fc-1311">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1311">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0e2fc-1312">Object</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1312">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0e2fc-1313">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1313">Example</span></span>

<span data-ttu-id="0e2fc-1314">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1314">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="0e2fc-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="0e2fc-1316">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1316">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1317">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1317">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e2fc-1318">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1318">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="0e2fc-p181">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1321">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1321">Parameters</span></span>

|<span data-ttu-id="0e2fc-1322">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1322">Name</span></span>|<span data-ttu-id="0e2fc-1323">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1323">Type</span></span>|<span data-ttu-id="0e2fc-1324">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1324">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="0e2fc-1325">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1325">String</span></span>|<span data-ttu-id="0e2fc-1326">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1326">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1327">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1327">Requirements</span></span>

|<span data-ttu-id="0e2fc-1328">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1328">Requirement</span></span>|<span data-ttu-id="0e2fc-1329">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1329">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1330">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1331">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1331">1.0</span></span>|
|[<span data-ttu-id="0e2fc-1332">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1332">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1333">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1334">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1334">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1335">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1335">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1336">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1336">Returns:</span></span>

<span data-ttu-id="0e2fc-1337">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1337">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="0e2fc-1338">Type : Array.< String ></span><span class="sxs-lookup"><span data-stu-id="0e2fc-1338">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="0e2fc-1339">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1339">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="0e2fc-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="0e2fc-1341">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1341">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="0e2fc-p182">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1344">Dans Outlook sur le Web, la méthode renvoie la chaîne « NULL » si aucun texte n’est sélectionné, mais que le curseur se trouve dans le corps.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1344">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="0e2fc-1345">Pour vérifier cette situation, incluez un code similaire à celui-ci :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1345">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1346">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1346">Parameters</span></span>

|<span data-ttu-id="0e2fc-1347">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1347">Name</span></span>|<span data-ttu-id="0e2fc-1348">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1348">Type</span></span>|<span data-ttu-id="0e2fc-1349">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1349">Attributes</span></span>|<span data-ttu-id="0e2fc-1350">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1350">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="0e2fc-1351">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1351">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="0e2fc-p184">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p184">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="0e2fc-1355">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1355">Object</span></span>|<span data-ttu-id="0e2fc-1356">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1356">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1357">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1357">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-1358">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1358">Object</span></span>|<span data-ttu-id="0e2fc-1359">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1359">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1360">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1360">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1361">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1361">function</span></span>||<span data-ttu-id="0e2fc-1362">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1362">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e2fc-1363">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1363">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="0e2fc-1364">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1364">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1365">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1365">Requirements</span></span>

|<span data-ttu-id="0e2fc-1366">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1366">Requirement</span></span>|<span data-ttu-id="0e2fc-1367">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1367">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1368">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1369">1.2</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1369">1.2</span></span>|
|[<span data-ttu-id="0e2fc-1370">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1371">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1372">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1373">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1373">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1374">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1374">Returns:</span></span>

<span data-ttu-id="0e2fc-1375">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1375">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="0e2fc-1376">Type : String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1376">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0e2fc-1377">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1377">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="0e2fc-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="0e2fc-1379">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1379">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="0e2fc-1380">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1380">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1381">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1381">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1382">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1382">Requirements</span></span>

|<span data-ttu-id="0e2fc-1383">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1383">Requirement</span></span>|<span data-ttu-id="0e2fc-1384">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1384">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1385">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1386">1.6</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1386">1.6</span></span>|
|[<span data-ttu-id="0e2fc-1387">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1387">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1388">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1389">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1389">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1390">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1390">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1391">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1391">Returns:</span></span>

<span data-ttu-id="0e2fc-1392">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1392">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="0e2fc-1393">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1393">Example</span></span>

<span data-ttu-id="0e2fc-1394">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1394">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="0e2fc-1395">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1395">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="0e2fc-p187">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p187">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1398">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1398">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e2fc-p188">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p188">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0e2fc-1402">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1402">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0e2fc-1403">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1403">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0e2fc-p189">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p189">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1407">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1407">Requirements</span></span>

|<span data-ttu-id="0e2fc-1408">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1408">Requirement</span></span>|<span data-ttu-id="0e2fc-1409">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1410">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1411">1.6</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1411">1.6</span></span>|
|[<span data-ttu-id="0e2fc-1412">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1413">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1414">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1415">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1415">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e2fc-1416">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1416">Returns:</span></span>

<span data-ttu-id="0e2fc-p190">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p190">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="0e2fc-1419">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1419">Example</span></span>

<span data-ttu-id="0e2fc-1420">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1420">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="0e2fc-1421">getSharedPropertiesAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1421">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="0e2fc-1422">Obtient les propriétés du rendez-vous ou du message sélectionné dans un dossier partagé, un calendrier ou une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1422">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1423">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1423">Parameters</span></span>

|<span data-ttu-id="0e2fc-1424">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1424">Name</span></span>|<span data-ttu-id="0e2fc-1425">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1425">Type</span></span>|<span data-ttu-id="0e2fc-1426">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1426">Attributes</span></span>|<span data-ttu-id="0e2fc-1427">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1427">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e2fc-1428">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1428">Object</span></span>|<span data-ttu-id="0e2fc-1429">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1429">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1430">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1430">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-1431">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1431">Object</span></span>|<span data-ttu-id="0e2fc-1432">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1433">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1433">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1434">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1434">function</span></span>||<span data-ttu-id="0e2fc-1435">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e2fc-1436">Les propriétés partagées sont fournies sous [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) la forme d' `asyncResult.value` un objet dans la propriété.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1436">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0e2fc-1437">Cet objet peut être utilisé pour obtenir les propriétés partagées de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1437">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1438">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1438">Requirements</span></span>

|<span data-ttu-id="0e2fc-1439">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1439">Requirement</span></span>|<span data-ttu-id="0e2fc-1440">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1440">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1441">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1442">1.8</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1442">1.8</span></span>|
|[<span data-ttu-id="0e2fc-1443">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1444">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1445">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1446">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1446">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-1447">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1447">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="0e2fc-1448">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1448">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="0e2fc-1449">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1449">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="0e2fc-p192">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p192">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1453">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1453">Parameters</span></span>

|<span data-ttu-id="0e2fc-1454">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1454">Name</span></span>|<span data-ttu-id="0e2fc-1455">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1455">Type</span></span>|<span data-ttu-id="0e2fc-1456">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1456">Attributes</span></span>|<span data-ttu-id="0e2fc-1457">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1457">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="0e2fc-1458">function</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1458">function</span></span>||<span data-ttu-id="0e2fc-1459">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1459">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e2fc-1460">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1460">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0e2fc-1461">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1461">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="0e2fc-1462">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1462">Object</span></span>|<span data-ttu-id="0e2fc-1463">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1464">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1464">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="0e2fc-1465">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1465">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1466">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1466">Requirements</span></span>

|<span data-ttu-id="0e2fc-1467">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1467">Requirement</span></span>|<span data-ttu-id="0e2fc-1468">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1468">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1469">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1470">1.0</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1470">1.0</span></span>|
|[<span data-ttu-id="0e2fc-1471">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1472">ReadItem</span></span>|
|[<span data-ttu-id="0e2fc-1473">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1474">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1474">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-1475">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1475">Example</span></span>

<span data-ttu-id="0e2fc-p195">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p195">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="0e2fc-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="0e2fc-1480">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1480">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="0e2fc-1481">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1481">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0e2fc-1482">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1482">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="0e2fc-1483">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1483">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0e2fc-1484">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1484">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1485">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1485">Parameters</span></span>

|<span data-ttu-id="0e2fc-1486">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1486">Name</span></span>|<span data-ttu-id="0e2fc-1487">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1487">Type</span></span>|<span data-ttu-id="0e2fc-1488">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1488">Attributes</span></span>|<span data-ttu-id="0e2fc-1489">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1489">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="0e2fc-1490">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1490">String</span></span>||<span data-ttu-id="0e2fc-1491">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1491">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="0e2fc-1492">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1492">Object</span></span>|<span data-ttu-id="0e2fc-1493">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1493">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1494">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1494">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-1495">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1495">Object</span></span>|<span data-ttu-id="0e2fc-1496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1496">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1497">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1497">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1498">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1498">function</span></span>|<span data-ttu-id="0e2fc-1499">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1499">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1500">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e2fc-1501">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1501">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e2fc-1502">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1502">Errors</span></span>

|<span data-ttu-id="0e2fc-1503">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1503">Error code</span></span>|<span data-ttu-id="0e2fc-1504">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1504">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="0e2fc-1505">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1505">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1506">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1506">Requirements</span></span>

|<span data-ttu-id="0e2fc-1507">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1507">Requirement</span></span>|<span data-ttu-id="0e2fc-1508">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1508">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1509">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1510">1.1</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1510">1.1</span></span>|
|[<span data-ttu-id="0e2fc-1511">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1512">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1512">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e2fc-1513">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1514">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1514">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-1515">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1515">Example</span></span>

<span data-ttu-id="0e2fc-1516">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1516">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="0e2fc-1517">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1517">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="0e2fc-1518">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1518">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="0e2fc-1519">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1519">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1520">Parameters</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1520">Parameters</span></span>

| <span data-ttu-id="0e2fc-1521">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1521">Name</span></span> | <span data-ttu-id="0e2fc-1522">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1522">Type</span></span> | <span data-ttu-id="0e2fc-1523">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1523">Attributes</span></span> | <span data-ttu-id="0e2fc-1524">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1524">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0e2fc-1525">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1525">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0e2fc-1526">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1526">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="0e2fc-1527">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1527">Object</span></span> | <span data-ttu-id="0e2fc-1528">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1528">&lt;optional&gt;</span></span> | <span data-ttu-id="0e2fc-1529">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1529">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0e2fc-1530">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1530">Object</span></span> | <span data-ttu-id="0e2fc-1531">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1531">&lt;optional&gt;</span></span> | <span data-ttu-id="0e2fc-1532">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1532">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0e2fc-1533">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1533">function</span></span>| <span data-ttu-id="0e2fc-1534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1534">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1535">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1535">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1536">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1536">Requirements</span></span>

|<span data-ttu-id="0e2fc-1537">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1537">Requirement</span></span>| <span data-ttu-id="0e2fc-1538">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1538">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1539">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e2fc-1540">1.7</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1540">1.7</span></span> |
|[<span data-ttu-id="0e2fc-1541">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e2fc-1542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1542">ReadItem</span></span> |
|[<span data-ttu-id="0e2fc-1543">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0e2fc-1544">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1544">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="0e2fc-1545">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1545">saveAsync([options], callback)</span></span>

<span data-ttu-id="0e2fc-1546">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1546">Asynchronously saves an item.</span></span>

<span data-ttu-id="0e2fc-1547">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1547">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="0e2fc-1548">Dans Outlook sur le web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1548">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="0e2fc-1549">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1549">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1550">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1550">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="0e2fc-1551">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1551">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="0e2fc-p199">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p199">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="0e2fc-1555">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1555">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="0e2fc-1556">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1556">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="0e2fc-1557">La méthode `saveAsync` échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1557">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="0e2fc-1558">Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1558">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="0e2fc-1559">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1559">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1560">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1560">Parameters</span></span>

|<span data-ttu-id="0e2fc-1561">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1561">Name</span></span>|<span data-ttu-id="0e2fc-1562">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1562">Type</span></span>|<span data-ttu-id="0e2fc-1563">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1563">Attributes</span></span>|<span data-ttu-id="0e2fc-1564">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1564">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e2fc-1565">Object</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1565">Object</span></span>|<span data-ttu-id="0e2fc-1566">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1566">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1567">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1567">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-1568">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1568">Object</span></span>|<span data-ttu-id="0e2fc-1569">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1569">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1570">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1570">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1571">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1571">function</span></span>||<span data-ttu-id="0e2fc-1572">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1572">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e2fc-1573">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1573">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1574">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1574">Requirements</span></span>

|<span data-ttu-id="0e2fc-1575">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1575">Requirement</span></span>|<span data-ttu-id="0e2fc-1576">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1576">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1577">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1578">1.3</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1578">1.3</span></span>|
|[<span data-ttu-id="0e2fc-1579">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1580">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1580">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e2fc-1581">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1582">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1582">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e2fc-1583">範例</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1583">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="0e2fc-p201">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p201">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="0e2fc-1586">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1586">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="0e2fc-1587">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1587">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="0e2fc-p202">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p202">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e2fc-1591">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1591">Parameters</span></span>

|<span data-ttu-id="0e2fc-1592">Nom</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1592">Name</span></span>|<span data-ttu-id="0e2fc-1593">Type</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1593">Type</span></span>|<span data-ttu-id="0e2fc-1594">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1594">Attributes</span></span>|<span data-ttu-id="0e2fc-1595">Description</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1595">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="0e2fc-1596">String</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1596">String</span></span>||<span data-ttu-id="0e2fc-p203">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-p203">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="0e2fc-1600">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1600">Object</span></span>|<span data-ttu-id="0e2fc-1601">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1602">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1602">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e2fc-1603">Objet</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1603">Object</span></span>|<span data-ttu-id="0e2fc-1604">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1604">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1605">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1605">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="0e2fc-1606">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1606">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="0e2fc-1607">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1607">&lt;optional&gt;</span></span>|<span data-ttu-id="0e2fc-1608">Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1608">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="0e2fc-1609">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1609">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="0e2fc-1610">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1610">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="0e2fc-1611">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1611">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="0e2fc-1612">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1612">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="0e2fc-1613">fonction</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1613">function</span></span>||<span data-ttu-id="0e2fc-1614">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e2fc-1615">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1615">Requirements</span></span>

|<span data-ttu-id="0e2fc-1616">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1616">Requirement</span></span>|<span data-ttu-id="0e2fc-1617">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1617">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e2fc-1618">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e2fc-1619">1.2</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1619">1.2</span></span>|
|[<span data-ttu-id="0e2fc-1620">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e2fc-1621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1621">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e2fc-1622">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e2fc-1623">Composition</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e2fc-1624">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e2fc-1624">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
