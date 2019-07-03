---
title: Office. Context. Mailbox. Item-Preview ensemble de conditions requises
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 537ac59649b149d9bb54b09f8e16704adb813f58
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454901"
---
# <a name="item"></a><span data-ttu-id="0e073-102">élément</span><span class="sxs-lookup"><span data-stu-id="0e073-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="0e073-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="0e073-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="0e073-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="0e073-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-106">Requirements</span></span>

|<span data-ttu-id="0e073-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-107">Requirement</span></span>|<span data-ttu-id="0e073-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-110">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-110">1.0</span></span>|
|[<span data-ttu-id="0e073-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="0e073-112">Restricted</span></span>|
|[<span data-ttu-id="0e073-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0e073-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="0e073-115">Members and methods</span></span>

| <span data-ttu-id="0e073-116">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-116">Member</span></span> | <span data-ttu-id="0e073-117">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0e073-118">attachments</span><span class="sxs-lookup"><span data-stu-id="0e073-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="0e073-119">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-119">Member</span></span> |
| [<span data-ttu-id="0e073-120">bcc</span><span class="sxs-lookup"><span data-stu-id="0e073-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="0e073-121">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-121">Member</span></span> |
| [<span data-ttu-id="0e073-122">body</span><span class="sxs-lookup"><span data-stu-id="0e073-122">body</span></span>](#body-body) | <span data-ttu-id="0e073-123">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-123">Member</span></span> |
| [<span data-ttu-id="0e073-124">catégories</span><span class="sxs-lookup"><span data-stu-id="0e073-124">categories</span></span>](#categories-categories) | <span data-ttu-id="0e073-125">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-125">Member</span></span> |
| [<span data-ttu-id="0e073-126">cc</span><span class="sxs-lookup"><span data-stu-id="0e073-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0e073-127">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-127">Member</span></span> |
| [<span data-ttu-id="0e073-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="0e073-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="0e073-129">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-129">Member</span></span> |
| [<span data-ttu-id="0e073-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="0e073-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="0e073-131">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-131">Member</span></span> |
| [<span data-ttu-id="0e073-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="0e073-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="0e073-133">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-133">Member</span></span> |
| [<span data-ttu-id="0e073-134">end</span><span class="sxs-lookup"><span data-stu-id="0e073-134">end</span></span>](#end-datetime) | <span data-ttu-id="0e073-135">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-135">Member</span></span> |
| [<span data-ttu-id="0e073-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="0e073-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="0e073-137">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-137">Member</span></span> |
| [<span data-ttu-id="0e073-138">from</span><span class="sxs-lookup"><span data-stu-id="0e073-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="0e073-139">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-139">Member</span></span> |
| [<span data-ttu-id="0e073-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="0e073-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="0e073-141">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-141">Member</span></span> |
| [<span data-ttu-id="0e073-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="0e073-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="0e073-143">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-143">Member</span></span> |
| [<span data-ttu-id="0e073-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="0e073-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="0e073-145">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-145">Member</span></span> |
| [<span data-ttu-id="0e073-146">itemId</span><span class="sxs-lookup"><span data-stu-id="0e073-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="0e073-147">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-147">Member</span></span> |
| [<span data-ttu-id="0e073-148">itemType</span><span class="sxs-lookup"><span data-stu-id="0e073-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="0e073-149">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-149">Member</span></span> |
| [<span data-ttu-id="0e073-150">location</span><span class="sxs-lookup"><span data-stu-id="0e073-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="0e073-151">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-151">Member</span></span> |
| [<span data-ttu-id="0e073-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="0e073-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="0e073-153">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-153">Member</span></span> |
| [<span data-ttu-id="0e073-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="0e073-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="0e073-155">Member</span><span class="sxs-lookup"><span data-stu-id="0e073-155">Member</span></span> |
| [<span data-ttu-id="0e073-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="0e073-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0e073-157">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-157">Member</span></span> |
| [<span data-ttu-id="0e073-158">organizer</span><span class="sxs-lookup"><span data-stu-id="0e073-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="0e073-159">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-159">Member</span></span> |
| [<span data-ttu-id="0e073-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="0e073-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="0e073-161">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-161">Member</span></span> |
| [<span data-ttu-id="0e073-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="0e073-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0e073-163">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-163">Member</span></span> |
| [<span data-ttu-id="0e073-164">sender</span><span class="sxs-lookup"><span data-stu-id="0e073-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="0e073-165">Member</span><span class="sxs-lookup"><span data-stu-id="0e073-165">Member</span></span> |
| [<span data-ttu-id="0e073-166">seriesId</span><span class="sxs-lookup"><span data-stu-id="0e073-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="0e073-167">Member</span><span class="sxs-lookup"><span data-stu-id="0e073-167">Member</span></span> |
| [<span data-ttu-id="0e073-168">start</span><span class="sxs-lookup"><span data-stu-id="0e073-168">start</span></span>](#start-datetime) | <span data-ttu-id="0e073-169">Member</span><span class="sxs-lookup"><span data-stu-id="0e073-169">Member</span></span> |
| [<span data-ttu-id="0e073-170">subject</span><span class="sxs-lookup"><span data-stu-id="0e073-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="0e073-171">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-171">Member</span></span> |
| [<span data-ttu-id="0e073-172">to</span><span class="sxs-lookup"><span data-stu-id="0e073-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0e073-173">Membre</span><span class="sxs-lookup"><span data-stu-id="0e073-173">Member</span></span> |
| [<span data-ttu-id="0e073-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="0e073-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-175">Method</span></span> |
| [<span data-ttu-id="0e073-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="0e073-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="0e073-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-177">Method</span></span> |
| [<span data-ttu-id="0e073-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="0e073-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-179">Method</span></span> |
| [<span data-ttu-id="0e073-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="0e073-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-181">Method</span></span> |
| [<span data-ttu-id="0e073-182">close</span><span class="sxs-lookup"><span data-stu-id="0e073-182">close</span></span>](#close) | <span data-ttu-id="0e073-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-183">Method</span></span> |
| [<span data-ttu-id="0e073-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="0e073-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="0e073-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-185">Method</span></span> |
| [<span data-ttu-id="0e073-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="0e073-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="0e073-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-187">Method</span></span> |
| [<span data-ttu-id="0e073-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="0e073-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-189">Method</span></span> |
| [<span data-ttu-id="0e073-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="0e073-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-191">Method</span></span> |
| [<span data-ttu-id="0e073-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="0e073-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="0e073-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-193">Method</span></span> |
| [<span data-ttu-id="0e073-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="0e073-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0e073-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-195">Method</span></span> |
| [<span data-ttu-id="0e073-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="0e073-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0e073-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-197">Method</span></span> |
| [<span data-ttu-id="0e073-198">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="0e073-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-199">Method</span></span> |
| [<span data-ttu-id="0e073-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="0e073-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-201">Method</span></span> |
| [<span data-ttu-id="0e073-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0e073-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="0e073-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-203">Method</span></span> |
| [<span data-ttu-id="0e073-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="0e073-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="0e073-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-205">Method</span></span> |
| [<span data-ttu-id="0e073-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="0e073-207">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-207">Method</span></span> |
| [<span data-ttu-id="0e073-208">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="0e073-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="0e073-209">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-209">Method</span></span> |
| [<span data-ttu-id="0e073-210">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0e073-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="0e073-211">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-211">Method</span></span> |
| [<span data-ttu-id="0e073-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="0e073-213">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-213">Method</span></span> |
| [<span data-ttu-id="0e073-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="0e073-215">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-215">Method</span></span> |
| [<span data-ttu-id="0e073-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="0e073-217">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-217">Method</span></span> |
| [<span data-ttu-id="0e073-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="0e073-219">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-219">Method</span></span> |
| [<span data-ttu-id="0e073-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="0e073-221">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-221">Method</span></span> |
| [<span data-ttu-id="0e073-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0e073-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="0e073-223">Méthode</span><span class="sxs-lookup"><span data-stu-id="0e073-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="0e073-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-224">Example</span></span>

<span data-ttu-id="0e073-225">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e073-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="0e073-226">Membres</span><span class="sxs-lookup"><span data-stu-id="0e073-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="0e073-227">pièces jointes: tableau. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0e073-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="0e073-228">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="0e073-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="0e073-229">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-230">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="0e073-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="0e073-231">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="0e073-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-232">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-232">Type</span></span>

*   <span data-ttu-id="0e073-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0e073-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-234">Requirements</span></span>

|<span data-ttu-id="0e073-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-235">Requirement</span></span>|<span data-ttu-id="0e073-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-238">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-238">1.0</span></span>|
|[<span data-ttu-id="0e073-239">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-240">ReadItem</span></span>|
|[<span data-ttu-id="0e073-241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-242">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-243">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-243">Example</span></span>

<span data-ttu-id="0e073-244">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e073-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0e073-245">CCI: [destinataires](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e073-245">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0e073-246">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e073-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="0e073-247">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-247">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-248">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-248">Type</span></span>

*   [<span data-ttu-id="0e073-249">Destinataires</span><span class="sxs-lookup"><span data-stu-id="0e073-249">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="0e073-250">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-250">Requirements</span></span>

|<span data-ttu-id="0e073-251">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-251">Requirement</span></span>|<span data-ttu-id="0e073-252">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-253">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-254">1.1</span><span class="sxs-lookup"><span data-stu-id="0e073-254">1.1</span></span>|
|[<span data-ttu-id="0e073-255">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-256">ReadItem</span></span>|
|[<span data-ttu-id="0e073-257">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-258">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-258">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-259">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-259">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

---
---

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="0e073-260">Body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="0e073-260">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="0e073-261">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-261">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-262">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-262">Type</span></span>

*   [<span data-ttu-id="0e073-263">Body</span><span class="sxs-lookup"><span data-stu-id="0e073-263">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="0e073-264">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-264">Requirements</span></span>

|<span data-ttu-id="0e073-265">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-265">Requirement</span></span>|<span data-ttu-id="0e073-266">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-267">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-268">1.1</span><span class="sxs-lookup"><span data-stu-id="0e073-268">1.1</span></span>|
|[<span data-ttu-id="0e073-269">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-270">ReadItem</span></span>|
|[<span data-ttu-id="0e073-271">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-272">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-273">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-273">Example</span></span>

<span data-ttu-id="0e073-274">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="0e073-274">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="0e073-275">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-275">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="0e073-276">Catégories: [catégories](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="0e073-276">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="0e073-277">Obtient un objet qui fournit des méthodes pour la gestion des catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-277">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-278">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-278">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-279">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-279">Type</span></span>

*   [<span data-ttu-id="0e073-280">Catégories</span><span class="sxs-lookup"><span data-stu-id="0e073-280">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="0e073-281">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-281">Requirements</span></span>

|<span data-ttu-id="0e073-282">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-282">Requirement</span></span>|<span data-ttu-id="0e073-283">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-284">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-285">Aperçu</span><span class="sxs-lookup"><span data-stu-id="0e073-285">Preview</span></span>|
|[<span data-ttu-id="0e073-286">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-287">ReadItem</span></span>|
|[<span data-ttu-id="0e073-288">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-289">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-289">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-290">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-290">Example</span></span>

<span data-ttu-id="0e073-291">Cet exemple obtient les catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-291">This example gets the item's categories.</span></span>

```javascript
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0e073-292">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e073-292">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0e073-293">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e073-293">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="0e073-294">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e073-294">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-295">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-295">Read mode</span></span>

<span data-ttu-id="0e073-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e073-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-298">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-298">Compose mode</span></span>

<span data-ttu-id="0e073-299">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="0e073-299">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e073-300">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-300">Type</span></span>

*   <span data-ttu-id="0e073-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e073-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-302">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-302">Requirements</span></span>

|<span data-ttu-id="0e073-303">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-303">Requirement</span></span>|<span data-ttu-id="0e073-304">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-305">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-306">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-306">1.0</span></span>|
|[<span data-ttu-id="0e073-307">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-308">ReadItem</span></span>|
|[<span data-ttu-id="0e073-309">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-310">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-310">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="0e073-311">(Nullable) conversationId: chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-311">(nullable) conversationId: String</span></span>

<span data-ttu-id="0e073-312">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="0e073-312">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="0e073-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="0e073-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="0e073-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="0e073-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-317">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-317">Type</span></span>

*   <span data-ttu-id="0e073-318">String</span><span class="sxs-lookup"><span data-stu-id="0e073-318">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-319">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-319">Requirements</span></span>

|<span data-ttu-id="0e073-320">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-320">Requirement</span></span>|<span data-ttu-id="0e073-321">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-321">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-322">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-322">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-323">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-323">1.0</span></span>|
|[<span data-ttu-id="0e073-324">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-324">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-325">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-325">ReadItem</span></span>|
|[<span data-ttu-id="0e073-326">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-326">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-327">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-327">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-328">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-328">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="0e073-329">dateTimeCreated: date</span><span class="sxs-lookup"><span data-stu-id="0e073-329">dateTimeCreated: Date</span></span>

<span data-ttu-id="0e073-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-332">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-332">Type</span></span>

*   <span data-ttu-id="0e073-333">Date</span><span class="sxs-lookup"><span data-stu-id="0e073-333">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-334">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-334">Requirements</span></span>

|<span data-ttu-id="0e073-335">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-335">Requirement</span></span>|<span data-ttu-id="0e073-336">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-337">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-338">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-338">1.0</span></span>|
|[<span data-ttu-id="0e073-339">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-340">ReadItem</span></span>|
|[<span data-ttu-id="0e073-341">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-342">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-343">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-343">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="0e073-344">dateTimeModified: date</span><span class="sxs-lookup"><span data-stu-id="0e073-344">dateTimeModified: Date</span></span>

<span data-ttu-id="0e073-345">Obtient la date et l’heure de la dernière modification d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-345">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="0e073-346">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-346">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-347">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-347">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-348">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-348">Type</span></span>

*   <span data-ttu-id="0e073-349">Date</span><span class="sxs-lookup"><span data-stu-id="0e073-349">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-350">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-350">Requirements</span></span>

|<span data-ttu-id="0e073-351">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-351">Requirement</span></span>|<span data-ttu-id="0e073-352">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-353">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-354">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-354">1.0</span></span>|
|[<span data-ttu-id="0e073-355">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-356">ReadItem</span></span>|
|[<span data-ttu-id="0e073-357">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-358">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-358">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-359">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-359">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="0e073-360">fin: date | [Fois](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="0e073-360">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="0e073-361">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-361">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="0e073-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="0e073-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-364">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-364">Read mode</span></span>

<span data-ttu-id="0e073-365">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="0e073-365">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-366">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-366">Compose mode</span></span>

<span data-ttu-id="0e073-367">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e073-367">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="0e073-368">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e073-368">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0e073-369">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e073-369">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="0e073-370">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-370">Type</span></span>

*   <span data-ttu-id="0e073-371">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="0e073-371">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-372">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-372">Requirements</span></span>

|<span data-ttu-id="0e073-373">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-373">Requirement</span></span>|<span data-ttu-id="0e073-374">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-375">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-376">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-376">1.0</span></span>|
|[<span data-ttu-id="0e073-377">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-378">ReadItem</span></span>|
|[<span data-ttu-id="0e073-379">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-380">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-380">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="0e073-381">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="0e073-381">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="0e073-382">Obtient ou définit les emplacements d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-382">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-383">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-383">Read mode</span></span>

<span data-ttu-id="0e073-384">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui vous permet d’obtenir l’ensemble des emplacements (chacun représenté par un objet [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associé au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0e073-385">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-385">Compose mode</span></span>

<span data-ttu-id="0e073-386">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui fournit des méthodes pour obtenir, supprimer ou ajouter des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-386">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-387">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-387">Type</span></span>

*   [<span data-ttu-id="0e073-388">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="0e073-388">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="0e073-389">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-389">Requirements</span></span>

|<span data-ttu-id="0e073-390">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-390">Requirement</span></span>|<span data-ttu-id="0e073-391">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-392">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-393">Aperçu</span><span class="sxs-lookup"><span data-stu-id="0e073-393">Preview</span></span>|
|[<span data-ttu-id="0e073-394">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-395">ReadItem</span></span>|
|[<span data-ttu-id="0e073-396">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-397">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-397">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-398">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-398">Example</span></span>

<span data-ttu-id="0e073-399">L’exemple suivant obtient les emplacements actuels associés au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-399">The following example gets the current locations associated with the appointment.</span></span>

```javascript
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

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="0e073-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="0e073-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="0e073-401">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e073-401">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="0e073-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="0e073-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-404">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0e073-404">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-405">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-405">Read mode</span></span>

<span data-ttu-id="0e073-406">La `from` propriété renvoie un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="0e073-406">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-407">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-407">Compose mode</span></span>

<span data-ttu-id="0e073-408">La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.</span><span class="sxs-lookup"><span data-stu-id="0e073-408">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e073-409">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-409">Type</span></span>

*   <span data-ttu-id="0e073-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [à partir de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="0e073-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-411">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-411">Requirements</span></span>

|<span data-ttu-id="0e073-412">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-412">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="0e073-413">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-414">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-414">1.0</span></span>|<span data-ttu-id="0e073-415">1.7</span><span class="sxs-lookup"><span data-stu-id="0e073-415">1.7</span></span>|
|[<span data-ttu-id="0e073-416">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-417">ReadItem</span></span>|<span data-ttu-id="0e073-418">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e073-418">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e073-419">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-419">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-420">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-420">Read</span></span>|<span data-ttu-id="0e073-421">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-421">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="0e073-422">internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="0e073-422">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="0e073-423">Obtient ou définit les en-têtes Internet personnalisés d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e073-423">Gets or sets custom internet headers on a message.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-424">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-424">Type</span></span>

*   [<span data-ttu-id="0e073-425">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="0e073-425">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="0e073-426">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-426">Requirements</span></span>

|<span data-ttu-id="0e073-427">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-427">Requirement</span></span>|<span data-ttu-id="0e073-428">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-429">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-430">Aperçu</span><span class="sxs-lookup"><span data-stu-id="0e073-430">Preview</span></span>|
|[<span data-ttu-id="0e073-431">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-432">ReadItem</span></span>|
|[<span data-ttu-id="0e073-433">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-434">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-435">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-435">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="0e073-436">internetMessageId: chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-436">internetMessageId: String</span></span>

<span data-ttu-id="0e073-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-439">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-439">Type</span></span>

*   <span data-ttu-id="0e073-440">String</span><span class="sxs-lookup"><span data-stu-id="0e073-440">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-441">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-441">Requirements</span></span>

|<span data-ttu-id="0e073-442">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-442">Requirement</span></span>|<span data-ttu-id="0e073-443">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-444">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-445">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-445">1.0</span></span>|
|[<span data-ttu-id="0e073-446">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-447">ReadItem</span></span>|
|[<span data-ttu-id="0e073-448">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-449">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-449">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-450">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-450">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="0e073-451">itemClass: chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-451">itemClass: String</span></span>

<span data-ttu-id="0e073-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="0e073-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="0e073-456">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-456">Type</span></span>|<span data-ttu-id="0e073-457">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-457">Description</span></span>|<span data-ttu-id="0e073-458">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="0e073-458">item class</span></span>|
|---|---|---|
|<span data-ttu-id="0e073-459">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="0e073-459">Appointment items</span></span>|<span data-ttu-id="0e073-460">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="0e073-460">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="0e073-461">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="0e073-461">Message items</span></span>|<span data-ttu-id="0e073-462">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="0e073-462">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="0e073-463">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="0e073-463">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-464">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-464">Type</span></span>

*   <span data-ttu-id="0e073-465">String</span><span class="sxs-lookup"><span data-stu-id="0e073-465">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-466">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-466">Requirements</span></span>

|<span data-ttu-id="0e073-467">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-467">Requirement</span></span>|<span data-ttu-id="0e073-468">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-469">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-470">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-470">1.0</span></span>|
|[<span data-ttu-id="0e073-471">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-472">ReadItem</span></span>|
|[<span data-ttu-id="0e073-473">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-474">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-474">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-475">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-475">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="0e073-476">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="0e073-476">(nullable) itemId: String</span></span>

<span data-ttu-id="0e073-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-479">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="0e073-479">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0e073-480">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e073-480">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="0e073-481">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="0e073-481">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0e073-482">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="0e073-482">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="0e073-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-485">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-485">Type</span></span>

*   <span data-ttu-id="0e073-486">String</span><span class="sxs-lookup"><span data-stu-id="0e073-486">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-487">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-487">Requirements</span></span>

|<span data-ttu-id="0e073-488">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-488">Requirement</span></span>|<span data-ttu-id="0e073-489">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-490">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-490">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-491">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-491">1.0</span></span>|
|[<span data-ttu-id="0e073-492">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-492">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-493">ReadItem</span></span>|
|[<span data-ttu-id="0e073-494">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-494">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-495">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-495">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-496">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-496">Example</span></span>

<span data-ttu-id="0e073-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0e073-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="0e073-499">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="0e073-499">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="0e073-500">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="0e073-500">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="0e073-501">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-501">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-502">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-502">Type</span></span>

*   [<span data-ttu-id="0e073-503">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="0e073-503">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="0e073-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-504">Requirements</span></span>

|<span data-ttu-id="0e073-505">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-505">Requirement</span></span>|<span data-ttu-id="0e073-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-507">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-508">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-508">1.0</span></span>|
|[<span data-ttu-id="0e073-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-510">ReadItem</span></span>|
|[<span data-ttu-id="0e073-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-512">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-512">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-513">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-513">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="0e073-514">Location: String | [Emplacement](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="0e073-514">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="0e073-515">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-515">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-516">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-516">Read mode</span></span>

<span data-ttu-id="0e073-517">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-517">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-518">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-518">Compose mode</span></span>

<span data-ttu-id="0e073-519">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-519">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e073-520">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-520">Type</span></span>

*   <span data-ttu-id="0e073-521">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="0e073-521">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-522">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-522">Requirements</span></span>

|<span data-ttu-id="0e073-523">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-523">Requirement</span></span>|<span data-ttu-id="0e073-524">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-525">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-526">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-526">1.0</span></span>|
|[<span data-ttu-id="0e073-527">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-528">ReadItem</span></span>|
|[<span data-ttu-id="0e073-529">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-530">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-530">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="0e073-531">normalizedSubject: chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-531">normalizedSubject: String</span></span>

<span data-ttu-id="0e073-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="0e073-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="0e073-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-536">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-536">Type</span></span>

*   <span data-ttu-id="0e073-537">String</span><span class="sxs-lookup"><span data-stu-id="0e073-537">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-538">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-538">Requirements</span></span>

|<span data-ttu-id="0e073-539">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-539">Requirement</span></span>|<span data-ttu-id="0e073-540">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-541">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-542">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-542">1.0</span></span>|
|[<span data-ttu-id="0e073-543">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-544">ReadItem</span></span>|
|[<span data-ttu-id="0e073-545">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-546">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-547">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-547">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="0e073-548">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="0e073-548">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="0e073-549">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-549">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-550">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-550">Type</span></span>

*   [<span data-ttu-id="0e073-551">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="0e073-551">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="0e073-552">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-552">Requirements</span></span>

|<span data-ttu-id="0e073-553">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-553">Requirement</span></span>|<span data-ttu-id="0e073-554">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-554">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-555">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-555">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-556">1.3</span><span class="sxs-lookup"><span data-stu-id="0e073-556">1.3</span></span>|
|[<span data-ttu-id="0e073-557">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-557">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-558">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-558">ReadItem</span></span>|
|[<span data-ttu-id="0e073-559">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-559">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-560">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-560">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-561">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-561">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0e073-562">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="0e073-562">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0e073-563">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="0e073-563">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="0e073-564">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e073-564">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-565">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-565">Read mode</span></span>

<span data-ttu-id="0e073-566">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-566">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-567">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-567">Compose mode</span></span>

<span data-ttu-id="0e073-568">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-568">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e073-569">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-569">Type</span></span>

*   <span data-ttu-id="0e073-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e073-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-571">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-571">Requirements</span></span>

|<span data-ttu-id="0e073-572">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-572">Requirement</span></span>|<span data-ttu-id="0e073-573">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-573">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-574">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-575">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-575">1.0</span></span>|
|[<span data-ttu-id="0e073-576">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-577">ReadItem</span></span>|
|[<span data-ttu-id="0e073-578">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-579">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-579">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="0e073-580">Organisateur: [](/javascript/api/outlook/office.emailaddressdetails)|[organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0e073-580">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="0e073-581">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="0e073-581">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-582">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-582">Read mode</span></span>

<span data-ttu-id="0e073-583">La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-583">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-584">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-584">Compose mode</span></span>

<span data-ttu-id="0e073-585">La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="0e073-585">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="0e073-586">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-586">Type</span></span>

*   <span data-ttu-id="0e073-587">[](/javascript/api/outlook/office.emailaddressdetails) | [Organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0e073-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-588">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-588">Requirements</span></span>

|<span data-ttu-id="0e073-589">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-589">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="0e073-590">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-591">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-591">1.0</span></span>|<span data-ttu-id="0e073-592">1.7</span><span class="sxs-lookup"><span data-stu-id="0e073-592">1.7</span></span>|
|[<span data-ttu-id="0e073-593">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-594">ReadItem</span></span>|<span data-ttu-id="0e073-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e073-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e073-596">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-597">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-597">Read</span></span>|<span data-ttu-id="0e073-598">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-598">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="0e073-599">(Nullable) récurrence: [périodicité](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="0e073-599">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="0e073-600">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-600">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="0e073-601">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-601">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="0e073-602">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-602">Read and compose modes for appointment items.</span></span> <span data-ttu-id="0e073-603">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-603">Read mode for meeting request items.</span></span>

<span data-ttu-id="0e073-604">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="0e073-604">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="0e073-605">`null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="0e073-605">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="0e073-606">`undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-606">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="0e073-607">Remarque: les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="0e073-607">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="0e073-608">Remarque: si l’objet de périodicité `null`est, cela indique que l’objet est un rendez-vous unique ou une demande de réunion d’un seul rendez-vous et non d’une série.</span><span class="sxs-lookup"><span data-stu-id="0e073-608">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-609">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-609">Read mode</span></span>

<span data-ttu-id="0e073-610">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-610">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="0e073-611">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-611">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-612">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-612">Compose mode</span></span>

<span data-ttu-id="0e073-613">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-613">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="0e073-614">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-614">This is available for appointments.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="0e073-615">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-615">Type</span></span>

* [<span data-ttu-id="0e073-616">Instances</span><span class="sxs-lookup"><span data-stu-id="0e073-616">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="0e073-617">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-617">Requirement</span></span>|<span data-ttu-id="0e073-618">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-619">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-620">1.7</span><span class="sxs-lookup"><span data-stu-id="0e073-620">1.7</span></span>|
|[<span data-ttu-id="0e073-621">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-622">ReadItem</span></span>|
|[<span data-ttu-id="0e073-623">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-624">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-624">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0e073-625">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="0e073-625">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0e073-626">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="0e073-626">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="0e073-627">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e073-627">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-628">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-628">Read mode</span></span>

<span data-ttu-id="0e073-629">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-629">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-630">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-630">Compose mode</span></span>

<span data-ttu-id="0e073-631">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-631">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="0e073-632">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-632">Type</span></span>

*   <span data-ttu-id="0e073-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e073-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-634">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-634">Requirements</span></span>

|<span data-ttu-id="0e073-635">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-635">Requirement</span></span>|<span data-ttu-id="0e073-636">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-637">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-638">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-638">1.0</span></span>|
|[<span data-ttu-id="0e073-639">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-640">ReadItem</span></span>|
|[<span data-ttu-id="0e073-641">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-642">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-642">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="0e073-643">expéditeur: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="0e073-643">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="0e073-p128">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="0e073-p129">Les propriétés [`from`](#from-emailaddressdetailsfrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="0e073-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-648">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0e073-648">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-649">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-649">Type</span></span>

*   [<span data-ttu-id="0e073-650">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0e073-650">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="0e073-651">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-651">Requirements</span></span>

|<span data-ttu-id="0e073-652">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-652">Requirement</span></span>|<span data-ttu-id="0e073-653">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-654">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-655">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-655">1.0</span></span>|
|[<span data-ttu-id="0e073-656">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-657">ReadItem</span></span>|
|[<span data-ttu-id="0e073-658">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-659">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-659">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-660">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-660">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="0e073-661">(Nullable) seriesId: chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-661">(nullable) seriesId: String</span></span>

<span data-ttu-id="0e073-662">Obtient l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="0e073-662">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="0e073-663">Dans Outlook sur le Web et les clients de bureau `seriesId` , le renvoie l’ID des services Web Exchange (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="0e073-663">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="0e073-664">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="0e073-664">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-665">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="0e073-665">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0e073-666">La `seriesId` propriété n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e073-666">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="0e073-667">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="0e073-667">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0e073-668">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="0e073-668">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="0e073-669">La `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-669">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="0e073-670">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-670">Type</span></span>

* <span data-ttu-id="0e073-671">String</span><span class="sxs-lookup"><span data-stu-id="0e073-671">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-672">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-672">Requirements</span></span>

|<span data-ttu-id="0e073-673">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-673">Requirement</span></span>|<span data-ttu-id="0e073-674">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-674">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-675">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-675">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-676">1.7</span><span class="sxs-lookup"><span data-stu-id="0e073-676">1.7</span></span>|
|[<span data-ttu-id="0e073-677">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-677">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-678">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-678">ReadItem</span></span>|
|[<span data-ttu-id="0e073-679">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-679">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-680">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-680">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-681">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-681">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="0e073-682">début: date | [Fois](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="0e073-682">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="0e073-683">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-683">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="0e073-p132">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="0e073-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-686">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-686">Read mode</span></span>

<span data-ttu-id="0e073-687">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="0e073-687">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-688">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-688">Compose mode</span></span>

<span data-ttu-id="0e073-689">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e073-689">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="0e073-690">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e073-690">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0e073-691">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e073-691">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="0e073-692">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-692">Type</span></span>

*   <span data-ttu-id="0e073-693">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="0e073-693">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-694">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-694">Requirements</span></span>

|<span data-ttu-id="0e073-695">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-695">Requirement</span></span>|<span data-ttu-id="0e073-696">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-697">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-698">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-698">1.0</span></span>|
|[<span data-ttu-id="0e073-699">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-700">ReadItem</span></span>|
|[<span data-ttu-id="0e073-701">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-702">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-702">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="0e073-703">Subject: String | [Objet](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="0e073-703">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="0e073-704">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-704">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="0e073-705">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="0e073-705">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-706">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-706">Read mode</span></span>

<span data-ttu-id="0e073-p133">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="0e073-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="0e073-709">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e073-709">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-710">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-710">Compose mode</span></span>
<span data-ttu-id="0e073-711">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="0e073-711">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="0e073-712">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-712">Type</span></span>

*   <span data-ttu-id="0e073-713">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="0e073-713">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-714">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-714">Requirements</span></span>

|<span data-ttu-id="0e073-715">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-715">Requirement</span></span>|<span data-ttu-id="0e073-716">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-716">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-717">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-717">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-718">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-718">1.0</span></span>|
|[<span data-ttu-id="0e073-719">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-719">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-720">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-720">ReadItem</span></span>|
|[<span data-ttu-id="0e073-721">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-721">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-722">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-722">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0e073-723">to: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e073-723">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0e073-724">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e073-724">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="0e073-725">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e073-725">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e073-726">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-726">Read mode</span></span>

<span data-ttu-id="0e073-p135">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0e073-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="0e073-729">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0e073-729">Compose mode</span></span>

<span data-ttu-id="0e073-730">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="0e073-730">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e073-731">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-731">Type</span></span>

*   <span data-ttu-id="0e073-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e073-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-733">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-733">Requirements</span></span>

|<span data-ttu-id="0e073-734">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-734">Requirement</span></span>|<span data-ttu-id="0e073-735">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-735">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-736">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-736">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-737">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-737">1.0</span></span>|
|[<span data-ttu-id="0e073-738">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-738">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-739">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-739">ReadItem</span></span>|
|[<span data-ttu-id="0e073-740">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-740">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-741">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-741">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="0e073-742">Méthodes</span><span class="sxs-lookup"><span data-stu-id="0e073-742">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="0e073-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e073-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0e073-744">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="0e073-744">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0e073-745">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="0e073-745">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="0e073-746">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0e073-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-747">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-747">Parameters</span></span>
|<span data-ttu-id="0e073-748">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-748">Name</span></span>|<span data-ttu-id="0e073-749">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-749">Type</span></span>|<span data-ttu-id="0e073-750">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-750">Attributes</span></span>|<span data-ttu-id="0e073-751">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-751">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="0e073-752">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-752">String</span></span>||<span data-ttu-id="0e073-p136">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e073-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="0e073-755">String</span><span class="sxs-lookup"><span data-stu-id="0e073-755">String</span></span>||<span data-ttu-id="0e073-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e073-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0e073-758">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-758">Object</span></span>|<span data-ttu-id="0e073-759">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-759">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-760">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-761">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-761">Object</span></span>|<span data-ttu-id="0e073-762">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-762">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-763">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="0e073-764">Boolean</span><span class="sxs-lookup"><span data-stu-id="0e073-764">Boolean</span></span>|<span data-ttu-id="0e073-765">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-765">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-766">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e073-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="0e073-767">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-767">function</span></span>|<span data-ttu-id="0e073-768">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-768">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-769">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e073-770">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e073-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0e073-771">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="0e073-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e073-772">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e073-772">Errors</span></span>

|<span data-ttu-id="0e073-773">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e073-773">Error code</span></span>|<span data-ttu-id="0e073-774">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="0e073-775">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="0e073-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="0e073-776">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="0e073-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0e073-777">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e073-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-778">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-778">Requirements</span></span>

|<span data-ttu-id="0e073-779">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-779">Requirement</span></span>|<span data-ttu-id="0e073-780">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-781">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-782">1.1</span><span class="sxs-lookup"><span data-stu-id="0e073-782">1.1</span></span>|
|[<span data-ttu-id="0e073-783">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e073-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e073-785">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-786">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e073-787">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e073-787">Examples</span></span>

```javascript
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

<span data-ttu-id="0e073-788">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="0e073-788">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```javascript
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

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="0e073-789">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e073-789">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0e073-790">Ajoute un fichier à partir du codage Base64 à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="0e073-790">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0e073-791">La `addFileAttachmentFromBase64Async` méthode charge le fichier à partir du codage Base64 et l’associe à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="0e073-791">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="0e073-792">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="0e073-792">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="0e073-793">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0e073-793">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-794">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-794">Parameters</span></span>

|<span data-ttu-id="0e073-795">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-795">Name</span></span>|<span data-ttu-id="0e073-796">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-796">Type</span></span>|<span data-ttu-id="0e073-797">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-797">Attributes</span></span>|<span data-ttu-id="0e073-798">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-798">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="0e073-799">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-799">String</span></span>||<span data-ttu-id="0e073-800">Contenu encodé en base64 d’une image ou d’un fichier à ajouter à un message électronique ou à un événement.</span><span class="sxs-lookup"><span data-stu-id="0e073-800">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="0e073-801">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-801">String</span></span>||<span data-ttu-id="0e073-p139">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e073-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0e073-804">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-804">Object</span></span>|<span data-ttu-id="0e073-805">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-805">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-806">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-806">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-807">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-807">Object</span></span>|<span data-ttu-id="0e073-808">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-808">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-809">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-809">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="0e073-810">Boolean</span><span class="sxs-lookup"><span data-stu-id="0e073-810">Boolean</span></span>|<span data-ttu-id="0e073-811">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-811">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-812">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e073-812">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="0e073-813">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-813">function</span></span>|<span data-ttu-id="0e073-814">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-814">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-815">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-815">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e073-816">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e073-816">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0e073-817">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="0e073-817">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e073-818">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e073-818">Errors</span></span>

|<span data-ttu-id="0e073-819">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e073-819">Error code</span></span>|<span data-ttu-id="0e073-820">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-820">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="0e073-821">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="0e073-821">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="0e073-822">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="0e073-822">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0e073-823">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e073-823">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-824">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-824">Requirements</span></span>

|<span data-ttu-id="0e073-825">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-825">Requirement</span></span>|<span data-ttu-id="0e073-826">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-827">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-828">Aperçu</span><span class="sxs-lookup"><span data-stu-id="0e073-828">Preview</span></span>|
|[<span data-ttu-id="0e073-829">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-830">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e073-830">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e073-831">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-832">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-832">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e073-833">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e073-833">Examples</span></span>

```javascript
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

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="0e073-834">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e073-834">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="0e073-835">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="0e073-835">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="0e073-836">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="0e073-836">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-837">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-837">Parameters</span></span>

| <span data-ttu-id="0e073-838">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-838">Name</span></span> | <span data-ttu-id="0e073-839">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-839">Type</span></span> | <span data-ttu-id="0e073-840">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-840">Attributes</span></span> | <span data-ttu-id="0e073-841">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-841">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0e073-842">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0e073-842">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0e073-843">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="0e073-843">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="0e073-844">Fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-844">Function</span></span> || <span data-ttu-id="0e073-p140">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="0e073-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="0e073-848">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-848">Object</span></span> | <span data-ttu-id="0e073-849">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-849">&lt;optional&gt;</span></span> | <span data-ttu-id="0e073-850">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-850">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0e073-851">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-851">Object</span></span> | <span data-ttu-id="0e073-852">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-852">&lt;optional&gt;</span></span> | <span data-ttu-id="0e073-853">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-853">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0e073-854">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-854">function</span></span>| <span data-ttu-id="0e073-855">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-855">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-856">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-857">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-857">Requirements</span></span>

|<span data-ttu-id="0e073-858">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-858">Requirement</span></span>| <span data-ttu-id="0e073-859">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-859">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-860">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e073-861">1.7</span><span class="sxs-lookup"><span data-stu-id="0e073-861">1.7</span></span> |
|[<span data-ttu-id="0e073-862">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e073-863">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-863">ReadItem</span></span> |
|[<span data-ttu-id="0e073-864">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0e073-865">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-865">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="0e073-866">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-866">Example</span></span>

```javascript
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="0e073-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e073-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0e073-868">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-868">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="0e073-p141">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="0e073-872">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0e073-872">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="0e073-873">Si votre complément Office est en cours d’exécution dans Outlook sur le Web, `addItemAttachmentAsync` la méthode peut joindre des éléments à des éléments autres que l’élément que vous modifiez; Toutefois, cette option n’est pas prise en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="0e073-873">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-874">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-874">Parameters</span></span>

|<span data-ttu-id="0e073-875">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-875">Name</span></span>|<span data-ttu-id="0e073-876">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-876">Type</span></span>|<span data-ttu-id="0e073-877">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-877">Attributes</span></span>|<span data-ttu-id="0e073-878">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-878">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="0e073-879">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-879">String</span></span>||<span data-ttu-id="0e073-p142">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e073-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="0e073-882">String</span><span class="sxs-lookup"><span data-stu-id="0e073-882">String</span></span>||<span data-ttu-id="0e073-883">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="0e073-883">The subject of the item to be attached.</span></span> <span data-ttu-id="0e073-884">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e073-884">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0e073-885">Object</span><span class="sxs-lookup"><span data-stu-id="0e073-885">Object</span></span>|<span data-ttu-id="0e073-886">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-886">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-887">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-887">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-888">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-888">Object</span></span>|<span data-ttu-id="0e073-889">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-889">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-890">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-890">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e073-891">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-891">function</span></span>|<span data-ttu-id="0e073-892">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-892">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-893">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-893">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e073-894">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e073-894">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0e073-895">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="0e073-895">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e073-896">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e073-896">Errors</span></span>

|<span data-ttu-id="0e073-897">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e073-897">Error code</span></span>|<span data-ttu-id="0e073-898">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-898">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0e073-899">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e073-899">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-900">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-900">Requirements</span></span>

|<span data-ttu-id="0e073-901">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-901">Requirement</span></span>|<span data-ttu-id="0e073-902">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-903">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-904">1.1</span><span class="sxs-lookup"><span data-stu-id="0e073-904">1.1</span></span>|
|[<span data-ttu-id="0e073-905">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e073-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e073-907">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-908">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-908">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-909">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-909">Example</span></span>

<span data-ttu-id="0e073-910">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="0e073-910">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

---
---

#### <a name="close"></a><span data-ttu-id="0e073-911">close()</span><span class="sxs-lookup"><span data-stu-id="0e073-911">close()</span></span>

<span data-ttu-id="0e073-912">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="0e073-912">Closes the current item that is being composed.</span></span>

<span data-ttu-id="0e073-p144">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="0e073-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-915">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-915">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="0e073-916">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="0e073-916">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-917">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-917">Requirements</span></span>

|<span data-ttu-id="0e073-918">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-918">Requirement</span></span>|<span data-ttu-id="0e073-919">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-919">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-920">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-920">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-921">1.3</span><span class="sxs-lookup"><span data-stu-id="0e073-921">1.3</span></span>|
|[<span data-ttu-id="0e073-922">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-922">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-923">Restreinte</span><span class="sxs-lookup"><span data-stu-id="0e073-923">Restricted</span></span>|
|[<span data-ttu-id="0e073-924">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-924">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-925">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-925">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="0e073-926">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0e073-926">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="0e073-927">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e073-927">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-928">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-928">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e073-929">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="0e073-929">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0e073-930">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="0e073-930">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="0e073-931">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="0e073-931">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="0e073-932">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="0e073-932">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="0e073-933">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="0e073-933">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-934">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-934">Parameters</span></span>

|<span data-ttu-id="0e073-935">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-935">Name</span></span>|<span data-ttu-id="0e073-936">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-936">Type</span></span>|<span data-ttu-id="0e073-937">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-937">Attributes</span></span>|<span data-ttu-id="0e073-938">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-938">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="0e073-939">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0e073-939">String &#124; Object</span></span>||<span data-ttu-id="0e073-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0e073-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0e073-942">**OU**</span><span class="sxs-lookup"><span data-stu-id="0e073-942">**OR**</span></span><br/><span data-ttu-id="0e073-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="0e073-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="0e073-945">String</span><span class="sxs-lookup"><span data-stu-id="0e073-945">String</span></span>|<span data-ttu-id="0e073-946">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-946">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0e073-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="0e073-949">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-949">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="0e073-950">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-950">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-951">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-951">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="0e073-952">String</span><span class="sxs-lookup"><span data-stu-id="0e073-952">String</span></span>||<span data-ttu-id="0e073-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="0e073-955">String</span><span class="sxs-lookup"><span data-stu-id="0e073-955">String</span></span>||<span data-ttu-id="0e073-956">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e073-956">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="0e073-957">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-957">String</span></span>||<span data-ttu-id="0e073-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="0e073-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="0e073-960">Booléen</span><span class="sxs-lookup"><span data-stu-id="0e073-960">Boolean</span></span>||<span data-ttu-id="0e073-p151">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e073-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="0e073-963">String</span><span class="sxs-lookup"><span data-stu-id="0e073-963">String</span></span>||<span data-ttu-id="0e073-p152">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e073-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="0e073-967">function</span><span class="sxs-lookup"><span data-stu-id="0e073-967">function</span></span>|<span data-ttu-id="0e073-968">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-968">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-969">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-969">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-970">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-970">Requirements</span></span>

|<span data-ttu-id="0e073-971">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-971">Requirement</span></span>|<span data-ttu-id="0e073-972">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-972">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-973">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-973">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-974">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-974">1.0</span></span>|
|[<span data-ttu-id="0e073-975">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-975">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-976">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-976">ReadItem</span></span>|
|[<span data-ttu-id="0e073-977">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-977">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-978">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-978">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e073-979">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e073-979">Examples</span></span>

<span data-ttu-id="0e073-980">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="0e073-980">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="0e073-981">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="0e073-981">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="0e073-982">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="0e073-982">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0e073-983">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="0e073-983">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0e073-984">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-984">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0e073-985">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-985">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="0e073-986">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0e073-986">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="0e073-987">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e073-987">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-988">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-988">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e073-989">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="0e073-989">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0e073-990">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="0e073-990">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="0e073-991">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="0e073-991">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="0e073-992">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="0e073-992">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="0e073-993">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="0e073-993">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-994">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-994">Parameters</span></span>

|<span data-ttu-id="0e073-995">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-995">Name</span></span>|<span data-ttu-id="0e073-996">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-996">Type</span></span>|<span data-ttu-id="0e073-997">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-997">Attributes</span></span>|<span data-ttu-id="0e073-998">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-998">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="0e073-999">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0e073-999">String &#124; Object</span></span>||<span data-ttu-id="0e073-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0e073-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0e073-1002">**OU**</span><span class="sxs-lookup"><span data-stu-id="0e073-1002">**OR**</span></span><br/><span data-ttu-id="0e073-p155">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="0e073-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="0e073-1005">String</span><span class="sxs-lookup"><span data-stu-id="0e073-1005">String</span></span>|<span data-ttu-id="0e073-1006">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-p156">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0e073-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="0e073-1009">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1009">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="0e073-1010">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1011">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-1011">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="0e073-1012">String</span><span class="sxs-lookup"><span data-stu-id="0e073-1012">String</span></span>||<span data-ttu-id="0e073-p157">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="0e073-1015">String</span><span class="sxs-lookup"><span data-stu-id="0e073-1015">String</span></span>||<span data-ttu-id="0e073-1016">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e073-1016">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="0e073-1017">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-1017">String</span></span>||<span data-ttu-id="0e073-p158">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="0e073-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="0e073-1020">Booléen</span><span class="sxs-lookup"><span data-stu-id="0e073-1020">Boolean</span></span>||<span data-ttu-id="0e073-p159">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0e073-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="0e073-1023">String</span><span class="sxs-lookup"><span data-stu-id="0e073-1023">String</span></span>||<span data-ttu-id="0e073-p160">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="0e073-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="0e073-1027">function</span><span class="sxs-lookup"><span data-stu-id="0e073-1027">function</span></span>|<span data-ttu-id="0e073-1028">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1028">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1029">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1030">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1030">Requirements</span></span>

|<span data-ttu-id="0e073-1031">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1031">Requirement</span></span>|<span data-ttu-id="0e073-1032">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1032">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1033">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1033">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1034">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-1034">1.0</span></span>|
|[<span data-ttu-id="0e073-1035">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1035">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1036">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1036">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1037">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1037">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1038">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1038">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e073-1039">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e073-1039">Examples</span></span>

<span data-ttu-id="0e073-1040">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1040">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="0e073-1041">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="0e073-1041">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="0e073-1042">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="0e073-1042">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0e073-1043">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="0e073-1043">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0e073-1044">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-1044">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0e073-1045">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1045">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="0e073-1046">getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="0e073-1046">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="0e073-1047">Obtient la pièce jointe spécifiée à partir d’un message ou d’un `AttachmentContent` rendez-vous et la renvoie en tant qu’objet.</span><span class="sxs-lookup"><span data-stu-id="0e073-1047">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="0e073-1048">La `getAttachmentContentAsync` méthode obtient la pièce jointe avec l’identificateur spécifié à partir de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-1048">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0e073-1049">Il est recommandé d’utiliser l’identificateur pour récupérer une pièce jointe dans la même session que l’attachmentIds a été récupérée avec l' `getAttachmentsAsync` appel ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="0e073-1049">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="0e073-1050">Dans Outlook sur le Web et les appareils mobiles, l’identificateur de pièce jointe est valide uniquement au sein de la même session.</span><span class="sxs-lookup"><span data-stu-id="0e073-1050">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0e073-1051">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="0e073-1051">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1052">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1052">Parameters</span></span>

|<span data-ttu-id="0e073-1053">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1053">Name</span></span>|<span data-ttu-id="0e073-1054">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1054">Type</span></span>|<span data-ttu-id="0e073-1055">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1055">Attributes</span></span>|<span data-ttu-id="0e073-1056">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1056">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="0e073-1057">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-1057">String</span></span>||<span data-ttu-id="0e073-1058">Identificateur de la pièce jointe que vous souhaitez obtenir.</span><span class="sxs-lookup"><span data-stu-id="0e073-1058">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="0e073-1059">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1059">Object</span></span>|<span data-ttu-id="0e073-1060">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1060">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1061">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1061">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-1062">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1062">Object</span></span>|<span data-ttu-id="0e073-1063">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1064">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1064">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e073-1065">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1065">function</span></span>|<span data-ttu-id="0e073-1066">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1067">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1067">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1068">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1068">Requirements</span></span>

|<span data-ttu-id="0e073-1069">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1069">Requirement</span></span>|<span data-ttu-id="0e073-1070">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1070">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1071">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1071">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1072">Aperçu</span><span class="sxs-lookup"><span data-stu-id="0e073-1072">Preview</span></span>|
|[<span data-ttu-id="0e073-1073">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1073">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1074">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1074">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1075">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1075">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1076">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1076">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1077">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1077">Returns:</span></span>

<span data-ttu-id="0e073-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="0e073-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="0e073-1079">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1079">Example</span></span>

```javascript
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
  if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="0e073-1080">getAttachmentsAsync ([options], [Rappel]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0e073-1080">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="0e073-1081">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="0e073-1081">Gets the item's attachments as an array.</span></span> <span data-ttu-id="0e073-1082">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-1082">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1083">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1083">Parameters</span></span>

|<span data-ttu-id="0e073-1084">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1084">Name</span></span>|<span data-ttu-id="0e073-1085">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1085">Type</span></span>|<span data-ttu-id="0e073-1086">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1086">Attributes</span></span>|<span data-ttu-id="0e073-1087">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1087">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e073-1088">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1088">Object</span></span>|<span data-ttu-id="0e073-1089">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1089">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1090">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1090">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-1091">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1091">Object</span></span>|<span data-ttu-id="0e073-1092">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1093">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1093">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e073-1094">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1094">function</span></span>|<span data-ttu-id="0e073-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1096">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1096">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1097">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1097">Requirements</span></span>

|<span data-ttu-id="0e073-1098">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1098">Requirement</span></span>|<span data-ttu-id="0e073-1099">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1099">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1100">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1100">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1101">Aperçu</span><span class="sxs-lookup"><span data-stu-id="0e073-1101">Preview</span></span>|
|[<span data-ttu-id="0e073-1102">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1102">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1103">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1103">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1104">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1104">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1105">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-1105">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1106">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1106">Returns:</span></span>

<span data-ttu-id="0e073-1107">Type: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0e073-1107">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="0e073-1108">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1108">Example</span></span>

<span data-ttu-id="0e073-1109">L’exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e073-1109">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="0e073-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="0e073-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="0e073-1111">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e073-1111">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1112">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-1112">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-1113">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1113">Requirements</span></span>

|<span data-ttu-id="0e073-1114">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1114">Requirement</span></span>|<span data-ttu-id="0e073-1115">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1115">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1116">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1116">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1117">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-1117">1.0</span></span>|
|[<span data-ttu-id="0e073-1118">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1118">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1119">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1119">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1120">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1120">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1121">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1121">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1122">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1122">Returns:</span></span>

<span data-ttu-id="0e073-1123">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="0e073-1123">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="0e073-1124">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1124">Example</span></span>

<span data-ttu-id="0e073-1125">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e073-1125">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="0e073-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="0e073-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="0e073-1127">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e073-1127">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1128">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-1128">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1129">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1129">Parameters</span></span>

|<span data-ttu-id="0e073-1130">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1130">Name</span></span>|<span data-ttu-id="0e073-1131">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1131">Type</span></span>|<span data-ttu-id="0e073-1132">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1132">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="0e073-1133">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="0e073-1133">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="0e073-1134">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="0e073-1134">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1135">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1135">Requirements</span></span>

|<span data-ttu-id="0e073-1136">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1136">Requirement</span></span>|<span data-ttu-id="0e073-1137">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1138">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1139">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-1139">1.0</span></span>|
|[<span data-ttu-id="0e073-1140">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1141">Restreinte</span><span class="sxs-lookup"><span data-stu-id="0e073-1141">Restricted</span></span>|
|[<span data-ttu-id="0e073-1142">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1143">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1144">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1144">Returns:</span></span>

<span data-ttu-id="0e073-1145">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="0e073-1145">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="0e073-1146">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="0e073-1146">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="0e073-1147">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1147">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="0e073-1148">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="0e073-1148">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="0e073-1149">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="0e073-1149">Value of `entityType`</span></span>|<span data-ttu-id="0e073-1150">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="0e073-1150">Type of objects in returned array</span></span>|<span data-ttu-id="0e073-1151">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="0e073-1151">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="0e073-1152">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-1152">String</span></span>|<span data-ttu-id="0e073-1153">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0e073-1153">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="0e073-1154">Contact</span><span class="sxs-lookup"><span data-stu-id="0e073-1154">Contact</span></span>|<span data-ttu-id="0e073-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e073-1155">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="0e073-1156">String</span><span class="sxs-lookup"><span data-stu-id="0e073-1156">String</span></span>|<span data-ttu-id="0e073-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e073-1157">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="0e073-1158">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="0e073-1158">MeetingSuggestion</span></span>|<span data-ttu-id="0e073-1159">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e073-1159">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="0e073-1160">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="0e073-1160">PhoneNumber</span></span>|<span data-ttu-id="0e073-1161">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0e073-1161">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="0e073-1162">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="0e073-1162">TaskSuggestion</span></span>|<span data-ttu-id="0e073-1163">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e073-1163">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="0e073-1164">String</span><span class="sxs-lookup"><span data-stu-id="0e073-1164">String</span></span>|<span data-ttu-id="0e073-1165">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0e073-1165">**Restricted**</span></span>|

<span data-ttu-id="0e073-1166">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="0e073-1166">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="0e073-1167">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1167">Example</span></span>

<span data-ttu-id="0e073-1168">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0e073-1168">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="0e073-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="0e073-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="0e073-1170">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0e073-1170">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1171">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-1171">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e073-1172">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="0e073-1172">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1173">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1173">Parameters</span></span>

|<span data-ttu-id="0e073-1174">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1174">Name</span></span>|<span data-ttu-id="0e073-1175">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1175">Type</span></span>|<span data-ttu-id="0e073-1176">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1176">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="0e073-1177">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-1177">String</span></span>|<span data-ttu-id="0e073-1178">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="0e073-1178">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1179">Requirements</span></span>

|<span data-ttu-id="0e073-1180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1180">Requirement</span></span>|<span data-ttu-id="0e073-1181">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1181">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1183">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-1183">1.0</span></span>|
|[<span data-ttu-id="0e073-1184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1185">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1187">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1187">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1188">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1188">Returns:</span></span>

<span data-ttu-id="0e073-p164">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="0e073-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="0e073-1191">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="0e073-1191">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="0e073-1192">getInitializationContextAsync ([options], [Rappel])</span><span class="sxs-lookup"><span data-stu-id="0e073-1192">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="0e073-1193">Obtient les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="0e073-1193">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1194">Cette méthode est uniquement prise en charge par Outlook 2016 ou une version ultérieure sur Windows (versions «démarrer en un clic» ultérieures à 16.0.8413.1000) et Outlook sur le Web pour Office 365.</span><span class="sxs-lookup"><span data-stu-id="0e073-1194">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1195">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1195">Parameters</span></span>

|<span data-ttu-id="0e073-1196">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1196">Name</span></span>|<span data-ttu-id="0e073-1197">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1197">Type</span></span>|<span data-ttu-id="0e073-1198">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1198">Attributes</span></span>|<span data-ttu-id="0e073-1199">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1199">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e073-1200">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1200">Object</span></span>|<span data-ttu-id="0e073-1201">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1201">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1202">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1202">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-1203">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1203">Object</span></span>|<span data-ttu-id="0e073-1204">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1204">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1205">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1205">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e073-1206">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1206">function</span></span>|<span data-ttu-id="0e073-1207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1207">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1208">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1208">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e073-1209">En cas de réussite, les données d’initialisation sont fournies `asyncResult.value` dans la propriété sous la forme d’une chaîne.</span><span class="sxs-lookup"><span data-stu-id="0e073-1209">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="0e073-1210">S’il n’existe pas de contexte d’initialisation `asyncResult` , l’objet contient `Error` un objet dont `code` la propriété est `9020` définie sur `name` et sa propriété `GenericResponseError`est définie sur.</span><span class="sxs-lookup"><span data-stu-id="0e073-1210">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1211">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1211">Requirements</span></span>

|<span data-ttu-id="0e073-1212">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1212">Requirement</span></span>|<span data-ttu-id="0e073-1213">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1213">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1214">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1215">Aperçu</span><span class="sxs-lookup"><span data-stu-id="0e073-1215">Preview</span></span>|
|[<span data-ttu-id="0e073-1216">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1217">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1218">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1219">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1219">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-1220">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1220">Example</span></span>

```javascript
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

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="0e073-1221">getItemIdAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="0e073-1221">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="0e073-1222">Obtient de manière asynchrone l’ID d’un élément enregistré.</span><span class="sxs-lookup"><span data-stu-id="0e073-1222">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="0e073-1223">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="0e073-1223">Compose mode only.</span></span>

<span data-ttu-id="0e073-1224">Lorsqu’elle est appelée, cette méthode renvoie l’ID de l’élément par le biais de la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1224">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1225">Si votre complément appelle `getItemIdAsync` sur un élément en mode composition (par exemple, pour obtenir un à utiliser avec `itemId` EWS ou l’API REST), sachez que lorsque Outlook est en mode mis en cache, l’élément peut prendre un certain temps avant la synchronisation de l’élément avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e073-1225">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="0e073-1226">Tant que l’élément n’est pas synchronisé `itemId` , le n’est pas reconnu et son utilisation renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="0e073-1226">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1227">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1227">Parameters</span></span>

|<span data-ttu-id="0e073-1228">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1228">Name</span></span>|<span data-ttu-id="0e073-1229">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1229">Type</span></span>|<span data-ttu-id="0e073-1230">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1230">Attributes</span></span>|<span data-ttu-id="0e073-1231">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1231">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e073-1232">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1232">Object</span></span>|<span data-ttu-id="0e073-1233">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1233">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1234">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1234">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-1235">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1235">Object</span></span>|<span data-ttu-id="0e073-1236">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1236">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1237">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1237">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e073-1238">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1238">function</span></span>||<span data-ttu-id="0e073-1239">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1239">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e073-1240">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1240">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e073-1241">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e073-1241">Errors</span></span>

|<span data-ttu-id="0e073-1242">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e073-1242">Error code</span></span>|<span data-ttu-id="0e073-1243">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1243">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="0e073-1244">L’ID ne peut pas être récupéré tant que l’élément n’est pas enregistré.</span><span class="sxs-lookup"><span data-stu-id="0e073-1244">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1245">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1245">Requirements</span></span>

|<span data-ttu-id="0e073-1246">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1246">Requirement</span></span>|<span data-ttu-id="0e073-1247">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1247">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1248">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1249">Aperçu</span><span class="sxs-lookup"><span data-stu-id="0e073-1249">Preview</span></span>|
|[<span data-ttu-id="0e073-1250">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1250">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1251">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1252">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1252">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1253">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-1253">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e073-1254">Exemples</span><span class="sxs-lookup"><span data-stu-id="0e073-1254">Examples</span></span>

```javascript
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="0e073-1255">L’exemple suivant montre la structure du `result` paramètre transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1255">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="0e073-1256">La `value` propriété contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-1256">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="0e073-1257">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0e073-1257">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="0e073-1258">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0e073-1258">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1259">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-1259">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e073-p168">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="0e073-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0e073-1263">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="0e073-1263">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0e073-1264">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1264">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0e073-p169">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-1268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1268">Requirements</span></span>

|<span data-ttu-id="0e073-1269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1269">Requirement</span></span>|<span data-ttu-id="0e073-1270">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-1272">1.0</span></span>|
|[<span data-ttu-id="0e073-1273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1274">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1276">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1277">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1277">Returns:</span></span>

<span data-ttu-id="0e073-p170">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="0e073-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="0e073-1280">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="0e073-1280">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0e073-1281">Object</span><span class="sxs-lookup"><span data-stu-id="0e073-1281">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0e073-1282">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1282">Example</span></span>

<span data-ttu-id="0e073-1283">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="0e073-1283">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="0e073-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="0e073-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="0e073-1285">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0e073-1285">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1286">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-1286">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e073-1287">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="0e073-1287">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="0e073-p171">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="0e073-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1290">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1290">Parameters</span></span>

|<span data-ttu-id="0e073-1291">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1291">Name</span></span>|<span data-ttu-id="0e073-1292">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1292">Type</span></span>|<span data-ttu-id="0e073-1293">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1293">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="0e073-1294">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0e073-1294">String</span></span>|<span data-ttu-id="0e073-1295">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="0e073-1295">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1296">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1296">Requirements</span></span>

|<span data-ttu-id="0e073-1297">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1297">Requirement</span></span>|<span data-ttu-id="0e073-1298">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1298">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1299">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1300">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-1300">1.0</span></span>|
|[<span data-ttu-id="0e073-1301">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1301">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1302">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1303">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1303">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1304">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1304">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1305">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1305">Returns:</span></span>

<span data-ttu-id="0e073-1306">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0e073-1306">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="0e073-1307">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="0e073-1307">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0e073-1308">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="0e073-1308">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0e073-1309">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1309">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="0e073-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="0e073-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="0e073-1311">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="0e073-1311">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="0e073-p172">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="0e073-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1314">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1314">Parameters</span></span>

|<span data-ttu-id="0e073-1315">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1315">Name</span></span>|<span data-ttu-id="0e073-1316">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1316">Type</span></span>|<span data-ttu-id="0e073-1317">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1317">Attributes</span></span>|<span data-ttu-id="0e073-1318">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1318">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="0e073-1319">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0e073-1319">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="0e073-p173">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="0e073-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="0e073-1323">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1323">Object</span></span>|<span data-ttu-id="0e073-1324">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1324">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1325">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1325">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-1326">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1326">Object</span></span>|<span data-ttu-id="0e073-1327">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1327">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1328">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1328">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e073-1329">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1329">function</span></span>||<span data-ttu-id="0e073-1330">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1330">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e073-1331">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1331">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="0e073-1332">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1332">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1333">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1333">Requirements</span></span>

|<span data-ttu-id="0e073-1334">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1334">Requirement</span></span>|<span data-ttu-id="0e073-1335">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1335">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1336">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1337">1.2</span><span class="sxs-lookup"><span data-stu-id="0e073-1337">1.2</span></span>|
|[<span data-ttu-id="0e073-1338">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1339">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1339">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e073-1340">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1341">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-1341">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1342">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1342">Returns:</span></span>

<span data-ttu-id="0e073-1343">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1343">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="0e073-1344">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="0e073-1344">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0e073-1345">String</span><span class="sxs-lookup"><span data-stu-id="0e073-1345">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0e073-1346">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1346">Example</span></span>

```javascript
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

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="0e073-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="0e073-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="0e073-1348">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e073-1348">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="0e073-1349">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="0e073-1349">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1350">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-1350">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-1351">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1351">Requirements</span></span>

|<span data-ttu-id="0e073-1352">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1352">Requirement</span></span>|<span data-ttu-id="0e073-1353">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1354">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1355">1.6</span><span class="sxs-lookup"><span data-stu-id="0e073-1355">1.6</span></span>|
|[<span data-ttu-id="0e073-1356">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1356">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1357">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1358">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1358">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1359">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1359">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1360">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1360">Returns:</span></span>

<span data-ttu-id="0e073-1361">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="0e073-1361">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="0e073-1362">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1362">Example</span></span>

<span data-ttu-id="0e073-1363">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0e073-1363">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="0e073-1364">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0e073-1364">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="0e073-p176">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="0e073-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1367">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0e073-1367">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0e073-p177">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="0e073-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0e073-1371">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="0e073-1371">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0e073-1372">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1372">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0e073-p178">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e073-1376">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1376">Requirements</span></span>

|<span data-ttu-id="0e073-1377">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1377">Requirement</span></span>|<span data-ttu-id="0e073-1378">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1378">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1379">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1380">1.6</span><span class="sxs-lookup"><span data-stu-id="0e073-1380">1.6</span></span>|
|[<span data-ttu-id="0e073-1381">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1382">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1383">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1384">Lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1384">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e073-1385">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0e073-1385">Returns:</span></span>

<span data-ttu-id="0e073-p179">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="0e073-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="0e073-1388">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1388">Example</span></span>

<span data-ttu-id="0e073-1389">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="0e073-1389">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="0e073-1390">getSharedPropertiesAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="0e073-1390">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="0e073-1391">Obtient les propriétés du rendez-vous ou du message sélectionné dans un dossier partagé, un calendrier ou une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="0e073-1391">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1392">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1392">Parameters</span></span>

|<span data-ttu-id="0e073-1393">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1393">Name</span></span>|<span data-ttu-id="0e073-1394">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1394">Type</span></span>|<span data-ttu-id="0e073-1395">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1395">Attributes</span></span>|<span data-ttu-id="0e073-1396">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1396">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e073-1397">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1397">Object</span></span>|<span data-ttu-id="0e073-1398">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1398">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1399">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1399">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-1400">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1400">Object</span></span>|<span data-ttu-id="0e073-1401">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1401">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1402">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1402">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e073-1403">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1403">function</span></span>||<span data-ttu-id="0e073-1404">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1404">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e073-1405">Les propriétés partagées sont fournies sous [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) la forme d' `asyncResult.value` un objet dans la propriété.</span><span class="sxs-lookup"><span data-stu-id="0e073-1405">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0e073-1406">Cet objet peut être utilisé pour obtenir les propriétés partagées de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-1406">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1407">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1407">Requirements</span></span>

|<span data-ttu-id="0e073-1408">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1408">Requirement</span></span>|<span data-ttu-id="0e073-1409">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1410">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1411">Aperçu</span><span class="sxs-lookup"><span data-stu-id="0e073-1411">Preview</span></span>|
|[<span data-ttu-id="0e073-1412">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1413">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1414">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1415">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1415">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-1416">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1416">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="0e073-1417">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0e073-1417">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="0e073-1418">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0e073-1418">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="0e073-p181">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="0e073-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1422">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1422">Parameters</span></span>

|<span data-ttu-id="0e073-1423">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1423">Name</span></span>|<span data-ttu-id="0e073-1424">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1424">Type</span></span>|<span data-ttu-id="0e073-1425">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1425">Attributes</span></span>|<span data-ttu-id="0e073-1426">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1426">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="0e073-1427">function</span><span class="sxs-lookup"><span data-stu-id="0e073-1427">function</span></span>||<span data-ttu-id="0e073-1428">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1428">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e073-1429">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1429">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0e073-1430">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e073-1430">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="0e073-1431">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1431">Object</span></span>|<span data-ttu-id="0e073-1432">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1433">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1433">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="0e073-1434">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1434">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1435">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1435">Requirements</span></span>

|<span data-ttu-id="0e073-1436">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1436">Requirement</span></span>|<span data-ttu-id="0e073-1437">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1437">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1438">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1439">1.0</span><span class="sxs-lookup"><span data-stu-id="0e073-1439">1.0</span></span>|
|[<span data-ttu-id="0e073-1440">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1440">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1441">ReadItem</span></span>|
|[<span data-ttu-id="0e073-1442">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1442">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1443">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1443">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-1444">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1444">Example</span></span>

<span data-ttu-id="0e073-p184">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0e073-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="0e073-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e073-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="0e073-1449">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0e073-1449">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="0e073-1450">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-1450">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0e073-1451">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="0e073-1451">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="0e073-1452">Dans Outlook sur le Web et les appareils mobiles, l’identificateur de pièce jointe est valide uniquement au sein de la même session.</span><span class="sxs-lookup"><span data-stu-id="0e073-1452">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0e073-1453">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="0e073-1453">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1454">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1454">Parameters</span></span>

|<span data-ttu-id="0e073-1455">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1455">Name</span></span>|<span data-ttu-id="0e073-1456">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1456">Type</span></span>|<span data-ttu-id="0e073-1457">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1457">Attributes</span></span>|<span data-ttu-id="0e073-1458">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1458">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="0e073-1459">String</span><span class="sxs-lookup"><span data-stu-id="0e073-1459">String</span></span>||<span data-ttu-id="0e073-1460">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="0e073-1460">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="0e073-1461">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1461">Object</span></span>|<span data-ttu-id="0e073-1462">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1462">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1463">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1463">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-1464">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1464">Object</span></span>|<span data-ttu-id="0e073-1465">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1466">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1466">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e073-1467">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1467">function</span></span>|<span data-ttu-id="0e073-1468">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1468">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1469">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1469">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e073-1470">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="0e073-1470">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e073-1471">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0e073-1471">Errors</span></span>

|<span data-ttu-id="0e073-1472">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0e073-1472">Error code</span></span>|<span data-ttu-id="0e073-1473">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1473">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="0e073-1474">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="0e073-1474">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1475">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1475">Requirements</span></span>

|<span data-ttu-id="0e073-1476">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1476">Requirement</span></span>|<span data-ttu-id="0e073-1477">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1477">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1478">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1478">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1479">1.1</span><span class="sxs-lookup"><span data-stu-id="0e073-1479">1.1</span></span>|
|[<span data-ttu-id="0e073-1480">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1480">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1481">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1481">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e073-1482">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1482">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1483">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-1483">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-1484">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1484">Example</span></span>

<span data-ttu-id="0e073-1485">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="0e073-1485">The following code removes an attachment with an identifier of '0'.</span></span>

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

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="0e073-1486">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e073-1486">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="0e073-1487">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="0e073-1487">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="0e073-1488">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="0e073-1488">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1489">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1489">Parameters</span></span>

| <span data-ttu-id="0e073-1490">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1490">Name</span></span> | <span data-ttu-id="0e073-1491">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1491">Type</span></span> | <span data-ttu-id="0e073-1492">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1492">Attributes</span></span> | <span data-ttu-id="0e073-1493">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1493">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0e073-1494">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0e073-1494">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0e073-1495">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="0e073-1495">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="0e073-1496">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1496">Object</span></span> | <span data-ttu-id="0e073-1497">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1497">&lt;optional&gt;</span></span> | <span data-ttu-id="0e073-1498">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1498">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0e073-1499">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1499">Object</span></span> | <span data-ttu-id="0e073-1500">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1500">&lt;optional&gt;</span></span> | <span data-ttu-id="0e073-1501">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1501">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0e073-1502">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1502">function</span></span>| <span data-ttu-id="0e073-1503">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1503">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1504">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1505">Requirements</span></span>

|<span data-ttu-id="0e073-1506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1506">Requirement</span></span>| <span data-ttu-id="0e073-1507">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e073-1509">1.7</span><span class="sxs-lookup"><span data-stu-id="0e073-1509">1.7</span></span> |
|[<span data-ttu-id="0e073-1510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e073-1511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1511">ReadItem</span></span> |
|[<span data-ttu-id="0e073-1512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0e073-1513">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0e073-1513">Compose or Read</span></span> |

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="0e073-1514">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0e073-1514">saveAsync([options], callback)</span></span>

<span data-ttu-id="0e073-1515">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0e073-1515">Asynchronously saves an item.</span></span>

<span data-ttu-id="0e073-1516">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1516">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="0e073-1517">Dans Outlook sur le Web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="0e073-1517">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="0e073-1518">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="0e073-1518">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1519">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="0e073-1519">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="0e073-1520">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="0e073-1520">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="0e073-p188">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="0e073-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="0e073-1524">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="0e073-1524">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="0e073-1525">Outlook sur Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="0e073-1525">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="0e073-1526">La `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="0e073-1526">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="0e073-1527">Consultez la rubrique [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide de l’API Office js](https://support.microsoft.com/help/4505745) pour obtenir une solution de contournement.</span><span class="sxs-lookup"><span data-stu-id="0e073-1527">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="0e073-1528">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="0e073-1528">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1529">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1529">Parameters</span></span>

|<span data-ttu-id="0e073-1530">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1530">Name</span></span>|<span data-ttu-id="0e073-1531">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1531">Type</span></span>|<span data-ttu-id="0e073-1532">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1532">Attributes</span></span>|<span data-ttu-id="0e073-1533">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1533">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0e073-1534">Object</span><span class="sxs-lookup"><span data-stu-id="0e073-1534">Object</span></span>|<span data-ttu-id="0e073-1535">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1535">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1536">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1536">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-1537">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1537">Object</span></span>|<span data-ttu-id="0e073-1538">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1538">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1539">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1539">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0e073-1540">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1540">function</span></span>||<span data-ttu-id="0e073-1541">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e073-1542">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e073-1542">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1543">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1543">Requirements</span></span>

|<span data-ttu-id="0e073-1544">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1544">Requirement</span></span>|<span data-ttu-id="0e073-1545">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1546">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1547">1.3</span><span class="sxs-lookup"><span data-stu-id="0e073-1547">1.3</span></span>|
|[<span data-ttu-id="0e073-1548">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1549">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1549">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e073-1550">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1551">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-1551">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e073-1552">範例</span><span class="sxs-lookup"><span data-stu-id="0e073-1552">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="0e073-p190">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0e073-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="0e073-1555">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="0e073-1555">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="0e073-1556">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0e073-1556">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="0e073-p191">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="0e073-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e073-1560">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0e073-1560">Parameters</span></span>

|<span data-ttu-id="0e073-1561">Nom</span><span class="sxs-lookup"><span data-stu-id="0e073-1561">Name</span></span>|<span data-ttu-id="0e073-1562">Type</span><span class="sxs-lookup"><span data-stu-id="0e073-1562">Type</span></span>|<span data-ttu-id="0e073-1563">Attributs</span><span class="sxs-lookup"><span data-stu-id="0e073-1563">Attributes</span></span>|<span data-ttu-id="0e073-1564">Description</span><span class="sxs-lookup"><span data-stu-id="0e073-1564">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="0e073-1565">String</span><span class="sxs-lookup"><span data-stu-id="0e073-1565">String</span></span>||<span data-ttu-id="0e073-p192">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="0e073-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="0e073-1569">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1569">Object</span></span>|<span data-ttu-id="0e073-1570">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1570">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1571">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0e073-1571">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0e073-1572">Objet</span><span class="sxs-lookup"><span data-stu-id="0e073-1572">Object</span></span>|<span data-ttu-id="0e073-1573">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1573">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1574">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0e073-1574">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="0e073-1575">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0e073-1575">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="0e073-1576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0e073-1576">&lt;optional&gt;</span></span>|<span data-ttu-id="0e073-1577">Si `text`, le style actuel est appliqué dans Outlook sur le Web et les clients de bureau.</span><span class="sxs-lookup"><span data-stu-id="0e073-1577">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="0e073-1578">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="0e073-1578">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="0e073-1579">Si `html` et que le champ prend en charge le format html (l’objet ne l’est pas), le style actuel est appliqué dans Outlook sur le Web et le style par défaut est appliqué dans les clients de bureau Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e073-1579">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="0e073-1580">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="0e073-1580">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="0e073-1581">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="0e073-1581">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="0e073-1582">fonction</span><span class="sxs-lookup"><span data-stu-id="0e073-1582">function</span></span>||<span data-ttu-id="0e073-1583">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e073-1583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e073-1584">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0e073-1584">Requirements</span></span>

|<span data-ttu-id="0e073-1585">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0e073-1585">Requirement</span></span>|<span data-ttu-id="0e073-1586">Valeur</span><span class="sxs-lookup"><span data-stu-id="0e073-1586">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e073-1587">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0e073-1587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0e073-1588">1.2</span><span class="sxs-lookup"><span data-stu-id="0e073-1588">1.2</span></span>|
|[<span data-ttu-id="0e073-1589">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0e073-1589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0e073-1590">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e073-1590">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e073-1591">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0e073-1591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0e073-1592">Composition</span><span class="sxs-lookup"><span data-stu-id="0e073-1592">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e073-1593">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e073-1593">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
