---
title: Office. Context. Mailbox. Item-Preview ensemble de conditions requises
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: 12ec5d5558b558c87587e34472c33116478d14b3
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589201"
---
# <a name="item"></a><span data-ttu-id="95bbb-102">élément</span><span class="sxs-lookup"><span data-stu-id="95bbb-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="95bbb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="95bbb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="95bbb-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="95bbb-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-106">Requirements</span></span>

|<span data-ttu-id="95bbb-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-107">Requirement</span></span>|<span data-ttu-id="95bbb-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-110">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-110">1.0</span></span>|
|[<span data-ttu-id="95bbb-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="95bbb-112">Restricted</span></span>|
|[<span data-ttu-id="95bbb-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="95bbb-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="95bbb-115">Members and methods</span></span>

| <span data-ttu-id="95bbb-116">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-116">Member</span></span> | <span data-ttu-id="95bbb-117">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="95bbb-118">attachments</span><span class="sxs-lookup"><span data-stu-id="95bbb-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="95bbb-119">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-119">Member</span></span> |
| [<span data-ttu-id="95bbb-120">bcc</span><span class="sxs-lookup"><span data-stu-id="95bbb-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="95bbb-121">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-121">Member</span></span> |
| [<span data-ttu-id="95bbb-122">body</span><span class="sxs-lookup"><span data-stu-id="95bbb-122">body</span></span>](#body-body) | <span data-ttu-id="95bbb-123">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-123">Member</span></span> |
| [<span data-ttu-id="95bbb-124">catégories</span><span class="sxs-lookup"><span data-stu-id="95bbb-124">categories</span></span>](#categories-categories) | <span data-ttu-id="95bbb-125">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-125">Member</span></span> |
| [<span data-ttu-id="95bbb-126">cc</span><span class="sxs-lookup"><span data-stu-id="95bbb-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="95bbb-127">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-127">Member</span></span> |
| [<span data-ttu-id="95bbb-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="95bbb-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="95bbb-129">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-129">Member</span></span> |
| [<span data-ttu-id="95bbb-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="95bbb-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="95bbb-131">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-131">Member</span></span> |
| [<span data-ttu-id="95bbb-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="95bbb-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="95bbb-133">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-133">Member</span></span> |
| [<span data-ttu-id="95bbb-134">end</span><span class="sxs-lookup"><span data-stu-id="95bbb-134">end</span></span>](#end-datetime) | <span data-ttu-id="95bbb-135">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-135">Member</span></span> |
| [<span data-ttu-id="95bbb-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="95bbb-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="95bbb-137">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-137">Member</span></span> |
| [<span data-ttu-id="95bbb-138">from</span><span class="sxs-lookup"><span data-stu-id="95bbb-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="95bbb-139">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-139">Member</span></span> |
| [<span data-ttu-id="95bbb-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="95bbb-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="95bbb-141">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-141">Member</span></span> |
| [<span data-ttu-id="95bbb-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="95bbb-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="95bbb-143">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-143">Member</span></span> |
| [<span data-ttu-id="95bbb-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="95bbb-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="95bbb-145">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-145">Member</span></span> |
| [<span data-ttu-id="95bbb-146">itemId</span><span class="sxs-lookup"><span data-stu-id="95bbb-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="95bbb-147">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-147">Member</span></span> |
| [<span data-ttu-id="95bbb-148">itemType</span><span class="sxs-lookup"><span data-stu-id="95bbb-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="95bbb-149">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-149">Member</span></span> |
| [<span data-ttu-id="95bbb-150">location</span><span class="sxs-lookup"><span data-stu-id="95bbb-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="95bbb-151">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-151">Member</span></span> |
| [<span data-ttu-id="95bbb-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="95bbb-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="95bbb-153">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-153">Member</span></span> |
| [<span data-ttu-id="95bbb-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="95bbb-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="95bbb-155">Member</span><span class="sxs-lookup"><span data-stu-id="95bbb-155">Member</span></span> |
| [<span data-ttu-id="95bbb-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="95bbb-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="95bbb-157">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-157">Member</span></span> |
| [<span data-ttu-id="95bbb-158">organizer</span><span class="sxs-lookup"><span data-stu-id="95bbb-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="95bbb-159">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-159">Member</span></span> |
| [<span data-ttu-id="95bbb-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="95bbb-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="95bbb-161">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-161">Member</span></span> |
| [<span data-ttu-id="95bbb-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="95bbb-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="95bbb-163">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-163">Member</span></span> |
| [<span data-ttu-id="95bbb-164">sender</span><span class="sxs-lookup"><span data-stu-id="95bbb-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="95bbb-165">Member</span><span class="sxs-lookup"><span data-stu-id="95bbb-165">Member</span></span> |
| [<span data-ttu-id="95bbb-166">seriesId</span><span class="sxs-lookup"><span data-stu-id="95bbb-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="95bbb-167">Member</span><span class="sxs-lookup"><span data-stu-id="95bbb-167">Member</span></span> |
| [<span data-ttu-id="95bbb-168">start</span><span class="sxs-lookup"><span data-stu-id="95bbb-168">start</span></span>](#start-datetime) | <span data-ttu-id="95bbb-169">Member</span><span class="sxs-lookup"><span data-stu-id="95bbb-169">Member</span></span> |
| [<span data-ttu-id="95bbb-170">subject</span><span class="sxs-lookup"><span data-stu-id="95bbb-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="95bbb-171">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-171">Member</span></span> |
| [<span data-ttu-id="95bbb-172">to</span><span class="sxs-lookup"><span data-stu-id="95bbb-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="95bbb-173">Membre</span><span class="sxs-lookup"><span data-stu-id="95bbb-173">Member</span></span> |
| [<span data-ttu-id="95bbb-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="95bbb-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-175">Method</span></span> |
| [<span data-ttu-id="95bbb-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="95bbb-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="95bbb-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-177">Method</span></span> |
| [<span data-ttu-id="95bbb-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="95bbb-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-179">Method</span></span> |
| [<span data-ttu-id="95bbb-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="95bbb-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-181">Method</span></span> |
| [<span data-ttu-id="95bbb-182">close</span><span class="sxs-lookup"><span data-stu-id="95bbb-182">close</span></span>](#close) | <span data-ttu-id="95bbb-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-183">Method</span></span> |
| [<span data-ttu-id="95bbb-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="95bbb-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="95bbb-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-185">Method</span></span> |
| [<span data-ttu-id="95bbb-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="95bbb-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="95bbb-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-187">Method</span></span> |
| [<span data-ttu-id="95bbb-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="95bbb-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-189">Method</span></span> |
| [<span data-ttu-id="95bbb-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="95bbb-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-191">Method</span></span> |
| [<span data-ttu-id="95bbb-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="95bbb-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="95bbb-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-193">Method</span></span> |
| [<span data-ttu-id="95bbb-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="95bbb-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="95bbb-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-195">Method</span></span> |
| [<span data-ttu-id="95bbb-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="95bbb-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="95bbb-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-197">Method</span></span> |
| [<span data-ttu-id="95bbb-198">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="95bbb-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-199">Method</span></span> |
| [<span data-ttu-id="95bbb-200">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="95bbb-200">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="95bbb-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-201">Method</span></span> |
| [<span data-ttu-id="95bbb-202">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="95bbb-202">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="95bbb-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-203">Method</span></span> |
| [<span data-ttu-id="95bbb-204">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-204">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="95bbb-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-205">Method</span></span> |
| [<span data-ttu-id="95bbb-206">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="95bbb-206">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="95bbb-207">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-207">Method</span></span> |
| [<span data-ttu-id="95bbb-208">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="95bbb-208">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="95bbb-209">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-209">Method</span></span> |
| [<span data-ttu-id="95bbb-210">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-210">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="95bbb-211">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-211">Method</span></span> |
| [<span data-ttu-id="95bbb-212">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-212">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="95bbb-213">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-213">Method</span></span> |
| [<span data-ttu-id="95bbb-214">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-214">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="95bbb-215">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-215">Method</span></span> |
| [<span data-ttu-id="95bbb-216">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-216">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="95bbb-217">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-217">Method</span></span> |
| [<span data-ttu-id="95bbb-218">saveAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-218">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="95bbb-219">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-219">Method</span></span> |
| [<span data-ttu-id="95bbb-220">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="95bbb-220">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="95bbb-221">Méthode</span><span class="sxs-lookup"><span data-stu-id="95bbb-221">Method</span></span> |

### <a name="example"></a><span data-ttu-id="95bbb-222">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-222">Example</span></span>

<span data-ttu-id="95bbb-223">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="95bbb-223">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="95bbb-224">Membres</span><span class="sxs-lookup"><span data-stu-id="95bbb-224">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="95bbb-225">pièces jointes: tableau. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="95bbb-225">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="95bbb-226">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="95bbb-226">Gets the item's attachments as an array.</span></span> <span data-ttu-id="95bbb-227">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-227">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-228">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="95bbb-228">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="95bbb-229">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="95bbb-229">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-230">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-230">Type</span></span>

*   <span data-ttu-id="95bbb-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="95bbb-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-232">Requirements</span></span>

|<span data-ttu-id="95bbb-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-233">Requirement</span></span>|<span data-ttu-id="95bbb-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-236">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-236">1.0</span></span>|
|[<span data-ttu-id="95bbb-237">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-238">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-240">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-240">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-241">Example</span></span>

<span data-ttu-id="95bbb-242">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="95bbb-242">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="95bbb-243">CCI: [destinataires](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="95bbb-243">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="95bbb-244">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="95bbb-244">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="95bbb-245">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-245">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-246">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-246">Type</span></span>

*   [<span data-ttu-id="95bbb-247">Destinataires</span><span class="sxs-lookup"><span data-stu-id="95bbb-247">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="95bbb-248">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-248">Requirements</span></span>

|<span data-ttu-id="95bbb-249">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-249">Requirement</span></span>|<span data-ttu-id="95bbb-250">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-251">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-252">1.1</span><span class="sxs-lookup"><span data-stu-id="95bbb-252">1.1</span></span>|
|[<span data-ttu-id="95bbb-253">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-254">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-255">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-256">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-256">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-257">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-257">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="95bbb-258">Body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="95bbb-258">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="95bbb-259">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-259">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-260">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-260">Type</span></span>

*   [<span data-ttu-id="95bbb-261">Body</span><span class="sxs-lookup"><span data-stu-id="95bbb-261">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="95bbb-262">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-262">Requirements</span></span>

|<span data-ttu-id="95bbb-263">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-263">Requirement</span></span>|<span data-ttu-id="95bbb-264">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-265">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-266">1.1</span><span class="sxs-lookup"><span data-stu-id="95bbb-266">1.1</span></span>|
|[<span data-ttu-id="95bbb-267">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-267">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-268">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-269">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-269">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-270">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-270">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-271">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-271">Example</span></span>

<span data-ttu-id="95bbb-272">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="95bbb-272">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="95bbb-273">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-273">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="95bbb-274">Catégories: [catégories](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="95bbb-274">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="95bbb-275">Obtient un objet qui fournit des méthodes pour la gestion des catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-275">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-276">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-276">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-277">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-277">Type</span></span>

*   [<span data-ttu-id="95bbb-278">Catégories</span><span class="sxs-lookup"><span data-stu-id="95bbb-278">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="95bbb-279">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-279">Requirements</span></span>

|<span data-ttu-id="95bbb-280">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-280">Requirement</span></span>|<span data-ttu-id="95bbb-281">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-282">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-283">Aperçu</span><span class="sxs-lookup"><span data-stu-id="95bbb-283">Preview</span></span>|
|[<span data-ttu-id="95bbb-284">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-285">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-286">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-287">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-288">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-288">Example</span></span>

<span data-ttu-id="95bbb-289">Cet exemple obtient les catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-289">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="95bbb-290">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="95bbb-290">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="95bbb-291">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="95bbb-291">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="95bbb-292">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="95bbb-292">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-293">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-293">Read mode</span></span>

<span data-ttu-id="95bbb-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-296">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-296">Compose mode</span></span>

<span data-ttu-id="95bbb-297">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="95bbb-297">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="95bbb-298">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-298">Type</span></span>

*   <span data-ttu-id="95bbb-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="95bbb-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-300">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-300">Requirements</span></span>

|<span data-ttu-id="95bbb-301">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-301">Requirement</span></span>|<span data-ttu-id="95bbb-302">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-303">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-304">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-304">1.0</span></span>|
|[<span data-ttu-id="95bbb-305">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-306">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-307">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-308">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-308">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="95bbb-309">(Nullable) conversationId: chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-309">(nullable) conversationId: String</span></span>

<span data-ttu-id="95bbb-310">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="95bbb-310">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="95bbb-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="95bbb-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-315">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-315">Type</span></span>

*   <span data-ttu-id="95bbb-316">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-316">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-317">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-317">Requirements</span></span>

|<span data-ttu-id="95bbb-318">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-318">Requirement</span></span>|<span data-ttu-id="95bbb-319">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-320">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-321">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-321">1.0</span></span>|
|[<span data-ttu-id="95bbb-322">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-323">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-324">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-325">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-325">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-326">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-326">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="95bbb-327">dateTimeCreated: date</span><span class="sxs-lookup"><span data-stu-id="95bbb-327">dateTimeCreated: Date</span></span>

<span data-ttu-id="95bbb-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-330">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-330">Type</span></span>

*   <span data-ttu-id="95bbb-331">Date</span><span class="sxs-lookup"><span data-stu-id="95bbb-331">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-332">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-332">Requirements</span></span>

|<span data-ttu-id="95bbb-333">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-333">Requirement</span></span>|<span data-ttu-id="95bbb-334">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-335">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-336">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-336">1.0</span></span>|
|[<span data-ttu-id="95bbb-337">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-338">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-339">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-340">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-340">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-341">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-341">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="95bbb-342">dateTimeModified: date</span><span class="sxs-lookup"><span data-stu-id="95bbb-342">dateTimeModified: Date</span></span>

<span data-ttu-id="95bbb-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-345">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-345">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-346">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-346">Type</span></span>

*   <span data-ttu-id="95bbb-347">Date</span><span class="sxs-lookup"><span data-stu-id="95bbb-347">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-348">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-348">Requirements</span></span>

|<span data-ttu-id="95bbb-349">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-349">Requirement</span></span>|<span data-ttu-id="95bbb-350">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-351">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-352">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-352">1.0</span></span>|
|[<span data-ttu-id="95bbb-353">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-354">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-355">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-356">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-357">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-357">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="95bbb-358">fin: date | [Fois](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="95bbb-358">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="95bbb-359">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-359">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="95bbb-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-362">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-362">Read mode</span></span>

<span data-ttu-id="95bbb-363">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-363">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-364">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-364">Compose mode</span></span>

<span data-ttu-id="95bbb-365">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-365">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="95bbb-366">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-366">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="95bbb-367">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-367">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="95bbb-368">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-368">Type</span></span>

*   <span data-ttu-id="95bbb-369">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="95bbb-369">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-370">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-370">Requirements</span></span>

|<span data-ttu-id="95bbb-371">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-371">Requirement</span></span>|<span data-ttu-id="95bbb-372">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-372">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-373">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-373">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-374">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-374">1.0</span></span>|
|[<span data-ttu-id="95bbb-375">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-375">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-376">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-376">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-377">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-377">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-378">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-378">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="95bbb-379">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="95bbb-379">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="95bbb-380">Obtient ou définit les emplacements d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-380">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-381">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-381">Read mode</span></span>

<span data-ttu-id="95bbb-382">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui vous permet d’obtenir l’ensemble des emplacements (chacun représenté par un objet [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associé au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-382">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-383">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-383">Compose mode</span></span>

<span data-ttu-id="95bbb-384">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui fournit des méthodes pour obtenir, supprimer ou ajouter des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-385">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-385">Type</span></span>

*   [<span data-ttu-id="95bbb-386">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="95bbb-386">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="95bbb-387">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-387">Requirements</span></span>

|<span data-ttu-id="95bbb-388">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-388">Requirement</span></span>|<span data-ttu-id="95bbb-389">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-390">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-391">Aperçu</span><span class="sxs-lookup"><span data-stu-id="95bbb-391">Preview</span></span>|
|[<span data-ttu-id="95bbb-392">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-393">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-394">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-395">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-395">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-396">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-396">Example</span></span>

<span data-ttu-id="95bbb-397">L’exemple suivant obtient les emplacements actuels associés au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-397">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="95bbb-398">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="95bbb-398">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="95bbb-399">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="95bbb-399">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="95bbb-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-402">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-402">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-403">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-403">Read mode</span></span>

<span data-ttu-id="95bbb-404">La `from` propriété renvoie un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="95bbb-404">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-405">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-405">Compose mode</span></span>

<span data-ttu-id="95bbb-406">La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.</span><span class="sxs-lookup"><span data-stu-id="95bbb-406">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="95bbb-407">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-407">Type</span></span>

*   <span data-ttu-id="95bbb-408">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [à partir de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="95bbb-408">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-409">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-409">Requirements</span></span>

|<span data-ttu-id="95bbb-410">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-410">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="95bbb-411">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-412">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-412">1.0</span></span>|<span data-ttu-id="95bbb-413">1.7</span><span class="sxs-lookup"><span data-stu-id="95bbb-413">1.7</span></span>|
|[<span data-ttu-id="95bbb-414">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-415">ReadItem</span></span>|<span data-ttu-id="95bbb-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-416">ReadWriteItem</span></span>|
|[<span data-ttu-id="95bbb-417">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-418">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-418">Read</span></span>|<span data-ttu-id="95bbb-419">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-419">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="95bbb-420">internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="95bbb-420">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="95bbb-421">Obtient ou définit les en-têtes Internet d’un message.</span><span class="sxs-lookup"><span data-stu-id="95bbb-421">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-422">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-422">Type</span></span>

*   [<span data-ttu-id="95bbb-423">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="95bbb-423">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="95bbb-424">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-424">Requirements</span></span>

|<span data-ttu-id="95bbb-425">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-425">Requirement</span></span>|<span data-ttu-id="95bbb-426">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-427">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-428">Aperçu</span><span class="sxs-lookup"><span data-stu-id="95bbb-428">Preview</span></span>|
|[<span data-ttu-id="95bbb-429">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-430">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-431">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-432">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-432">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-433">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-433">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="95bbb-434">internetMessageId: chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-434">internetMessageId: String</span></span>

<span data-ttu-id="95bbb-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-437">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-437">Type</span></span>

*   <span data-ttu-id="95bbb-438">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-438">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-439">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-439">Requirements</span></span>

|<span data-ttu-id="95bbb-440">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-440">Requirement</span></span>|<span data-ttu-id="95bbb-441">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-442">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-443">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-443">1.0</span></span>|
|[<span data-ttu-id="95bbb-444">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-445">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-446">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-447">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-447">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-448">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-448">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="95bbb-449">itemClass: chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-449">itemClass: String</span></span>

<span data-ttu-id="95bbb-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="95bbb-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="95bbb-454">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-454">Type</span></span>|<span data-ttu-id="95bbb-455">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-455">Description</span></span>|<span data-ttu-id="95bbb-456">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="95bbb-456">item class</span></span>|
|---|---|---|
|<span data-ttu-id="95bbb-457">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="95bbb-457">Appointment items</span></span>|<span data-ttu-id="95bbb-458">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-458">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="95bbb-459">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="95bbb-459">Message items</span></span>|<span data-ttu-id="95bbb-460">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="95bbb-460">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="95bbb-461">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-461">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-462">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-462">Type</span></span>

*   <span data-ttu-id="95bbb-463">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-464">Requirements</span></span>

|<span data-ttu-id="95bbb-465">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-465">Requirement</span></span>|<span data-ttu-id="95bbb-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-468">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-468">1.0</span></span>|
|[<span data-ttu-id="95bbb-469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-470">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-472">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-473">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-473">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="95bbb-474">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="95bbb-474">(nullable) itemId: String</span></span>

<span data-ttu-id="95bbb-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-477">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="95bbb-477">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="95bbb-478">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="95bbb-478">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="95bbb-479">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="95bbb-479">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="95bbb-480">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="95bbb-480">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="95bbb-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-483">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-483">Type</span></span>

*   <span data-ttu-id="95bbb-484">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-484">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-485">Requirements</span></span>

|<span data-ttu-id="95bbb-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-486">Requirement</span></span>|<span data-ttu-id="95bbb-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-489">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-489">1.0</span></span>|
|[<span data-ttu-id="95bbb-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-491">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-493">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-493">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-494">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-494">Example</span></span>

<span data-ttu-id="95bbb-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="95bbb-497">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="95bbb-497">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="95bbb-498">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="95bbb-498">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="95bbb-499">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-499">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-500">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-500">Type</span></span>

*   [<span data-ttu-id="95bbb-501">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="95bbb-501">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="95bbb-502">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-502">Requirements</span></span>

|<span data-ttu-id="95bbb-503">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-503">Requirement</span></span>|<span data-ttu-id="95bbb-504">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-505">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-506">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-506">1.0</span></span>|
|[<span data-ttu-id="95bbb-507">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-508">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-509">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-510">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-510">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-511">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-511">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="95bbb-512">Location: String | [Emplacement](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="95bbb-512">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="95bbb-513">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-513">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-514">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-514">Read mode</span></span>

<span data-ttu-id="95bbb-515">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-515">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-516">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-516">Compose mode</span></span>

<span data-ttu-id="95bbb-517">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-517">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="95bbb-518">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-518">Type</span></span>

*   <span data-ttu-id="95bbb-519">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="95bbb-519">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-520">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-520">Requirements</span></span>

|<span data-ttu-id="95bbb-521">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-521">Requirement</span></span>|<span data-ttu-id="95bbb-522">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-522">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-523">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-524">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-524">1.0</span></span>|
|[<span data-ttu-id="95bbb-525">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-526">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-527">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-528">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-528">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="95bbb-529">normalizedSubject: chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-529">normalizedSubject: String</span></span>

<span data-ttu-id="95bbb-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="95bbb-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="95bbb-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-534">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-534">Type</span></span>

*   <span data-ttu-id="95bbb-535">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-535">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-536">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-536">Requirements</span></span>

|<span data-ttu-id="95bbb-537">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-537">Requirement</span></span>|<span data-ttu-id="95bbb-538">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-539">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-540">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-540">1.0</span></span>|
|[<span data-ttu-id="95bbb-541">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-542">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-543">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-544">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-545">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-545">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="95bbb-546">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="95bbb-546">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="95bbb-547">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-547">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-548">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-548">Type</span></span>

*   [<span data-ttu-id="95bbb-549">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="95bbb-549">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="95bbb-550">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-550">Requirements</span></span>

|<span data-ttu-id="95bbb-551">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-551">Requirement</span></span>|<span data-ttu-id="95bbb-552">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-553">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-554">1.3</span><span class="sxs-lookup"><span data-stu-id="95bbb-554">1.3</span></span>|
|[<span data-ttu-id="95bbb-555">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-556">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-557">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-558">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-558">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-559">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-559">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="95bbb-560">optionalAttendees: les destinataires[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) Array. <</span><span class="sxs-lookup"><span data-stu-id="95bbb-560">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="95bbb-561">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-561">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="95bbb-562">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="95bbb-562">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-563">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-563">Read mode</span></span>

<span data-ttu-id="95bbb-564">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-564">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-565">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-565">Compose mode</span></span>

<span data-ttu-id="95bbb-566">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-566">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="95bbb-567">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-567">Type</span></span>

*   <span data-ttu-id="95bbb-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="95bbb-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-569">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-569">Requirements</span></span>

|<span data-ttu-id="95bbb-570">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-570">Requirement</span></span>|<span data-ttu-id="95bbb-571">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-571">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-572">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-572">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-573">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-573">1.0</span></span>|
|[<span data-ttu-id="95bbb-574">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-574">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-575">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-575">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-576">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-576">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-577">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-577">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="95bbb-578">Organisateur: [](/javascript/api/outlook/office.emailaddressdetails)|[organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="95bbb-578">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="95bbb-579">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="95bbb-579">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-580">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-580">Read mode</span></span>

<span data-ttu-id="95bbb-581">La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-581">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-582">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-582">Compose mode</span></span>

<span data-ttu-id="95bbb-583">La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-583">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="95bbb-584">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-584">Type</span></span>

*   <span data-ttu-id="95bbb-585">[](/javascript/api/outlook/office.emailaddressdetails) | [Organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="95bbb-585">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-586">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-586">Requirements</span></span>

|<span data-ttu-id="95bbb-587">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-587">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="95bbb-588">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-589">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-589">1.0</span></span>|<span data-ttu-id="95bbb-590">1.7</span><span class="sxs-lookup"><span data-stu-id="95bbb-590">1.7</span></span>|
|[<span data-ttu-id="95bbb-591">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-591">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-592">ReadItem</span></span>|<span data-ttu-id="95bbb-593">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-593">ReadWriteItem</span></span>|
|[<span data-ttu-id="95bbb-594">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-595">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-595">Read</span></span>|<span data-ttu-id="95bbb-596">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-596">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="95bbb-597">(Nullable) récurrence: [périodicité](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="95bbb-597">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="95bbb-598">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-598">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="95bbb-599">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-599">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="95bbb-600">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-600">Read and compose modes for appointment items.</span></span> <span data-ttu-id="95bbb-601">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-601">Read mode for meeting request items.</span></span>

<span data-ttu-id="95bbb-602">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="95bbb-602">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="95bbb-603">`null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="95bbb-603">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="95bbb-604">`undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-604">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="95bbb-605">Remarque: les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="95bbb-605">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="95bbb-606">Remarque: si l’objet de périodicité `null`est, cela indique que l’objet est un rendez-vous unique ou une demande de réunion d’un seul rendez-vous et non d’une série.</span><span class="sxs-lookup"><span data-stu-id="95bbb-606">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-607">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-607">Read mode</span></span>

<span data-ttu-id="95bbb-608">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-608">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="95bbb-609">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-609">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-610">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-610">Compose mode</span></span>

<span data-ttu-id="95bbb-611">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-611">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="95bbb-612">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-612">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="95bbb-613">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-613">Type</span></span>

* [<span data-ttu-id="95bbb-614">Instances</span><span class="sxs-lookup"><span data-stu-id="95bbb-614">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="95bbb-615">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-615">Requirement</span></span>|<span data-ttu-id="95bbb-616">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-617">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-618">1.7</span><span class="sxs-lookup"><span data-stu-id="95bbb-618">1.7</span></span>|
|[<span data-ttu-id="95bbb-619">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-619">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-620">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-621">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-621">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-622">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-622">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="95bbb-623">requiredAttendees: les destinataires[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) Array. <</span><span class="sxs-lookup"><span data-stu-id="95bbb-623">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="95bbb-624">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-624">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="95bbb-625">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="95bbb-625">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-626">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-626">Read mode</span></span>

<span data-ttu-id="95bbb-627">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-627">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-628">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-628">Compose mode</span></span>

<span data-ttu-id="95bbb-629">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-629">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="95bbb-630">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-630">Type</span></span>

*   <span data-ttu-id="95bbb-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="95bbb-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-632">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-632">Requirements</span></span>

|<span data-ttu-id="95bbb-633">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-633">Requirement</span></span>|<span data-ttu-id="95bbb-634">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-635">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-636">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-636">1.0</span></span>|
|[<span data-ttu-id="95bbb-637">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-638">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-639">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-640">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-640">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="95bbb-641">expéditeur: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="95bbb-641">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="95bbb-p128">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="95bbb-p129">Les propriétés [`from`](#from-emailaddressdetailsfrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-646">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-646">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-647">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-647">Type</span></span>

*   [<span data-ttu-id="95bbb-648">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="95bbb-648">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="95bbb-649">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-649">Requirements</span></span>

|<span data-ttu-id="95bbb-650">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-650">Requirement</span></span>|<span data-ttu-id="95bbb-651">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-651">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-652">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-653">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-653">1.0</span></span>|
|[<span data-ttu-id="95bbb-654">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-655">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-656">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-657">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-657">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-658">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-658">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="95bbb-659">(Nullable) seriesId: chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-659">(nullable) seriesId: String</span></span>

<span data-ttu-id="95bbb-660">Obtient l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="95bbb-660">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="95bbb-661">Dans OWA et Outlook, le `seriesId` renvoie l’ID des services Web Exchange (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="95bbb-661">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="95bbb-662">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="95bbb-662">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-663">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="95bbb-663">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="95bbb-664">La `seriesId` propriété n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="95bbb-664">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="95bbb-665">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="95bbb-665">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="95bbb-666">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="95bbb-666">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="95bbb-667">La `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="95bbb-667">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="95bbb-668">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-668">Type</span></span>

* <span data-ttu-id="95bbb-669">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-669">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-670">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-670">Requirements</span></span>

|<span data-ttu-id="95bbb-671">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-671">Requirement</span></span>|<span data-ttu-id="95bbb-672">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-673">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-674">1.7</span><span class="sxs-lookup"><span data-stu-id="95bbb-674">1.7</span></span>|
|[<span data-ttu-id="95bbb-675">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-676">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-677">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-678">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-678">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-679">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-679">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="95bbb-680">début: date | [Fois](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="95bbb-680">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="95bbb-681">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-681">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="95bbb-p132">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-684">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-684">Read mode</span></span>

<span data-ttu-id="95bbb-685">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-685">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-686">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-686">Compose mode</span></span>

<span data-ttu-id="95bbb-687">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-687">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="95bbb-688">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-688">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="95bbb-689">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-689">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="95bbb-690">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-690">Type</span></span>

*   <span data-ttu-id="95bbb-691">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="95bbb-691">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-692">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-692">Requirements</span></span>

|<span data-ttu-id="95bbb-693">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-693">Requirement</span></span>|<span data-ttu-id="95bbb-694">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-695">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-696">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-696">1.0</span></span>|
|[<span data-ttu-id="95bbb-697">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-698">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-699">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-700">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-700">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="95bbb-701">Subject: String | [Objet](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="95bbb-701">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="95bbb-702">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-702">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="95bbb-703">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="95bbb-703">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-704">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-704">Read mode</span></span>

<span data-ttu-id="95bbb-p133">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="95bbb-707">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="95bbb-707">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-708">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-708">Compose mode</span></span>
<span data-ttu-id="95bbb-709">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="95bbb-709">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="95bbb-710">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-710">Type</span></span>

*   <span data-ttu-id="95bbb-711">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="95bbb-711">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-712">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-712">Requirements</span></span>

|<span data-ttu-id="95bbb-713">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-713">Requirement</span></span>|<span data-ttu-id="95bbb-714">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-715">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-716">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-716">1.0</span></span>|
|[<span data-ttu-id="95bbb-717">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-718">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-719">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-720">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-720">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="95bbb-721">to: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="95bbb-721">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="95bbb-722">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="95bbb-722">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="95bbb-723">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="95bbb-723">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="95bbb-724">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-724">Read mode</span></span>

<span data-ttu-id="95bbb-p135">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="95bbb-727">Mode composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-727">Compose mode</span></span>

<span data-ttu-id="95bbb-728">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="95bbb-728">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="95bbb-729">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-729">Type</span></span>

*   <span data-ttu-id="95bbb-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="95bbb-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-731">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-731">Requirements</span></span>

|<span data-ttu-id="95bbb-732">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-732">Requirement</span></span>|<span data-ttu-id="95bbb-733">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-734">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-735">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-735">1.0</span></span>|
|[<span data-ttu-id="95bbb-736">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-737">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-738">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-739">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-739">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="95bbb-740">Méthodes</span><span class="sxs-lookup"><span data-stu-id="95bbb-740">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="95bbb-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="95bbb-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="95bbb-742">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="95bbb-742">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="95bbb-743">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="95bbb-743">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="95bbb-744">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="95bbb-744">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-745">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-745">Parameters</span></span>
|<span data-ttu-id="95bbb-746">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-746">Name</span></span>|<span data-ttu-id="95bbb-747">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-747">Type</span></span>|<span data-ttu-id="95bbb-748">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-748">Attributes</span></span>|<span data-ttu-id="95bbb-749">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-749">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="95bbb-750">Chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-750">String</span></span>||<span data-ttu-id="95bbb-p136">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="95bbb-753">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-753">String</span></span>||<span data-ttu-id="95bbb-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="95bbb-756">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-756">Object</span></span>|<span data-ttu-id="95bbb-757">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-757">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-758">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-758">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-759">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-759">Object</span></span>|<span data-ttu-id="95bbb-760">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-760">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-761">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-761">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="95bbb-762">Boolean</span><span class="sxs-lookup"><span data-stu-id="95bbb-762">Boolean</span></span>|<span data-ttu-id="95bbb-763">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-763">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-764">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-764">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="95bbb-765">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-765">function</span></span>|<span data-ttu-id="95bbb-766">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-766">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-767">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-767">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="95bbb-768">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-768">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="95bbb-769">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-769">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="95bbb-770">Erreurs</span><span class="sxs-lookup"><span data-stu-id="95bbb-770">Errors</span></span>

|<span data-ttu-id="95bbb-771">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="95bbb-771">Error code</span></span>|<span data-ttu-id="95bbb-772">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-772">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="95bbb-773">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="95bbb-773">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="95bbb-774">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="95bbb-774">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="95bbb-775">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-775">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-776">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-776">Requirements</span></span>

|<span data-ttu-id="95bbb-777">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-777">Requirement</span></span>|<span data-ttu-id="95bbb-778">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-778">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-779">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-779">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-780">1.1</span><span class="sxs-lookup"><span data-stu-id="95bbb-780">1.1</span></span>|
|[<span data-ttu-id="95bbb-781">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-781">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-782">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-782">ReadWriteItem</span></span>|
|[<span data-ttu-id="95bbb-783">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-783">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-784">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-784">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="95bbb-785">Exemples</span><span class="sxs-lookup"><span data-stu-id="95bbb-785">Examples</span></span>

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

<span data-ttu-id="95bbb-786">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="95bbb-786">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="95bbb-787">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="95bbb-787">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="95bbb-788">Ajoute un fichier à partir du codage Base64 à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="95bbb-788">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="95bbb-789">La `addFileAttachmentFromBase64Async` méthode charge le fichier à partir du codage Base64 et l’associe à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="95bbb-789">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="95bbb-790">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="95bbb-790">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="95bbb-791">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="95bbb-791">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-792">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-792">Parameters</span></span>

|<span data-ttu-id="95bbb-793">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-793">Name</span></span>|<span data-ttu-id="95bbb-794">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-794">Type</span></span>|<span data-ttu-id="95bbb-795">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-795">Attributes</span></span>|<span data-ttu-id="95bbb-796">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-796">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="95bbb-797">Chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-797">String</span></span>||<span data-ttu-id="95bbb-798">Contenu encodé en base64 d’une image ou d’un fichier à ajouter à un message électronique ou à un événement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-798">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="95bbb-799">Chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-799">String</span></span>||<span data-ttu-id="95bbb-p139">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="95bbb-802">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-802">Object</span></span>|<span data-ttu-id="95bbb-803">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-803">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-804">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-804">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-805">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-805">Object</span></span>|<span data-ttu-id="95bbb-806">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-806">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-807">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-807">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="95bbb-808">Boolean</span><span class="sxs-lookup"><span data-stu-id="95bbb-808">Boolean</span></span>|<span data-ttu-id="95bbb-809">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-809">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-810">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-810">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="95bbb-811">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-811">function</span></span>|<span data-ttu-id="95bbb-812">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-812">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-813">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-813">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="95bbb-814">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-814">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="95bbb-815">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-815">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="95bbb-816">Erreurs</span><span class="sxs-lookup"><span data-stu-id="95bbb-816">Errors</span></span>

|<span data-ttu-id="95bbb-817">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="95bbb-817">Error code</span></span>|<span data-ttu-id="95bbb-818">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-818">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="95bbb-819">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="95bbb-819">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="95bbb-820">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="95bbb-820">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="95bbb-821">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-821">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-822">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-822">Requirements</span></span>

|<span data-ttu-id="95bbb-823">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-823">Requirement</span></span>|<span data-ttu-id="95bbb-824">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-825">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-826">Aperçu</span><span class="sxs-lookup"><span data-stu-id="95bbb-826">Preview</span></span>|
|[<span data-ttu-id="95bbb-827">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-827">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-828">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-828">ReadWriteItem</span></span>|
|[<span data-ttu-id="95bbb-829">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-829">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-830">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-830">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="95bbb-831">Exemples</span><span class="sxs-lookup"><span data-stu-id="95bbb-831">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="95bbb-832">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="95bbb-832">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="95bbb-833">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="95bbb-833">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="95bbb-834">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="95bbb-834">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-835">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-835">Parameters</span></span>

| <span data-ttu-id="95bbb-836">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-836">Name</span></span> | <span data-ttu-id="95bbb-837">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-837">Type</span></span> | <span data-ttu-id="95bbb-838">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-838">Attributes</span></span> | <span data-ttu-id="95bbb-839">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-839">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="95bbb-840">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="95bbb-840">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="95bbb-841">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="95bbb-841">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="95bbb-842">Fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-842">Function</span></span> || <span data-ttu-id="95bbb-p140">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="95bbb-846">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-846">Object</span></span> | <span data-ttu-id="95bbb-847">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-847">&lt;optional&gt;</span></span> | <span data-ttu-id="95bbb-848">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-848">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="95bbb-849">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-849">Object</span></span> | <span data-ttu-id="95bbb-850">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-850">&lt;optional&gt;</span></span> | <span data-ttu-id="95bbb-851">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-851">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="95bbb-852">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-852">function</span></span>| <span data-ttu-id="95bbb-853">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-853">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-854">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-854">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-855">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-855">Requirements</span></span>

|<span data-ttu-id="95bbb-856">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-856">Requirement</span></span>| <span data-ttu-id="95bbb-857">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-857">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-858">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-858">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95bbb-859">1.7</span><span class="sxs-lookup"><span data-stu-id="95bbb-859">1.7</span></span> |
|[<span data-ttu-id="95bbb-860">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-860">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95bbb-861">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-861">ReadItem</span></span> |
|[<span data-ttu-id="95bbb-862">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-862">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95bbb-863">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-863">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="95bbb-864">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-864">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="95bbb-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="95bbb-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="95bbb-866">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-866">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="95bbb-p141">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="95bbb-870">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="95bbb-870">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="95bbb-871">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="95bbb-871">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-872">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-872">Parameters</span></span>

|<span data-ttu-id="95bbb-873">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-873">Name</span></span>|<span data-ttu-id="95bbb-874">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-874">Type</span></span>|<span data-ttu-id="95bbb-875">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-875">Attributes</span></span>|<span data-ttu-id="95bbb-876">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-876">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="95bbb-877">Chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-877">String</span></span>||<span data-ttu-id="95bbb-p142">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="95bbb-880">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-880">String</span></span>||<span data-ttu-id="95bbb-881">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="95bbb-881">The subject of the item to be attached.</span></span> <span data-ttu-id="95bbb-882">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="95bbb-882">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="95bbb-883">Object</span><span class="sxs-lookup"><span data-stu-id="95bbb-883">Object</span></span>|<span data-ttu-id="95bbb-884">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-884">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-885">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-885">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-886">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-886">Object</span></span>|<span data-ttu-id="95bbb-887">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-887">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-888">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-888">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="95bbb-889">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-889">function</span></span>|<span data-ttu-id="95bbb-890">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-890">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-891">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-891">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="95bbb-892">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-892">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="95bbb-893">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-893">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="95bbb-894">Erreurs</span><span class="sxs-lookup"><span data-stu-id="95bbb-894">Errors</span></span>

|<span data-ttu-id="95bbb-895">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="95bbb-895">Error code</span></span>|<span data-ttu-id="95bbb-896">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-896">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="95bbb-897">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-897">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-898">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-898">Requirements</span></span>

|<span data-ttu-id="95bbb-899">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-899">Requirement</span></span>|<span data-ttu-id="95bbb-900">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-900">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-901">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-901">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-902">1.1</span><span class="sxs-lookup"><span data-stu-id="95bbb-902">1.1</span></span>|
|[<span data-ttu-id="95bbb-903">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-903">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-904">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-904">ReadWriteItem</span></span>|
|[<span data-ttu-id="95bbb-905">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-905">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-906">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-906">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-907">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-907">Example</span></span>

<span data-ttu-id="95bbb-908">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-908">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="95bbb-909">close()</span><span class="sxs-lookup"><span data-stu-id="95bbb-909">close()</span></span>

<span data-ttu-id="95bbb-910">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="95bbb-910">Closes the current item that is being composed.</span></span>

<span data-ttu-id="95bbb-p144">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-913">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-913">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="95bbb-914">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="95bbb-914">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-915">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-915">Requirements</span></span>

|<span data-ttu-id="95bbb-916">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-916">Requirement</span></span>|<span data-ttu-id="95bbb-917">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-917">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-918">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-918">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-919">1.3</span><span class="sxs-lookup"><span data-stu-id="95bbb-919">1.3</span></span>|
|[<span data-ttu-id="95bbb-920">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-920">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-921">Restreinte</span><span class="sxs-lookup"><span data-stu-id="95bbb-921">Restricted</span></span>|
|[<span data-ttu-id="95bbb-922">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-922">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-923">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-923">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="95bbb-924">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="95bbb-924">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="95bbb-925">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="95bbb-925">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-926">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-926">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="95bbb-927">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-927">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="95bbb-928">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="95bbb-928">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="95bbb-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-932">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-932">Parameters</span></span>

|<span data-ttu-id="95bbb-933">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-933">Name</span></span>|<span data-ttu-id="95bbb-934">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-934">Type</span></span>|<span data-ttu-id="95bbb-935">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-935">Attributes</span></span>|<span data-ttu-id="95bbb-936">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-936">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="95bbb-937">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="95bbb-937">String &#124; Object</span></span>||<span data-ttu-id="95bbb-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="95bbb-940">**OU**</span><span class="sxs-lookup"><span data-stu-id="95bbb-940">**OR**</span></span><br/><span data-ttu-id="95bbb-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="95bbb-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="95bbb-943">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-943">String</span></span>|<span data-ttu-id="95bbb-944">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-944">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="95bbb-947">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-947">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="95bbb-948">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-948">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-949">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-949">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="95bbb-950">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-950">String</span></span>||<span data-ttu-id="95bbb-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="95bbb-953">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-953">String</span></span>||<span data-ttu-id="95bbb-954">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="95bbb-954">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="95bbb-955">Chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-955">String</span></span>||<span data-ttu-id="95bbb-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="95bbb-958">Booléen</span><span class="sxs-lookup"><span data-stu-id="95bbb-958">Boolean</span></span>||<span data-ttu-id="95bbb-p151">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="95bbb-961">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-961">String</span></span>||<span data-ttu-id="95bbb-p152">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="95bbb-965">function</span><span class="sxs-lookup"><span data-stu-id="95bbb-965">function</span></span>|<span data-ttu-id="95bbb-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-966">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-967">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-968">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-968">Requirements</span></span>

|<span data-ttu-id="95bbb-969">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-969">Requirement</span></span>|<span data-ttu-id="95bbb-970">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-970">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-971">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-971">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-972">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-972">1.0</span></span>|
|[<span data-ttu-id="95bbb-973">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-973">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-974">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-974">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-975">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-975">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-976">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-976">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="95bbb-977">Exemples</span><span class="sxs-lookup"><span data-stu-id="95bbb-977">Examples</span></span>

<span data-ttu-id="95bbb-978">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-978">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="95bbb-979">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="95bbb-979">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="95bbb-980">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="95bbb-980">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="95bbb-981">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="95bbb-981">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="95bbb-982">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-982">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="95bbb-983">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-983">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="95bbb-984">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="95bbb-984">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="95bbb-985">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="95bbb-985">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-986">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-986">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="95bbb-987">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-987">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="95bbb-988">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="95bbb-988">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="95bbb-p153">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-992">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-992">Parameters</span></span>

|<span data-ttu-id="95bbb-993">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-993">Name</span></span>|<span data-ttu-id="95bbb-994">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-994">Type</span></span>|<span data-ttu-id="95bbb-995">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-995">Attributes</span></span>|<span data-ttu-id="95bbb-996">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-996">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="95bbb-997">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="95bbb-997">String &#124; Object</span></span>||<span data-ttu-id="95bbb-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="95bbb-1000">**OU**</span><span class="sxs-lookup"><span data-stu-id="95bbb-1000">**OR**</span></span><br/><span data-ttu-id="95bbb-p155">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="95bbb-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="95bbb-1003">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1003">String</span></span>|<span data-ttu-id="95bbb-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-p156">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="95bbb-1007">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1007">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="95bbb-1008">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1009">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1009">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="95bbb-1010">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1010">String</span></span>||<span data-ttu-id="95bbb-p157">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="95bbb-1013">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1013">String</span></span>||<span data-ttu-id="95bbb-1014">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1014">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="95bbb-1015">Chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-1015">String</span></span>||<span data-ttu-id="95bbb-p158">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="95bbb-1018">Booléen</span><span class="sxs-lookup"><span data-stu-id="95bbb-1018">Boolean</span></span>||<span data-ttu-id="95bbb-p159">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="95bbb-1021">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1021">String</span></span>||<span data-ttu-id="95bbb-p160">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="95bbb-1025">function</span><span class="sxs-lookup"><span data-stu-id="95bbb-1025">function</span></span>|<span data-ttu-id="95bbb-1026">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1026">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1027">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1027">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1028">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1028">Requirements</span></span>

|<span data-ttu-id="95bbb-1029">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1029">Requirement</span></span>|<span data-ttu-id="95bbb-1030">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1030">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1031">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1031">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1032">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-1032">1.0</span></span>|
|[<span data-ttu-id="95bbb-1033">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1033">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1034">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1034">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1035">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1035">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1036">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1036">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="95bbb-1037">Exemples</span><span class="sxs-lookup"><span data-stu-id="95bbb-1037">Examples</span></span>

<span data-ttu-id="95bbb-1038">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1038">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="95bbb-1039">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1039">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="95bbb-1040">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1040">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="95bbb-1041">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1041">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="95bbb-1042">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1042">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="95bbb-1043">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1043">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="95bbb-1044">getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="95bbb-1044">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="95bbb-1045">Obtient la pièce jointe spécifiée à partir d’un message ou d’un `AttachmentContent` rendez-vous et la renvoie en tant qu’objet.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1045">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="95bbb-1046">La `getAttachmentContentAsync` méthode obtient la pièce jointe avec l’identificateur spécifié à partir de l’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1046">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="95bbb-1047">Il est recommandé d’utiliser l’identificateur pour récupérer une pièce jointe dans la même session que l’attachmentIds a été récupérée avec l' `getAttachmentsAsync` appel ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="95bbb-1047">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="95bbb-1048">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1048">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="95bbb-1049">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1049">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1050">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1050">Parameters</span></span>

|<span data-ttu-id="95bbb-1051">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1051">Name</span></span>|<span data-ttu-id="95bbb-1052">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1052">Type</span></span>|<span data-ttu-id="95bbb-1053">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1053">Attributes</span></span>|<span data-ttu-id="95bbb-1054">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1054">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="95bbb-1055">Chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-1055">String</span></span>||<span data-ttu-id="95bbb-1056">Identificateur de la pièce jointe que vous souhaitez obtenir.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1056">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="95bbb-1057">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1057">Object</span></span>|<span data-ttu-id="95bbb-1058">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1058">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1059">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1059">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-1060">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1060">Object</span></span>|<span data-ttu-id="95bbb-1061">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1061">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1062">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1062">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="95bbb-1063">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-1063">function</span></span>|<span data-ttu-id="95bbb-1064">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1065">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1065">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1066">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1066">Requirements</span></span>

|<span data-ttu-id="95bbb-1067">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1067">Requirement</span></span>|<span data-ttu-id="95bbb-1068">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1069">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1070">Aperçu</span><span class="sxs-lookup"><span data-stu-id="95bbb-1070">Preview</span></span>|
|[<span data-ttu-id="95bbb-1071">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1072">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1073">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1074">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1074">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1075">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1075">Returns:</span></span>

<span data-ttu-id="95bbb-1076">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="95bbb-1076">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="95bbb-1077">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1077">Example</span></span>

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
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  if (result.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="95bbb-1078">getAttachmentsAsync ([options], [Rappel]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="95bbb-1078">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="95bbb-1079">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1079">Gets the item's attachments as an array.</span></span> <span data-ttu-id="95bbb-1080">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1080">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1081">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1081">Parameters</span></span>

|<span data-ttu-id="95bbb-1082">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1082">Name</span></span>|<span data-ttu-id="95bbb-1083">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1083">Type</span></span>|<span data-ttu-id="95bbb-1084">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1084">Attributes</span></span>|<span data-ttu-id="95bbb-1085">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1085">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="95bbb-1086">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1086">Object</span></span>|<span data-ttu-id="95bbb-1087">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1088">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1088">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-1089">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1089">Object</span></span>|<span data-ttu-id="95bbb-1090">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1091">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1091">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="95bbb-1092">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-1092">function</span></span>|<span data-ttu-id="95bbb-1093">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1093">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1094">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1094">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1095">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1095">Requirements</span></span>

|<span data-ttu-id="95bbb-1096">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1096">Requirement</span></span>|<span data-ttu-id="95bbb-1097">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1098">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1099">Aperçu</span><span class="sxs-lookup"><span data-stu-id="95bbb-1099">Preview</span></span>|
|[<span data-ttu-id="95bbb-1100">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1101">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1102">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1103">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-1103">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1104">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1104">Returns:</span></span>

<span data-ttu-id="95bbb-1105">Type: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="95bbb-1105">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="95bbb-1106">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1106">Example</span></span>

<span data-ttu-id="95bbb-1107">L’exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1107">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="95bbb-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="95bbb-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="95bbb-1109">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1109">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1110">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1110">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-1111">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1111">Requirements</span></span>

|<span data-ttu-id="95bbb-1112">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1112">Requirement</span></span>|<span data-ttu-id="95bbb-1113">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1114">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1115">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-1115">1.0</span></span>|
|[<span data-ttu-id="95bbb-1116">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1116">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1117">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1117">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1118">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1118">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1119">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1119">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1120">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1120">Returns:</span></span>

<span data-ttu-id="95bbb-1121">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="95bbb-1121">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="95bbb-1122">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1122">Example</span></span>

<span data-ttu-id="95bbb-1123">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1123">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="95bbb-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="95bbb-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="95bbb-1125">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1125">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1126">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1127">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1127">Parameters</span></span>

|<span data-ttu-id="95bbb-1128">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1128">Name</span></span>|<span data-ttu-id="95bbb-1129">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1129">Type</span></span>|<span data-ttu-id="95bbb-1130">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1130">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="95bbb-1131">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="95bbb-1131">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="95bbb-1132">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1132">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1133">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1133">Requirements</span></span>

|<span data-ttu-id="95bbb-1134">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1134">Requirement</span></span>|<span data-ttu-id="95bbb-1135">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1136">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1137">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-1137">1.0</span></span>|
|[<span data-ttu-id="95bbb-1138">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1139">Restreinte</span><span class="sxs-lookup"><span data-stu-id="95bbb-1139">Restricted</span></span>|
|[<span data-ttu-id="95bbb-1140">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1141">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1142">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1142">Returns:</span></span>

<span data-ttu-id="95bbb-1143">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1143">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="95bbb-1144">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1144">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="95bbb-1145">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1145">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="95bbb-1146">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1146">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="95bbb-1147">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="95bbb-1147">Value of `entityType`</span></span>|<span data-ttu-id="95bbb-1148">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="95bbb-1148">Type of objects in returned array</span></span>|<span data-ttu-id="95bbb-1149">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="95bbb-1149">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="95bbb-1150">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1150">String</span></span>|<span data-ttu-id="95bbb-1151">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="95bbb-1151">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="95bbb-1152">Contact</span><span class="sxs-lookup"><span data-stu-id="95bbb-1152">Contact</span></span>|<span data-ttu-id="95bbb-1153">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="95bbb-1153">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="95bbb-1154">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1154">String</span></span>|<span data-ttu-id="95bbb-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="95bbb-1155">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="95bbb-1156">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="95bbb-1156">MeetingSuggestion</span></span>|<span data-ttu-id="95bbb-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="95bbb-1157">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="95bbb-1158">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="95bbb-1158">PhoneNumber</span></span>|<span data-ttu-id="95bbb-1159">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="95bbb-1159">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="95bbb-1160">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="95bbb-1160">TaskSuggestion</span></span>|<span data-ttu-id="95bbb-1161">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="95bbb-1161">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="95bbb-1162">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1162">String</span></span>|<span data-ttu-id="95bbb-1163">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="95bbb-1163">**Restricted**</span></span>|

<span data-ttu-id="95bbb-1164">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="95bbb-1164">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="95bbb-1165">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1165">Example</span></span>

<span data-ttu-id="95bbb-1166">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1166">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="95bbb-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="95bbb-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="95bbb-1168">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1168">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1169">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1169">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="95bbb-1170">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1170">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1171">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1171">Parameters</span></span>

|<span data-ttu-id="95bbb-1172">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1172">Name</span></span>|<span data-ttu-id="95bbb-1173">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1173">Type</span></span>|<span data-ttu-id="95bbb-1174">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1174">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="95bbb-1175">Chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-1175">String</span></span>|<span data-ttu-id="95bbb-1176">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1176">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1177">Requirements</span></span>

|<span data-ttu-id="95bbb-1178">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1178">Requirement</span></span>|<span data-ttu-id="95bbb-1179">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1179">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1180">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1181">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-1181">1.0</span></span>|
|[<span data-ttu-id="95bbb-1182">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1183">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1185">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1185">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1186">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1186">Returns:</span></span>

<span data-ttu-id="95bbb-p164">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="95bbb-1189">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="95bbb-1189">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="95bbb-1190">getInitializationContextAsync ([options], [Rappel])</span><span class="sxs-lookup"><span data-stu-id="95bbb-1190">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="95bbb-1191">Obtient les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1191">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1192">Cette méthode est uniquement prise en charge par Outlook 2016 ou une version ultérieure sur Windows (versions «démarrer en un clic» ultérieures à 16.0.8413.1000) et Outlook sur le Web pour Office 365.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1192">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1193">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1193">Parameters</span></span>

|<span data-ttu-id="95bbb-1194">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1194">Name</span></span>|<span data-ttu-id="95bbb-1195">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1195">Type</span></span>|<span data-ttu-id="95bbb-1196">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1196">Attributes</span></span>|<span data-ttu-id="95bbb-1197">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1197">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="95bbb-1198">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1198">Object</span></span>|<span data-ttu-id="95bbb-1199">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1199">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1200">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1200">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-1201">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1201">Object</span></span>|<span data-ttu-id="95bbb-1202">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1202">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1203">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1203">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="95bbb-1204">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-1204">function</span></span>|<span data-ttu-id="95bbb-1205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1205">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1206">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1206">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="95bbb-1207">En cas de réussite, les données d’initialisation sont fournies `asyncResult.value` dans la propriété sous la forme d’une chaîne.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1207">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="95bbb-1208">S’il n’existe pas de contexte d’initialisation `asyncResult` , l’objet contient `Error` un objet dont `code` la propriété est `9020` définie sur `name` et sa propriété `GenericResponseError`est définie sur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1208">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1209">Requirements</span></span>

|<span data-ttu-id="95bbb-1210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1210">Requirement</span></span>|<span data-ttu-id="95bbb-1211">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1213">Aperçu</span><span class="sxs-lookup"><span data-stu-id="95bbb-1213">Preview</span></span>|
|[<span data-ttu-id="95bbb-1214">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1215">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1216">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1217">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-1218">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1218">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="95bbb-1219">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="95bbb-1219">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="95bbb-1220">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1220">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1221">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1221">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="95bbb-p165">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="95bbb-1225">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1225">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="95bbb-1226">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1226">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="95bbb-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-1230">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1230">Requirements</span></span>

|<span data-ttu-id="95bbb-1231">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1231">Requirement</span></span>|<span data-ttu-id="95bbb-1232">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1233">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1234">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-1234">1.0</span></span>|
|[<span data-ttu-id="95bbb-1235">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1236">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1238">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1238">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1239">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1239">Returns:</span></span>

<span data-ttu-id="95bbb-p167">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="95bbb-1242">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="95bbb-1242">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="95bbb-1243">Object</span><span class="sxs-lookup"><span data-stu-id="95bbb-1243">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="95bbb-1244">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1244">Example</span></span>

<span data-ttu-id="95bbb-1245">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1245">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="95bbb-1246">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="95bbb-1246">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="95bbb-1247">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1247">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1248">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1248">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="95bbb-1249">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1249">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="95bbb-p168">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1252">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1252">Parameters</span></span>

|<span data-ttu-id="95bbb-1253">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1253">Name</span></span>|<span data-ttu-id="95bbb-1254">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1254">Type</span></span>|<span data-ttu-id="95bbb-1255">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1255">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="95bbb-1256">Chaîne</span><span class="sxs-lookup"><span data-stu-id="95bbb-1256">String</span></span>|<span data-ttu-id="95bbb-1257">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1257">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1258">Requirements</span></span>

|<span data-ttu-id="95bbb-1259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1259">Requirement</span></span>|<span data-ttu-id="95bbb-1260">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1260">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1262">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-1262">1.0</span></span>|
|[<span data-ttu-id="95bbb-1263">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1263">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1264">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1265">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1265">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1266">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1266">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1267">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1267">Returns:</span></span>

<span data-ttu-id="95bbb-1268">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1268">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="95bbb-1269">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="95bbb-1269">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="95bbb-1270">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="95bbb-1270">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="95bbb-1271">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1271">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="95bbb-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="95bbb-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="95bbb-1273">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1273">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="95bbb-p169">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1276">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1276">Parameters</span></span>

|<span data-ttu-id="95bbb-1277">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1277">Name</span></span>|<span data-ttu-id="95bbb-1278">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1278">Type</span></span>|<span data-ttu-id="95bbb-1279">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1279">Attributes</span></span>|<span data-ttu-id="95bbb-1280">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1280">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="95bbb-1281">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="95bbb-1281">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="95bbb-p170">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="95bbb-1285">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1285">Object</span></span>|<span data-ttu-id="95bbb-1286">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1286">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1287">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1287">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-1288">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1288">Object</span></span>|<span data-ttu-id="95bbb-1289">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1289">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1290">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1290">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="95bbb-1291">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-1291">function</span></span>||<span data-ttu-id="95bbb-1292">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1292">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="95bbb-1293">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1293">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="95bbb-1294">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1294">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1295">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1295">Requirements</span></span>

|<span data-ttu-id="95bbb-1296">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1296">Requirement</span></span>|<span data-ttu-id="95bbb-1297">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1298">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1299">1.2</span><span class="sxs-lookup"><span data-stu-id="95bbb-1299">1.2</span></span>|
|[<span data-ttu-id="95bbb-1300">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1301">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1301">ReadWriteItem</span></span>|
|[<span data-ttu-id="95bbb-1302">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1303">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-1303">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1304">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1304">Returns:</span></span>

<span data-ttu-id="95bbb-1305">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1305">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="95bbb-1306">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="95bbb-1306">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="95bbb-1307">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1307">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="95bbb-1308">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1308">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="95bbb-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="95bbb-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="95bbb-1310">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1310">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="95bbb-1311">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1311">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1312">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1312">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-1313">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1313">Requirements</span></span>

|<span data-ttu-id="95bbb-1314">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1314">Requirement</span></span>|<span data-ttu-id="95bbb-1315">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1315">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1316">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1317">1.6</span><span class="sxs-lookup"><span data-stu-id="95bbb-1317">1.6</span></span>|
|[<span data-ttu-id="95bbb-1318">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1319">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1320">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1321">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1321">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1322">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1322">Returns:</span></span>

<span data-ttu-id="95bbb-1323">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="95bbb-1323">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="95bbb-1324">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1324">Example</span></span>

<span data-ttu-id="95bbb-1325">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1325">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="95bbb-1326">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="95bbb-1326">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="95bbb-p173">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="95bbb-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1329">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1329">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="95bbb-p174">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="95bbb-1333">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1333">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="95bbb-1334">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1334">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="95bbb-p175">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="95bbb-1338">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1338">Requirements</span></span>

|<span data-ttu-id="95bbb-1339">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1339">Requirement</span></span>|<span data-ttu-id="95bbb-1340">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1340">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1341">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1342">1.6</span><span class="sxs-lookup"><span data-stu-id="95bbb-1342">1.6</span></span>|
|[<span data-ttu-id="95bbb-1343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1344">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1346">Lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1346">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95bbb-1347">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1347">Returns:</span></span>

<span data-ttu-id="95bbb-p176">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="95bbb-1350">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1350">Example</span></span>

<span data-ttu-id="95bbb-1351">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1351">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="95bbb-1352">getSharedPropertiesAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="95bbb-1352">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="95bbb-1353">Obtient les propriétés du rendez-vous ou du message sélectionné dans un dossier partagé, un calendrier ou une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1353">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1354">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1354">Parameters</span></span>

|<span data-ttu-id="95bbb-1355">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1355">Name</span></span>|<span data-ttu-id="95bbb-1356">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1356">Type</span></span>|<span data-ttu-id="95bbb-1357">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1357">Attributes</span></span>|<span data-ttu-id="95bbb-1358">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1358">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="95bbb-1359">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1359">Object</span></span>|<span data-ttu-id="95bbb-1360">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1360">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1361">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1361">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-1362">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1362">Object</span></span>|<span data-ttu-id="95bbb-1363">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1363">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1364">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1364">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="95bbb-1365">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-1365">function</span></span>||<span data-ttu-id="95bbb-1366">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="95bbb-1367">Les propriétés partagées sont fournies sous [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) la forme d' `asyncResult.value` un objet dans la propriété.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1367">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="95bbb-1368">Cet objet peut être utilisé pour obtenir les propriétés partagées de l’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1368">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1369">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1369">Requirements</span></span>

|<span data-ttu-id="95bbb-1370">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1370">Requirement</span></span>|<span data-ttu-id="95bbb-1371">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1371">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1372">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1373">Aperçu</span><span class="sxs-lookup"><span data-stu-id="95bbb-1373">Preview</span></span>|
|[<span data-ttu-id="95bbb-1374">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1374">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1375">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1376">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1376">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1377">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-1378">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1378">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="95bbb-1379">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="95bbb-1379">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="95bbb-1380">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1380">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="95bbb-p178">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1384">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1384">Parameters</span></span>

|<span data-ttu-id="95bbb-1385">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1385">Name</span></span>|<span data-ttu-id="95bbb-1386">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1386">Type</span></span>|<span data-ttu-id="95bbb-1387">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1387">Attributes</span></span>|<span data-ttu-id="95bbb-1388">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1388">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="95bbb-1389">function</span><span class="sxs-lookup"><span data-stu-id="95bbb-1389">function</span></span>||<span data-ttu-id="95bbb-1390">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1390">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="95bbb-1391">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1391">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="95bbb-1392">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1392">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="95bbb-1393">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1393">Object</span></span>|<span data-ttu-id="95bbb-1394">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1394">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1395">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1395">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="95bbb-1396">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1396">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1397">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1397">Requirements</span></span>

|<span data-ttu-id="95bbb-1398">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1398">Requirement</span></span>|<span data-ttu-id="95bbb-1399">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1399">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1400">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1401">1.0</span><span class="sxs-lookup"><span data-stu-id="95bbb-1401">1.0</span></span>|
|[<span data-ttu-id="95bbb-1402">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1403">ReadItem</span></span>|
|[<span data-ttu-id="95bbb-1404">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1405">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1405">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-1406">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1406">Example</span></span>

<span data-ttu-id="95bbb-p181">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="95bbb-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="95bbb-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="95bbb-1411">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1411">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="95bbb-1412">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1412">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="95bbb-1413">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1413">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="95bbb-1414">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1414">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="95bbb-1415">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1415">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1416">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1416">Parameters</span></span>

|<span data-ttu-id="95bbb-1417">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1417">Name</span></span>|<span data-ttu-id="95bbb-1418">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1418">Type</span></span>|<span data-ttu-id="95bbb-1419">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1419">Attributes</span></span>|<span data-ttu-id="95bbb-1420">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1420">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="95bbb-1421">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1421">String</span></span>||<span data-ttu-id="95bbb-1422">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1422">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="95bbb-1423">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1423">Object</span></span>|<span data-ttu-id="95bbb-1424">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1424">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1425">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1425">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-1426">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1426">Object</span></span>|<span data-ttu-id="95bbb-1427">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1427">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1428">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1428">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="95bbb-1429">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-1429">function</span></span>|<span data-ttu-id="95bbb-1430">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1430">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1431">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1431">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="95bbb-1432">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1432">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="95bbb-1433">Erreurs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1433">Errors</span></span>

|<span data-ttu-id="95bbb-1434">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1434">Error code</span></span>|<span data-ttu-id="95bbb-1435">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1435">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="95bbb-1436">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1436">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1437">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1437">Requirements</span></span>

|<span data-ttu-id="95bbb-1438">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1438">Requirement</span></span>|<span data-ttu-id="95bbb-1439">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1439">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1440">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1441">1.1</span><span class="sxs-lookup"><span data-stu-id="95bbb-1441">1.1</span></span>|
|[<span data-ttu-id="95bbb-1442">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1443">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1443">ReadWriteItem</span></span>|
|[<span data-ttu-id="95bbb-1444">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1445">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-1445">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-1446">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1446">Example</span></span>

<span data-ttu-id="95bbb-1447">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="95bbb-1447">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="95bbb-1448">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="95bbb-1448">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="95bbb-1449">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1449">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="95bbb-1450">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1450">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1451">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1451">Parameters</span></span>

| <span data-ttu-id="95bbb-1452">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1452">Name</span></span> | <span data-ttu-id="95bbb-1453">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1453">Type</span></span> | <span data-ttu-id="95bbb-1454">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1454">Attributes</span></span> | <span data-ttu-id="95bbb-1455">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1455">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="95bbb-1456">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="95bbb-1456">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="95bbb-1457">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1457">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="95bbb-1458">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1458">Object</span></span> | <span data-ttu-id="95bbb-1459">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1459">&lt;optional&gt;</span></span> | <span data-ttu-id="95bbb-1460">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1460">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="95bbb-1461">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1461">Object</span></span> | <span data-ttu-id="95bbb-1462">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1462">&lt;optional&gt;</span></span> | <span data-ttu-id="95bbb-1463">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1463">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="95bbb-1464">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-1464">function</span></span>| <span data-ttu-id="95bbb-1465">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1466">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1466">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1467">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1467">Requirements</span></span>

|<span data-ttu-id="95bbb-1468">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1468">Requirement</span></span>| <span data-ttu-id="95bbb-1469">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1470">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95bbb-1471">1.7</span><span class="sxs-lookup"><span data-stu-id="95bbb-1471">1.7</span></span> |
|[<span data-ttu-id="95bbb-1472">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95bbb-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1473">ReadItem</span></span> |
|[<span data-ttu-id="95bbb-1474">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95bbb-1475">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="95bbb-1475">Compose or Read</span></span> |

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="95bbb-1476">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="95bbb-1476">saveAsync([options], callback)</span></span>

<span data-ttu-id="95bbb-1477">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1477">Asynchronously saves an item.</span></span>

<span data-ttu-id="95bbb-p183">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1481">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1481">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="95bbb-1482">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1482">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="95bbb-p185">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="95bbb-1486">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="95bbb-1486">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="95bbb-1487">Outlook pour Mac ne prend pas `saveAsync` en charge les réunions en mode composition.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1487">Outlook for Mac does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="95bbb-1488">En tant que tel `saveAsync` , l’appel dans ce scénario renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1488">As such, calling `saveAsync` in that scenario returns an error.</span></span> <span data-ttu-id="95bbb-1489">Consultez la rubrique [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide de l’API Office js](https://support.microsoft.com/help/4505745) pour obtenir une solution de contournement.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1489">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="95bbb-1490">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1490">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1491">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1491">Parameters</span></span>

|<span data-ttu-id="95bbb-1492">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1492">Name</span></span>|<span data-ttu-id="95bbb-1493">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1493">Type</span></span>|<span data-ttu-id="95bbb-1494">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1494">Attributes</span></span>|<span data-ttu-id="95bbb-1495">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1495">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="95bbb-1496">Object</span><span class="sxs-lookup"><span data-stu-id="95bbb-1496">Object</span></span>|<span data-ttu-id="95bbb-1497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1497">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1498">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1498">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-1499">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1499">Object</span></span>|<span data-ttu-id="95bbb-1500">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1500">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1501">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1501">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="95bbb-1502">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-1502">function</span></span>||<span data-ttu-id="95bbb-1503">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1503">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="95bbb-1504">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1504">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1505">Requirements</span></span>

|<span data-ttu-id="95bbb-1506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1506">Requirement</span></span>|<span data-ttu-id="95bbb-1507">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1509">1.3</span><span class="sxs-lookup"><span data-stu-id="95bbb-1509">1.3</span></span>|
|[<span data-ttu-id="95bbb-1510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1511">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1511">ReadWriteItem</span></span>|
|[<span data-ttu-id="95bbb-1512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1513">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-1513">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="95bbb-1514">範例</span><span class="sxs-lookup"><span data-stu-id="95bbb-1514">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="95bbb-p187">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="95bbb-1517">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="95bbb-1517">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="95bbb-1518">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1518">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="95bbb-p188">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95bbb-1522">Paramètres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1522">Parameters</span></span>

|<span data-ttu-id="95bbb-1523">Nom</span><span class="sxs-lookup"><span data-stu-id="95bbb-1523">Name</span></span>|<span data-ttu-id="95bbb-1524">Type</span><span class="sxs-lookup"><span data-stu-id="95bbb-1524">Type</span></span>|<span data-ttu-id="95bbb-1525">Attributs</span><span class="sxs-lookup"><span data-stu-id="95bbb-1525">Attributes</span></span>|<span data-ttu-id="95bbb-1526">Description</span><span class="sxs-lookup"><span data-stu-id="95bbb-1526">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="95bbb-1527">String</span><span class="sxs-lookup"><span data-stu-id="95bbb-1527">String</span></span>||<span data-ttu-id="95bbb-p189">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="95bbb-1531">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1531">Object</span></span>|<span data-ttu-id="95bbb-1532">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1532">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1533">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="95bbb-1534">Objet</span><span class="sxs-lookup"><span data-stu-id="95bbb-1534">Object</span></span>|<span data-ttu-id="95bbb-1535">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1535">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-1536">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="95bbb-1537">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="95bbb-1537">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="95bbb-1538">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="95bbb-1538">&lt;optional&gt;</span></span>|<span data-ttu-id="95bbb-p190">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="95bbb-p191">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="95bbb-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="95bbb-1543">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="95bbb-1543">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="95bbb-1544">fonction</span><span class="sxs-lookup"><span data-stu-id="95bbb-1544">function</span></span>||<span data-ttu-id="95bbb-1545">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="95bbb-1545">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95bbb-1546">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="95bbb-1546">Requirements</span></span>

|<span data-ttu-id="95bbb-1547">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95bbb-1547">Requirement</span></span>|<span data-ttu-id="95bbb-1548">Valeur</span><span class="sxs-lookup"><span data-stu-id="95bbb-1548">Value</span></span>|
|---|---|
|[<span data-ttu-id="95bbb-1549">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="95bbb-1549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="95bbb-1550">1.2</span><span class="sxs-lookup"><span data-stu-id="95bbb-1550">1.2</span></span>|
|[<span data-ttu-id="95bbb-1551">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="95bbb-1551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="95bbb-1552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="95bbb-1552">ReadWriteItem</span></span>|
|[<span data-ttu-id="95bbb-1553">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="95bbb-1553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="95bbb-1554">Composition</span><span class="sxs-lookup"><span data-stu-id="95bbb-1554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="95bbb-1555">Exemple</span><span class="sxs-lookup"><span data-stu-id="95bbb-1555">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
