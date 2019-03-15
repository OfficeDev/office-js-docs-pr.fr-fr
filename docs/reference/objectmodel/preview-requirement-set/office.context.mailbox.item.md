---
title: Office. Context. Mailbox. Item-Preview ensemble de conditions requises
description: ''
ms.date: 02/26/2019
localization_priority: Normal
ms.openlocfilehash: 83ebbad97df829b1ec64748ec6671ecf8f137496
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600304"
---
# <a name="item"></a><span data-ttu-id="76c97-102">élément</span><span class="sxs-lookup"><span data-stu-id="76c97-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="76c97-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="76c97-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="76c97-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="76c97-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-106">Requirements</span></span>

|<span data-ttu-id="76c97-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-107">Requirement</span></span>|<span data-ttu-id="76c97-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-110">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-110">1.0</span></span>|
|[<span data-ttu-id="76c97-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="76c97-112">Restricted</span></span>|
|[<span data-ttu-id="76c97-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="76c97-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="76c97-115">Members and methods</span></span>

| <span data-ttu-id="76c97-116">Membre	</span><span class="sxs-lookup"><span data-stu-id="76c97-116">Member</span></span> | <span data-ttu-id="76c97-117">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="76c97-118">attachments</span><span class="sxs-lookup"><span data-stu-id="76c97-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="76c97-119">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-119">Member</span></span> |
| [<span data-ttu-id="76c97-120">bcc</span><span class="sxs-lookup"><span data-stu-id="76c97-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="76c97-121">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-121">Member</span></span> |
| [<span data-ttu-id="76c97-122">body</span><span class="sxs-lookup"><span data-stu-id="76c97-122">body</span></span>](#body-body) | <span data-ttu-id="76c97-123">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-123">Member</span></span> |
| [<span data-ttu-id="76c97-124">cc</span><span class="sxs-lookup"><span data-stu-id="76c97-124">cc</span></span>](#cc-arrayemailaddressdetails) | <span data-ttu-id="76c97-125">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-125">Member</span></span> |
| [<span data-ttu-id="76c97-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="76c97-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="76c97-127">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-127">Member</span></span> |
| [<span data-ttu-id="76c97-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="76c97-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="76c97-129">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-129">Member</span></span> |
| [<span data-ttu-id="76c97-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="76c97-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="76c97-131">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-131">Member</span></span> |
| [<span data-ttu-id="76c97-132">end</span><span class="sxs-lookup"><span data-stu-id="76c97-132">end</span></span>](#end-datetime) | <span data-ttu-id="76c97-133">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-133">Member</span></span> |
| [<span data-ttu-id="76c97-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="76c97-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="76c97-135">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-135">Member</span></span> |
| [<span data-ttu-id="76c97-136">from</span><span class="sxs-lookup"><span data-stu-id="76c97-136">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="76c97-137">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-137">Member</span></span> |
| [<span data-ttu-id="76c97-138">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="76c97-138">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="76c97-139">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-139">Member</span></span> |
| [<span data-ttu-id="76c97-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="76c97-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="76c97-141">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-141">Member</span></span> |
| [<span data-ttu-id="76c97-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="76c97-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="76c97-143">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-143">Member</span></span> |
| [<span data-ttu-id="76c97-144">itemId</span><span class="sxs-lookup"><span data-stu-id="76c97-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="76c97-145">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-145">Member</span></span> |
| [<span data-ttu-id="76c97-146">itemType</span><span class="sxs-lookup"><span data-stu-id="76c97-146">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="76c97-147">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-147">Member</span></span> |
| [<span data-ttu-id="76c97-148">location</span><span class="sxs-lookup"><span data-stu-id="76c97-148">location</span></span>](#location-stringlocation) | <span data-ttu-id="76c97-149">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-149">Member</span></span> |
| [<span data-ttu-id="76c97-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="76c97-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="76c97-151">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-151">Member</span></span> |
| [<span data-ttu-id="76c97-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="76c97-152">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="76c97-153">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-153">Member</span></span> |
| [<span data-ttu-id="76c97-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="76c97-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetails) | <span data-ttu-id="76c97-155">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-155">Member</span></span> |
| [<span data-ttu-id="76c97-156">organizer</span><span class="sxs-lookup"><span data-stu-id="76c97-156">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="76c97-157">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-157">Member</span></span> |
| [<span data-ttu-id="76c97-158">recurrence</span><span class="sxs-lookup"><span data-stu-id="76c97-158">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="76c97-159">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-159">Member</span></span> |
| [<span data-ttu-id="76c97-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="76c97-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetails) | <span data-ttu-id="76c97-161">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-161">Member</span></span> |
| [<span data-ttu-id="76c97-162">sender</span><span class="sxs-lookup"><span data-stu-id="76c97-162">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="76c97-163">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-163">Member</span></span> |
| [<span data-ttu-id="76c97-164">seriesId</span><span class="sxs-lookup"><span data-stu-id="76c97-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="76c97-165">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-165">Member</span></span> |
| [<span data-ttu-id="76c97-166">start</span><span class="sxs-lookup"><span data-stu-id="76c97-166">start</span></span>](#start-datetime) | <span data-ttu-id="76c97-167">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-167">Member</span></span> |
| [<span data-ttu-id="76c97-168">subject</span><span class="sxs-lookup"><span data-stu-id="76c97-168">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="76c97-169">Member</span><span class="sxs-lookup"><span data-stu-id="76c97-169">Member</span></span> |
| [<span data-ttu-id="76c97-170">to</span><span class="sxs-lookup"><span data-stu-id="76c97-170">to</span></span>](#to-arrayemailaddressdetails) | <span data-ttu-id="76c97-171">Membre</span><span class="sxs-lookup"><span data-stu-id="76c97-171">Member</span></span> |
| [<span data-ttu-id="76c97-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="76c97-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-173">Method</span></span> |
| [<span data-ttu-id="76c97-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="76c97-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="76c97-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-175">Method</span></span> |
| [<span data-ttu-id="76c97-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="76c97-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-177">Method</span></span> |
| [<span data-ttu-id="76c97-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="76c97-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-179">Method</span></span> |
| [<span data-ttu-id="76c97-180">close</span><span class="sxs-lookup"><span data-stu-id="76c97-180">close</span></span>](#close) | <span data-ttu-id="76c97-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-181">Method</span></span> |
| [<span data-ttu-id="76c97-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="76c97-182">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="76c97-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-183">Method</span></span> |
| [<span data-ttu-id="76c97-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="76c97-184">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="76c97-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-185">Method</span></span> |
| [<span data-ttu-id="76c97-186">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="76c97-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-187">Method</span></span> |
| [<span data-ttu-id="76c97-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="76c97-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-189">Method</span></span> |
| [<span data-ttu-id="76c97-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="76c97-190">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="76c97-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-191">Method</span></span> |
| [<span data-ttu-id="76c97-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="76c97-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontact) | <span data-ttu-id="76c97-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-193">Method</span></span> |
| [<span data-ttu-id="76c97-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="76c97-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontact) | <span data-ttu-id="76c97-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-195">Method</span></span> |
| [<span data-ttu-id="76c97-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="76c97-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-197">Method</span></span> |
| [<span data-ttu-id="76c97-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="76c97-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="76c97-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-199">Method</span></span> |
| [<span data-ttu-id="76c97-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="76c97-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="76c97-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-201">Method</span></span> |
| [<span data-ttu-id="76c97-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="76c97-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-203">Method</span></span> |
| [<span data-ttu-id="76c97-204">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="76c97-204">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="76c97-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-205">Method</span></span> |
| [<span data-ttu-id="76c97-206">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="76c97-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="76c97-207">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-207">Method</span></span> |
| [<span data-ttu-id="76c97-208">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="76c97-209">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-209">Method</span></span> |
| [<span data-ttu-id="76c97-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="76c97-211">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-211">Method</span></span> |
| [<span data-ttu-id="76c97-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="76c97-213">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-213">Method</span></span> |
| [<span data-ttu-id="76c97-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="76c97-215">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-215">Method</span></span> |
| [<span data-ttu-id="76c97-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="76c97-217">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-217">Method</span></span> |
| [<span data-ttu-id="76c97-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="76c97-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="76c97-219">Méthode</span><span class="sxs-lookup"><span data-stu-id="76c97-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="76c97-220">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-220">Example</span></span>

<span data-ttu-id="76c97-221">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="76c97-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="76c97-222">Membres</span><span class="sxs-lookup"><span data-stu-id="76c97-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="76c97-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="76c97-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="76c97-224">Obtient les pièces jointes de l'élément sous la forme d'un tableau.</span><span class="sxs-lookup"><span data-stu-id="76c97-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="76c97-225">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-226">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="76c97-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="76c97-227">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="76c97-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-228">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-228">Type</span></span>

*   <span data-ttu-id="76c97-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="76c97-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-230">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-230">Requirements</span></span>

|<span data-ttu-id="76c97-231">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-231">Requirement</span></span>|<span data-ttu-id="76c97-232">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-233">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-234">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-234">1.0</span></span>|
|[<span data-ttu-id="76c97-235">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-235">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-236">ReadItem</span></span>|
|[<span data-ttu-id="76c97-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-237">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-238">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-239">Example</span></span>

<span data-ttu-id="76c97-240">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="76c97-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="76c97-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="76c97-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="76c97-242">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="76c97-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="76c97-243">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-244">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-244">Type</span></span>

*   [<span data-ttu-id="76c97-245">Destinataires</span><span class="sxs-lookup"><span data-stu-id="76c97-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="76c97-246">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-246">Requirements</span></span>

|<span data-ttu-id="76c97-247">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-247">Requirement</span></span>|<span data-ttu-id="76c97-248">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-249">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-250">1.1</span><span class="sxs-lookup"><span data-stu-id="76c97-250">1.1</span></span>|
|[<span data-ttu-id="76c97-251">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-252">ReadItem</span></span>|
|[<span data-ttu-id="76c97-253">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-254">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-255">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-255">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="76c97-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="76c97-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="76c97-257">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-258">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-258">Type</span></span>

*   [<span data-ttu-id="76c97-259">Body</span><span class="sxs-lookup"><span data-stu-id="76c97-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="76c97-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-260">Requirements</span></span>

|<span data-ttu-id="76c97-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-261">Requirement</span></span>|<span data-ttu-id="76c97-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-264">1.1</span><span class="sxs-lookup"><span data-stu-id="76c97-264">1.1</span></span>|
|[<span data-ttu-id="76c97-265">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-266">ReadItem</span></span>|
|[<span data-ttu-id="76c97-267">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-268">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-268">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-269">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-269">Example</span></span>

<span data-ttu-id="76c97-270">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="76c97-270">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="76c97-271">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-271">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="76c97-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="76c97-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="76c97-273">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="76c97-273">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="76c97-274">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="76c97-274">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-275">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-275">Read mode</span></span>

<span data-ttu-id="76c97-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="76c97-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-278">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-278">Compose mode</span></span>

<span data-ttu-id="76c97-279">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="76c97-279">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="76c97-280">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-280">Type</span></span>

*   <span data-ttu-id="76c97-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="76c97-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-282">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-282">Requirements</span></span>

|<span data-ttu-id="76c97-283">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-283">Requirement</span></span>|<span data-ttu-id="76c97-284">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-285">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-286">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-286">1.0</span></span>|
|[<span data-ttu-id="76c97-287">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-288">ReadItem</span></span>|
|[<span data-ttu-id="76c97-289">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-290">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-290">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="76c97-291">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="76c97-291">(nullable) conversationId :String</span></span>

<span data-ttu-id="76c97-292">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="76c97-292">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="76c97-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="76c97-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="76c97-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="76c97-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-297">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-297">Type</span></span>

*   <span data-ttu-id="76c97-298">String</span><span class="sxs-lookup"><span data-stu-id="76c97-298">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-299">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-299">Requirements</span></span>

|<span data-ttu-id="76c97-300">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-300">Requirement</span></span>|<span data-ttu-id="76c97-301">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-302">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-303">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-303">1.0</span></span>|
|[<span data-ttu-id="76c97-304">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-305">ReadItem</span></span>|
|[<span data-ttu-id="76c97-306">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-307">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-308">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-308">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="76c97-309">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="76c97-309">dateTimeCreated :Date</span></span>

<span data-ttu-id="76c97-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-312">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-312">Type</span></span>

*   <span data-ttu-id="76c97-313">Date</span><span class="sxs-lookup"><span data-stu-id="76c97-313">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-314">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-314">Requirements</span></span>

|<span data-ttu-id="76c97-315">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-315">Requirement</span></span>|<span data-ttu-id="76c97-316">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-316">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-317">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-318">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-318">1.0</span></span>|
|[<span data-ttu-id="76c97-319">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-319">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-320">ReadItem</span></span>|
|[<span data-ttu-id="76c97-321">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-321">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-322">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-322">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-323">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-323">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="76c97-324">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="76c97-324">dateTimeModified :Date</span></span>

<span data-ttu-id="76c97-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-327">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-327">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-328">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-328">Type</span></span>

*   <span data-ttu-id="76c97-329">Date</span><span class="sxs-lookup"><span data-stu-id="76c97-329">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-330">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-330">Requirements</span></span>

|<span data-ttu-id="76c97-331">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-331">Requirement</span></span>|<span data-ttu-id="76c97-332">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-333">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-334">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-334">1.0</span></span>|
|[<span data-ttu-id="76c97-335">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-335">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-336">ReadItem</span></span>|
|[<span data-ttu-id="76c97-337">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-337">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-338">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-339">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-339">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="76c97-340">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="76c97-340">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="76c97-341">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-341">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="76c97-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="76c97-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-344">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-344">Read mode</span></span>

<span data-ttu-id="76c97-345">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="76c97-345">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-346">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-346">Compose mode</span></span>

<span data-ttu-id="76c97-347">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="76c97-347">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="76c97-348">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="76c97-348">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="76c97-349">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="76c97-349">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="76c97-350">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-350">Type</span></span>

*   <span data-ttu-id="76c97-351">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="76c97-351">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-352">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-352">Requirements</span></span>

|<span data-ttu-id="76c97-353">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-353">Requirement</span></span>|<span data-ttu-id="76c97-354">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-355">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-356">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-356">1.0</span></span>|
|[<span data-ttu-id="76c97-357">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-358">ReadItem</span></span>|
|[<span data-ttu-id="76c97-359">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-360">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-360">Compose or Read</span></span>|

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="76c97-361">enhancedLocation:[enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="76c97-361">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="76c97-362">Obtient ou définit les emplacements d'un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-362">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-363">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-363">Read mode</span></span>

<span data-ttu-id="76c97-364">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui vous permet d'obtenir l'ensemble des emplacements (chacun représenté par un objet [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associé au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="76c97-365">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-365">Compose mode</span></span>

<span data-ttu-id="76c97-366">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui fournit des méthodes pour obtenir, supprimer ou ajouter des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-366">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-367">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-367">Type</span></span>

*   [<span data-ttu-id="76c97-368">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="76c97-368">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="76c97-369">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-369">Requirements</span></span>

|<span data-ttu-id="76c97-370">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-370">Requirement</span></span>|<span data-ttu-id="76c97-371">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-371">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-372">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-373">Aperçu</span><span class="sxs-lookup"><span data-stu-id="76c97-373">Preview</span></span>|
|[<span data-ttu-id="76c97-374">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-374">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-375">ReadItem</span></span>|
|[<span data-ttu-id="76c97-376">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-376">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-377">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-378">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-378">Example</span></span>

<span data-ttu-id="76c97-379">L'exemple suivant obtient les emplacements actuels associés au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-379">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="76c97-380">from:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="76c97-380">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="76c97-381">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="76c97-381">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="76c97-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="76c97-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-384">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="76c97-384">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-385">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-385">Read mode</span></span>

<span data-ttu-id="76c97-386">La `from` propriété renvoie un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="76c97-386">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-387">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-387">Compose mode</span></span>

<span data-ttu-id="76c97-388">La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.</span><span class="sxs-lookup"><span data-stu-id="76c97-388">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="76c97-389">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-389">Type</span></span>

*   <span data-ttu-id="76c97-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [à partir de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="76c97-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-391">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-391">Requirements</span></span>

|<span data-ttu-id="76c97-392">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-392">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="76c97-393">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-394">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-394">1.0</span></span>|<span data-ttu-id="76c97-395">1.7</span><span class="sxs-lookup"><span data-stu-id="76c97-395">1.7</span></span>|
|[<span data-ttu-id="76c97-396">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-397">ReadItem</span></span>|<span data-ttu-id="76c97-398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="76c97-398">ReadWriteItem</span></span>|
|[<span data-ttu-id="76c97-399">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-400">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-400">Read</span></span>|<span data-ttu-id="76c97-401">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-401">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="76c97-402">internetHeaders:[internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="76c97-402">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="76c97-403">Obtient ou définit les en-têtes Internet d'un message.</span><span class="sxs-lookup"><span data-stu-id="76c97-403">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-404">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-404">Type</span></span>

*   [<span data-ttu-id="76c97-405">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="76c97-405">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="76c97-406">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-406">Requirements</span></span>

|<span data-ttu-id="76c97-407">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-407">Requirement</span></span>|<span data-ttu-id="76c97-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-409">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-410">Aperçu</span><span class="sxs-lookup"><span data-stu-id="76c97-410">Preview</span></span>|
|[<span data-ttu-id="76c97-411">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-412">ReadItem</span></span>|
|[<span data-ttu-id="76c97-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-414">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-415">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="76c97-416">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="76c97-416">internetMessageId :String</span></span>

<span data-ttu-id="76c97-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-419">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-419">Type</span></span>

*   <span data-ttu-id="76c97-420">String</span><span class="sxs-lookup"><span data-stu-id="76c97-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-421">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-421">Requirements</span></span>

|<span data-ttu-id="76c97-422">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-422">Requirement</span></span>|<span data-ttu-id="76c97-423">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-425">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-425">1.0</span></span>|
|[<span data-ttu-id="76c97-426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-427">ReadItem</span></span>|
|[<span data-ttu-id="76c97-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-429">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-430">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-430">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

#### <a name="itemclass-string"></a><span data-ttu-id="76c97-431">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="76c97-431">itemClass :String</span></span>

<span data-ttu-id="76c97-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="76c97-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="76c97-436">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-436">Type</span></span>|<span data-ttu-id="76c97-437">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-437">Description</span></span>|<span data-ttu-id="76c97-438">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="76c97-438">item class</span></span>|
|---|---|---|
|<span data-ttu-id="76c97-439">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="76c97-439">Appointment items</span></span>|<span data-ttu-id="76c97-440">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="76c97-440">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="76c97-441">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="76c97-441">Message items</span></span>|<span data-ttu-id="76c97-442">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="76c97-442">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="76c97-443">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="76c97-443">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-444">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-444">Type</span></span>

*   <span data-ttu-id="76c97-445">String</span><span class="sxs-lookup"><span data-stu-id="76c97-445">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-446">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-446">Requirements</span></span>

|<span data-ttu-id="76c97-447">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-447">Requirement</span></span>|<span data-ttu-id="76c97-448">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-449">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-450">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-450">1.0</span></span>|
|[<span data-ttu-id="76c97-451">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-451">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-452">ReadItem</span></span>|
|[<span data-ttu-id="76c97-453">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-453">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-454">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-454">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-455">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-455">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="76c97-456">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="76c97-456">(nullable) itemId :String</span></span>

<span data-ttu-id="76c97-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-459">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="76c97-459">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="76c97-460">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="76c97-460">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="76c97-461">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="76c97-461">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="76c97-462">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="76c97-462">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="76c97-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-465">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-465">Type</span></span>

*   <span data-ttu-id="76c97-466">String</span><span class="sxs-lookup"><span data-stu-id="76c97-466">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-467">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-467">Requirements</span></span>

|<span data-ttu-id="76c97-468">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-468">Requirement</span></span>|<span data-ttu-id="76c97-469">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-470">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-471">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-471">1.0</span></span>|
|[<span data-ttu-id="76c97-472">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-473">ReadItem</span></span>|
|[<span data-ttu-id="76c97-474">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-475">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-475">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-476">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-476">Example</span></span>

<span data-ttu-id="76c97-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="76c97-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="76c97-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="76c97-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="76c97-480">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="76c97-480">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="76c97-481">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-481">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-482">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-482">Type</span></span>

*   [<span data-ttu-id="76c97-483">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="76c97-483">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="76c97-484">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-484">Requirements</span></span>

|<span data-ttu-id="76c97-485">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-485">Requirement</span></span>|<span data-ttu-id="76c97-486">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-487">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-488">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-488">1.0</span></span>|
|[<span data-ttu-id="76c97-489">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-490">ReadItem</span></span>|
|[<span data-ttu-id="76c97-491">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-492">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-492">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-493">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-493">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="76c97-494">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="76c97-494">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="76c97-495">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-495">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-496">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-496">Read mode</span></span>

<span data-ttu-id="76c97-497">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-497">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-498">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-498">Compose mode</span></span>

<span data-ttu-id="76c97-499">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-499">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="76c97-500">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-500">Type</span></span>

*   <span data-ttu-id="76c97-501">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="76c97-501">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-502">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-502">Requirements</span></span>

|<span data-ttu-id="76c97-503">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-503">Requirement</span></span>|<span data-ttu-id="76c97-504">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-505">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-506">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-506">1.0</span></span>|
|[<span data-ttu-id="76c97-507">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-508">ReadItem</span></span>|
|[<span data-ttu-id="76c97-509">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-510">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-510">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="76c97-511">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="76c97-511">normalizedSubject :String</span></span>

<span data-ttu-id="76c97-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="76c97-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="76c97-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-516">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-516">Type</span></span>

*   <span data-ttu-id="76c97-517">String</span><span class="sxs-lookup"><span data-stu-id="76c97-517">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-518">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-518">Requirements</span></span>

|<span data-ttu-id="76c97-519">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-519">Requirement</span></span>|<span data-ttu-id="76c97-520">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-521">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-522">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-522">1.0</span></span>|
|[<span data-ttu-id="76c97-523">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-524">ReadItem</span></span>|
|[<span data-ttu-id="76c97-525">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-526">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-526">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-527">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-527">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="76c97-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="76c97-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="76c97-529">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-529">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-530">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-530">Type</span></span>

*   [<span data-ttu-id="76c97-531">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="76c97-531">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="76c97-532">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-532">Requirements</span></span>

|<span data-ttu-id="76c97-533">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-533">Requirement</span></span>|<span data-ttu-id="76c97-534">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-535">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-536">1.3</span><span class="sxs-lookup"><span data-stu-id="76c97-536">1.3</span></span>|
|[<span data-ttu-id="76c97-537">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-537">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-538">ReadItem</span></span>|
|[<span data-ttu-id="76c97-539">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-539">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-540">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-540">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-541">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-541">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="76c97-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="76c97-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="76c97-543">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="76c97-543">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="76c97-544">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="76c97-544">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-545">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-545">Read mode</span></span>

<span data-ttu-id="76c97-546">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-546">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-547">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-547">Compose mode</span></span>

<span data-ttu-id="76c97-548">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-548">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="76c97-549">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-549">Type</span></span>

*   <span data-ttu-id="76c97-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="76c97-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-551">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-551">Requirements</span></span>

|<span data-ttu-id="76c97-552">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-552">Requirement</span></span>|<span data-ttu-id="76c97-553">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-553">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-554">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-554">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-555">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-555">1.0</span></span>|
|[<span data-ttu-id="76c97-556">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-556">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-557">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-557">ReadItem</span></span>|
|[<span data-ttu-id="76c97-558">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-558">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-559">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-559">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="76c97-560">Organisateur:[](/javascript/api/outlook/office.emailaddressdetails)|[organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="76c97-560">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="76c97-561">Obtient l'adresse de messagerie de l'organisateur d'une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="76c97-561">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-562">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-562">Read mode</span></span>

<span data-ttu-id="76c97-563">La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l'organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-563">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-564">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-564">Compose mode</span></span>

<span data-ttu-id="76c97-565">La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur de l'organisateur.</span><span class="sxs-lookup"><span data-stu-id="76c97-565">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="76c97-566">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-566">Type</span></span>

*   <span data-ttu-id="76c97-567">[](/javascript/api/outlook/office.emailaddressdetails) | [Organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="76c97-567">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-568">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-568">Requirements</span></span>

|<span data-ttu-id="76c97-569">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-569">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="76c97-570">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-571">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-571">1.0</span></span>|<span data-ttu-id="76c97-572">1.7</span><span class="sxs-lookup"><span data-stu-id="76c97-572">1.7</span></span>|
|[<span data-ttu-id="76c97-573">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-573">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-574">ReadItem</span></span>|<span data-ttu-id="76c97-575">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="76c97-575">ReadWriteItem</span></span>|
|[<span data-ttu-id="76c97-576">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-576">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-577">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-577">Read</span></span>|<span data-ttu-id="76c97-578">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-578">Compose</span></span>|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="76c97-579">(Nullable) récurrence:[périodicité](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="76c97-579">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="76c97-580">Obtient ou définit la périodicité d'un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-580">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="76c97-581">Obtient la périodicité d'une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-581">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="76c97-582">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-582">Read and compose modes for appointment items.</span></span> <span data-ttu-id="76c97-583">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-583">Read mode for meeting request items.</span></span>

<span data-ttu-id="76c97-584">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="76c97-584">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="76c97-585">`null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="76c97-585">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="76c97-586">`undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-586">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="76c97-587">Remarque: les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="76c97-587">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="76c97-588">Remarque: si l'objet de périodicité `null`est, cela indique que l'objet est un rendez-vous unique ou une demande de réunion d'un seul rendez-vous et non d'une série.</span><span class="sxs-lookup"><span data-stu-id="76c97-588">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-589">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-589">Read mode</span></span>

<span data-ttu-id="76c97-590">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-590">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="76c97-591">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-591">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-592">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-592">Compose mode</span></span>

<span data-ttu-id="76c97-593">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-593">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="76c97-594">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-594">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="76c97-595">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-595">Type</span></span>

* [<span data-ttu-id="76c97-596">Instances</span><span class="sxs-lookup"><span data-stu-id="76c97-596">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="76c97-597">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-597">Requirement</span></span>|<span data-ttu-id="76c97-598">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-599">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-600">1.7</span><span class="sxs-lookup"><span data-stu-id="76c97-600">1.7</span></span>|
|[<span data-ttu-id="76c97-601">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-601">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-602">ReadItem</span></span>|
|[<span data-ttu-id="76c97-603">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-603">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-604">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-604">Compose or Read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="76c97-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="76c97-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="76c97-606">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="76c97-606">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="76c97-607">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="76c97-607">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-608">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-608">Read mode</span></span>

<span data-ttu-id="76c97-609">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-609">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-610">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-610">Compose mode</span></span>

<span data-ttu-id="76c97-611">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-611">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="76c97-612">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-612">Type</span></span>

*   <span data-ttu-id="76c97-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="76c97-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-614">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-614">Requirements</span></span>

|<span data-ttu-id="76c97-615">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-615">Requirement</span></span>|<span data-ttu-id="76c97-616">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-617">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-618">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-618">1.0</span></span>|
|[<span data-ttu-id="76c97-619">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-619">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-620">ReadItem</span></span>|
|[<span data-ttu-id="76c97-621">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-621">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-622">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-622">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="76c97-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="76c97-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="76c97-p128">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="76c97-p129">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="76c97-p129">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-628">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="76c97-628">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-629">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-629">Type</span></span>

*   [<span data-ttu-id="76c97-630">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="76c97-630">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="76c97-631">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-631">Requirements</span></span>

|<span data-ttu-id="76c97-632">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-632">Requirement</span></span>|<span data-ttu-id="76c97-633">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-633">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-634">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-634">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-635">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-635">1.0</span></span>|
|[<span data-ttu-id="76c97-636">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-636">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-637">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-637">ReadItem</span></span>|
|[<span data-ttu-id="76c97-638">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-638">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-639">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-639">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-640">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-640">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="76c97-641">(Nullable) seriesId: chaîne</span><span class="sxs-lookup"><span data-stu-id="76c97-641">(nullable) seriesId :String</span></span>

<span data-ttu-id="76c97-642">Obtient l'ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="76c97-642">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="76c97-643">Dans OWA et Outlook, le `seriesId` renvoie l'ID des services Web Exchange (EWS) de l'élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="76c97-643">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="76c97-644">Toutefois, dans iOS et Android, le `seriesId` renvoie l'ID REST de l'élément parent.</span><span class="sxs-lookup"><span data-stu-id="76c97-644">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-645">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="76c97-645">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="76c97-646">La `seriesId` propriété n'est pas identique aux ID Outlook utilisés par l'API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="76c97-646">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="76c97-647">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="76c97-647">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="76c97-648">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="76c97-648">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="76c97-649">La `seriesId` propriété renvoie `null` pour les éléments qui n'ont pas d'éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="76c97-649">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="76c97-650">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-650">Type</span></span>

* <span data-ttu-id="76c97-651">String</span><span class="sxs-lookup"><span data-stu-id="76c97-651">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-652">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-652">Requirements</span></span>

|<span data-ttu-id="76c97-653">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-653">Requirement</span></span>|<span data-ttu-id="76c97-654">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-654">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-655">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-656">1.7</span><span class="sxs-lookup"><span data-stu-id="76c97-656">1.7</span></span>|
|[<span data-ttu-id="76c97-657">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-657">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-658">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-658">ReadItem</span></span>|
|[<span data-ttu-id="76c97-659">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-659">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-660">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-660">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-661">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-661">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="76c97-662">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="76c97-662">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="76c97-663">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-663">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="76c97-p132">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="76c97-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-666">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-666">Read mode</span></span>

<span data-ttu-id="76c97-667">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="76c97-667">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-668">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-668">Compose mode</span></span>

<span data-ttu-id="76c97-669">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="76c97-669">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="76c97-670">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="76c97-670">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="76c97-671">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="76c97-671">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="76c97-672">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-672">Type</span></span>

*   <span data-ttu-id="76c97-673">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="76c97-673">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-674">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-674">Requirements</span></span>

|<span data-ttu-id="76c97-675">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-675">Requirement</span></span>|<span data-ttu-id="76c97-676">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-676">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-677">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-677">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-678">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-678">1.0</span></span>|
|[<span data-ttu-id="76c97-679">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-679">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-680">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-680">ReadItem</span></span>|
|[<span data-ttu-id="76c97-681">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-681">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-682">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-682">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="76c97-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="76c97-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="76c97-684">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-684">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="76c97-685">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="76c97-685">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-686">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-686">Read mode</span></span>

<span data-ttu-id="76c97-p133">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="76c97-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="76c97-689">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="76c97-689">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-690">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-690">Compose mode</span></span>
<span data-ttu-id="76c97-691">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="76c97-691">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="76c97-692">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-692">Type</span></span>

*   <span data-ttu-id="76c97-693">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="76c97-693">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-694">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-694">Requirements</span></span>

|<span data-ttu-id="76c97-695">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-695">Requirement</span></span>|<span data-ttu-id="76c97-696">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-697">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-698">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-698">1.0</span></span>|
|[<span data-ttu-id="76c97-699">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-699">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-700">ReadItem</span></span>|
|[<span data-ttu-id="76c97-701">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-701">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-702">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-702">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="76c97-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="76c97-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="76c97-704">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="76c97-704">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="76c97-705">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="76c97-705">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="76c97-706">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-706">Read mode</span></span>

<span data-ttu-id="76c97-p135">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="76c97-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="76c97-709">Mode composition</span><span class="sxs-lookup"><span data-stu-id="76c97-709">Compose mode</span></span>

<span data-ttu-id="76c97-710">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="76c97-710">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="76c97-711">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-711">Type</span></span>

*   <span data-ttu-id="76c97-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="76c97-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-713">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-713">Requirements</span></span>

|<span data-ttu-id="76c97-714">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-714">Requirement</span></span>|<span data-ttu-id="76c97-715">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-716">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-717">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-717">1.0</span></span>|
|[<span data-ttu-id="76c97-718">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-718">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-719">ReadItem</span></span>|
|[<span data-ttu-id="76c97-720">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-720">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-721">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-721">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="76c97-722">Méthodes</span><span class="sxs-lookup"><span data-stu-id="76c97-722">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="76c97-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="76c97-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="76c97-724">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="76c97-724">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="76c97-725">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="76c97-725">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="76c97-726">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="76c97-726">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-727">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-727">Parameters</span></span>
|<span data-ttu-id="76c97-728">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-728">Name</span></span>|<span data-ttu-id="76c97-729">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-729">Type</span></span>|<span data-ttu-id="76c97-730">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-730">Attributes</span></span>|<span data-ttu-id="76c97-731">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-731">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="76c97-732">String</span><span class="sxs-lookup"><span data-stu-id="76c97-732">String</span></span>||<span data-ttu-id="76c97-p136">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="76c97-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="76c97-735">String</span><span class="sxs-lookup"><span data-stu-id="76c97-735">String</span></span>||<span data-ttu-id="76c97-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="76c97-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="76c97-738">Object</span><span class="sxs-lookup"><span data-stu-id="76c97-738">Object</span></span>|<span data-ttu-id="76c97-739">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-739">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-740">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-741">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-741">Object</span></span>|<span data-ttu-id="76c97-742">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-742">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-743">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="76c97-744">Boolean</span><span class="sxs-lookup"><span data-stu-id="76c97-744">Boolean</span></span>|<span data-ttu-id="76c97-745">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-745">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-746">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="76c97-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="76c97-747">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-747">function</span></span>|<span data-ttu-id="76c97-748">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-748">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-749">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="76c97-750">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="76c97-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="76c97-751">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="76c97-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="76c97-752">Erreurs</span><span class="sxs-lookup"><span data-stu-id="76c97-752">Errors</span></span>

|<span data-ttu-id="76c97-753">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="76c97-753">Error code</span></span>|<span data-ttu-id="76c97-754">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="76c97-755">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="76c97-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="76c97-756">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="76c97-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="76c97-757">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="76c97-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-758">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-758">Requirements</span></span>

|<span data-ttu-id="76c97-759">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-759">Requirement</span></span>|<span data-ttu-id="76c97-760">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-761">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-762">1.1</span><span class="sxs-lookup"><span data-stu-id="76c97-762">1.1</span></span>|
|[<span data-ttu-id="76c97-763">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="76c97-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="76c97-765">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-766">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="76c97-767">Exemples</span><span class="sxs-lookup"><span data-stu-id="76c97-767">Examples</span></span>

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

<span data-ttu-id="76c97-768">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="76c97-768">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="76c97-769">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="76c97-769">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="76c97-770">Ajoute un fichier à partir du codage Base64 à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="76c97-770">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="76c97-771">La `addFileAttachmentFromBase64Async` méthode charge le fichier à partir du codage Base64 et l'associe à l'élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="76c97-771">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="76c97-772">Cette méthode renvoie l'identificateur de pièce jointe dans l'objet AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="76c97-772">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="76c97-773">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="76c97-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-774">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-774">Parameters</span></span>
|<span data-ttu-id="76c97-775">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-775">Name</span></span>|<span data-ttu-id="76c97-776">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-776">Type</span></span>|<span data-ttu-id="76c97-777">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-777">Attributes</span></span>|<span data-ttu-id="76c97-778">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-778">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="76c97-779">String</span><span class="sxs-lookup"><span data-stu-id="76c97-779">String</span></span>||<span data-ttu-id="76c97-780">Contenu encodé en base64 d'une image ou d'un fichier à ajouter à un message électronique ou à un événement.</span><span class="sxs-lookup"><span data-stu-id="76c97-780">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="76c97-781">String</span><span class="sxs-lookup"><span data-stu-id="76c97-781">String</span></span>||<span data-ttu-id="76c97-p139">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="76c97-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="76c97-784">Object</span><span class="sxs-lookup"><span data-stu-id="76c97-784">Object</span></span>|<span data-ttu-id="76c97-785">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-785">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-786">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-786">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-787">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-787">Object</span></span>|<span data-ttu-id="76c97-788">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-788">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-789">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-789">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="76c97-790">Boolean</span><span class="sxs-lookup"><span data-stu-id="76c97-790">Boolean</span></span>|<span data-ttu-id="76c97-791">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-791">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-792">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="76c97-792">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="76c97-793">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-793">function</span></span>|<span data-ttu-id="76c97-794">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-794">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-795">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-795">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="76c97-796">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="76c97-796">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="76c97-797">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="76c97-797">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="76c97-798">Erreurs</span><span class="sxs-lookup"><span data-stu-id="76c97-798">Errors</span></span>

|<span data-ttu-id="76c97-799">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="76c97-799">Error code</span></span>|<span data-ttu-id="76c97-800">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-800">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="76c97-801">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="76c97-801">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="76c97-802">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="76c97-802">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="76c97-803">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="76c97-803">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-804">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-804">Requirements</span></span>

|<span data-ttu-id="76c97-805">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-805">Requirement</span></span>|<span data-ttu-id="76c97-806">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-807">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-808">Aperçu</span><span class="sxs-lookup"><span data-stu-id="76c97-808">Preview</span></span>|
|[<span data-ttu-id="76c97-809">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-809">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-810">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="76c97-810">ReadWriteItem</span></span>|
|[<span data-ttu-id="76c97-811">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-811">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-812">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-812">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="76c97-813">範例</span><span class="sxs-lookup"><span data-stu-id="76c97-813">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="76c97-814">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="76c97-814">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="76c97-815">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="76c97-815">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="76c97-816">Actuellement, les types d'événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="76c97-816">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-817">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-817">Parameters</span></span>

| <span data-ttu-id="76c97-818">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-818">Name</span></span> | <span data-ttu-id="76c97-819">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-819">Type</span></span> | <span data-ttu-id="76c97-820">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-820">Attributes</span></span> | <span data-ttu-id="76c97-821">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-821">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="76c97-822">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="76c97-822">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="76c97-823">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="76c97-823">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="76c97-824">Fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-824">Function</span></span> || <span data-ttu-id="76c97-p140">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="76c97-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="76c97-828">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-828">Object</span></span> | <span data-ttu-id="76c97-829">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-829">&lt;optional&gt;</span></span> | <span data-ttu-id="76c97-830">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-830">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="76c97-831">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-831">Object</span></span> | <span data-ttu-id="76c97-832">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-832">&lt;optional&gt;</span></span> | <span data-ttu-id="76c97-833">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-833">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="76c97-834">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-834">function</span></span>| <span data-ttu-id="76c97-835">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-835">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-836">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-836">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-837">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-837">Requirements</span></span>

|<span data-ttu-id="76c97-838">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-838">Requirement</span></span>| <span data-ttu-id="76c97-839">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-840">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76c97-841">1.7</span><span class="sxs-lookup"><span data-stu-id="76c97-841">1.7</span></span> |
|[<span data-ttu-id="76c97-842">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76c97-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-843">ReadItem</span></span> |
|[<span data-ttu-id="76c97-844">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76c97-845">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-845">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="76c97-846">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-846">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="76c97-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="76c97-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="76c97-848">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-848">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="76c97-p141">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="76c97-852">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="76c97-852">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="76c97-853">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="76c97-853">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-854">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-854">Parameters</span></span>

|<span data-ttu-id="76c97-855">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-855">Name</span></span>|<span data-ttu-id="76c97-856">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-856">Type</span></span>|<span data-ttu-id="76c97-857">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-857">Attributes</span></span>|<span data-ttu-id="76c97-858">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-858">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="76c97-859">String</span><span class="sxs-lookup"><span data-stu-id="76c97-859">String</span></span>||<span data-ttu-id="76c97-p142">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="76c97-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="76c97-862">String</span><span class="sxs-lookup"><span data-stu-id="76c97-862">String</span></span>||<span data-ttu-id="76c97-863">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="76c97-863">The subject of the item to be attached.</span></span> <span data-ttu-id="76c97-864">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="76c97-864">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="76c97-865">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-865">Object</span></span>|<span data-ttu-id="76c97-866">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-866">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-867">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-867">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-868">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-868">Object</span></span>|<span data-ttu-id="76c97-869">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-869">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-870">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-870">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="76c97-871">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-871">function</span></span>|<span data-ttu-id="76c97-872">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-872">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-873">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-873">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="76c97-874">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="76c97-874">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="76c97-875">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="76c97-875">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="76c97-876">Erreurs</span><span class="sxs-lookup"><span data-stu-id="76c97-876">Errors</span></span>

|<span data-ttu-id="76c97-877">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="76c97-877">Error code</span></span>|<span data-ttu-id="76c97-878">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-878">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="76c97-879">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="76c97-879">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-880">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-880">Requirements</span></span>

|<span data-ttu-id="76c97-881">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-881">Requirement</span></span>|<span data-ttu-id="76c97-882">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-883">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-884">1.1</span><span class="sxs-lookup"><span data-stu-id="76c97-884">1.1</span></span>|
|[<span data-ttu-id="76c97-885">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-885">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-886">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="76c97-886">ReadWriteItem</span></span>|
|[<span data-ttu-id="76c97-887">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-887">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-888">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-888">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-889">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-889">Example</span></span>

<span data-ttu-id="76c97-890">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="76c97-890">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="76c97-891">close()</span><span class="sxs-lookup"><span data-stu-id="76c97-891">close()</span></span>

<span data-ttu-id="76c97-892">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="76c97-892">Closes the current item that is being composed.</span></span>

<span data-ttu-id="76c97-p144">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="76c97-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-895">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-895">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="76c97-896">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="76c97-896">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-897">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-897">Requirements</span></span>

|<span data-ttu-id="76c97-898">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-898">Requirement</span></span>|<span data-ttu-id="76c97-899">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-899">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-900">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-900">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-901">1.3</span><span class="sxs-lookup"><span data-stu-id="76c97-901">1.3</span></span>|
|[<span data-ttu-id="76c97-902">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-902">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-903">Restreinte</span><span class="sxs-lookup"><span data-stu-id="76c97-903">Restricted</span></span>|
|[<span data-ttu-id="76c97-904">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-904">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-905">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-905">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="76c97-906">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="76c97-906">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="76c97-907">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="76c97-907">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-908">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-908">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76c97-909">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="76c97-909">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="76c97-910">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="76c97-910">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="76c97-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="76c97-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-914">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-914">Parameters</span></span>

|<span data-ttu-id="76c97-915">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-915">Name</span></span>|<span data-ttu-id="76c97-916">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-916">Type</span></span>|<span data-ttu-id="76c97-917">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-917">Attributes</span></span>|<span data-ttu-id="76c97-918">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-918">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="76c97-919">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="76c97-919">String &#124; Object</span></span>||<span data-ttu-id="76c97-920">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="76c97-920">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="76c97-921">La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="76c97-921">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="76c97-922">**OU**</span><span class="sxs-lookup"><span data-stu-id="76c97-922">**OR**</span></span><br/><span data-ttu-id="76c97-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="76c97-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="76c97-925">String</span><span class="sxs-lookup"><span data-stu-id="76c97-925">String</span></span>|<span data-ttu-id="76c97-926">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-926">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="76c97-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="76c97-929">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-929">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="76c97-930">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-930">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-931">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-931">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="76c97-932">Chaîne</span><span class="sxs-lookup"><span data-stu-id="76c97-932">String</span></span>||<span data-ttu-id="76c97-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="76c97-935">String</span><span class="sxs-lookup"><span data-stu-id="76c97-935">String</span></span>||<span data-ttu-id="76c97-936">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="76c97-936">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="76c97-937">Chaîne</span><span class="sxs-lookup"><span data-stu-id="76c97-937">String</span></span>||<span data-ttu-id="76c97-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="76c97-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="76c97-940">Booléen</span><span class="sxs-lookup"><span data-stu-id="76c97-940">Boolean</span></span>||<span data-ttu-id="76c97-p151">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="76c97-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="76c97-943">String</span><span class="sxs-lookup"><span data-stu-id="76c97-943">String</span></span>||<span data-ttu-id="76c97-p152">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="76c97-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="76c97-947">function</span><span class="sxs-lookup"><span data-stu-id="76c97-947">function</span></span>|<span data-ttu-id="76c97-948">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-948">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-949">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-949">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-950">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-950">Requirements</span></span>

|<span data-ttu-id="76c97-951">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-951">Requirement</span></span>|<span data-ttu-id="76c97-952">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-952">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-953">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-953">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-954">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-954">1.0</span></span>|
|[<span data-ttu-id="76c97-955">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-955">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-956">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-956">ReadItem</span></span>|
|[<span data-ttu-id="76c97-957">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-957">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-958">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-958">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="76c97-959">Exemples</span><span class="sxs-lookup"><span data-stu-id="76c97-959">Examples</span></span>

<span data-ttu-id="76c97-960">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="76c97-960">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="76c97-961">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="76c97-961">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="76c97-962">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="76c97-962">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="76c97-963">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="76c97-963">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="76c97-964">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-964">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="76c97-965">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-965">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="76c97-966">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="76c97-966">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="76c97-967">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="76c97-967">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-968">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-968">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76c97-969">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="76c97-969">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="76c97-970">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="76c97-970">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="76c97-p153">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="76c97-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-974">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-974">Parameters</span></span>

|<span data-ttu-id="76c97-975">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-975">Name</span></span>|<span data-ttu-id="76c97-976">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-976">Type</span></span>|<span data-ttu-id="76c97-977">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-977">Attributes</span></span>|<span data-ttu-id="76c97-978">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-978">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="76c97-979">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="76c97-979">String &#124; Object</span></span>||<span data-ttu-id="76c97-980">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="76c97-980">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="76c97-981">La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="76c97-981">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="76c97-982">**OU**</span><span class="sxs-lookup"><span data-stu-id="76c97-982">**OR**</span></span><br/><span data-ttu-id="76c97-p155">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="76c97-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="76c97-985">String</span><span class="sxs-lookup"><span data-stu-id="76c97-985">String</span></span>|<span data-ttu-id="76c97-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-986">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-p156">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="76c97-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="76c97-989">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-989">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="76c97-990">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-990">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-991">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-991">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="76c97-992">Chaîne</span><span class="sxs-lookup"><span data-stu-id="76c97-992">String</span></span>||<span data-ttu-id="76c97-p157">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="76c97-995">String</span><span class="sxs-lookup"><span data-stu-id="76c97-995">String</span></span>||<span data-ttu-id="76c97-996">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="76c97-996">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="76c97-997">Chaîne</span><span class="sxs-lookup"><span data-stu-id="76c97-997">String</span></span>||<span data-ttu-id="76c97-p158">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="76c97-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="76c97-1000">Booléen</span><span class="sxs-lookup"><span data-stu-id="76c97-1000">Boolean</span></span>||<span data-ttu-id="76c97-p159">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="76c97-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="76c97-1003">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1003">String</span></span>||<span data-ttu-id="76c97-p160">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="76c97-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="76c97-1007">function</span><span class="sxs-lookup"><span data-stu-id="76c97-1007">function</span></span>|<span data-ttu-id="76c97-1008">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1009">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1010">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1010">Requirements</span></span>

|<span data-ttu-id="76c97-1011">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1011">Requirement</span></span>|<span data-ttu-id="76c97-1012">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1013">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1014">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-1014">1.0</span></span>|
|[<span data-ttu-id="76c97-1015">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1015">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1016">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1017">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1017">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1018">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1018">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="76c97-1019">Exemples</span><span class="sxs-lookup"><span data-stu-id="76c97-1019">Examples</span></span>

<span data-ttu-id="76c97-1020">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="76c97-1020">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="76c97-1021">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="76c97-1021">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="76c97-1022">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="76c97-1022">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="76c97-1023">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="76c97-1023">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="76c97-1024">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-1024">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="76c97-1025">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1025">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="76c97-1026">getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="76c97-1026">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="76c97-1027">Obtient la pièce jointe spécifiée à partir d'un message ou d'un `AttachmentContent` rendez-vous et la renvoie en tant qu'objet.</span><span class="sxs-lookup"><span data-stu-id="76c97-1027">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="76c97-1028">La `getAttachmentContentAsync` méthode obtient la pièce jointe avec l'identificateur spécifié à partir de l'élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-1028">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="76c97-1029">Il est recommandé d'utiliser l'identificateur pour récupérer une pièce jointe dans la même session que l'attachmentIds a été récupérée avec l' `getAttachmentsAsync` appel ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="76c97-1029">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="76c97-1030">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="76c97-1030">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="76c97-1031">Une session est terminée lorsque l'utilisateur ferme l'application, ou si l'utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="76c97-1031">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1032">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1032">Parameters</span></span>

|<span data-ttu-id="76c97-1033">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1033">Name</span></span>|<span data-ttu-id="76c97-1034">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1034">Type</span></span>|<span data-ttu-id="76c97-1035">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1035">Attributes</span></span>|<span data-ttu-id="76c97-1036">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1036">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="76c97-1037">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1037">String</span></span>||<span data-ttu-id="76c97-1038">Identificateur de la pièce jointe que vous souhaitez obtenir.</span><span class="sxs-lookup"><span data-stu-id="76c97-1038">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="76c97-1039">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1039">Object</span></span>|<span data-ttu-id="76c97-1040">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1041">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-1041">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-1042">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1042">Object</span></span>|<span data-ttu-id="76c97-1043">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1044">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1044">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="76c97-1045">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-1045">function</span></span>|<span data-ttu-id="76c97-1046">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1046">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1047">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1047">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1048">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1048">Requirements</span></span>

|<span data-ttu-id="76c97-1049">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1049">Requirement</span></span>|<span data-ttu-id="76c97-1050">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1051">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1052">Aperçu</span><span class="sxs-lookup"><span data-stu-id="76c97-1052">Preview</span></span>|
|[<span data-ttu-id="76c97-1053">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1054">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1055">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1056">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1056">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1057">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1057">Returns:</span></span>

<span data-ttu-id="76c97-1058">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="76c97-1058">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="76c97-1059">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1059">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="76c97-1060">getAttachmentsAsync ([options], [Rappel]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="76c97-1060">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="76c97-1061">Obtient les pièces jointes de l'élément sous la forme d'un tableau.</span><span class="sxs-lookup"><span data-stu-id="76c97-1061">Gets the item's attachments as an array.</span></span> <span data-ttu-id="76c97-1062">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="76c97-1062">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1063">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1063">Parameters</span></span>

|<span data-ttu-id="76c97-1064">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1064">Name</span></span>|<span data-ttu-id="76c97-1065">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1065">Type</span></span>|<span data-ttu-id="76c97-1066">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1066">Attributes</span></span>|<span data-ttu-id="76c97-1067">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="76c97-1068">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1068">Object</span></span>|<span data-ttu-id="76c97-1069">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1070">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-1071">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1071">Object</span></span>|<span data-ttu-id="76c97-1072">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1073">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="76c97-1074">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-1074">function</span></span>|<span data-ttu-id="76c97-1075">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1076">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1077">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1077">Requirements</span></span>

|<span data-ttu-id="76c97-1078">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1078">Requirement</span></span>|<span data-ttu-id="76c97-1079">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1080">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1081">Aperçu</span><span class="sxs-lookup"><span data-stu-id="76c97-1081">Preview</span></span>|
|[<span data-ttu-id="76c97-1082">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1082">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1083">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1083">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1084">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1084">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1085">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-1085">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1086">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1086">Returns:</span></span>

<span data-ttu-id="76c97-1087">Type: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="76c97-1087">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="76c97-1088">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1088">Example</span></span>

<span data-ttu-id="76c97-1089">L'exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l'élément actif.</span><span class="sxs-lookup"><span data-stu-id="76c97-1089">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="76c97-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="76c97-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="76c97-1091">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="76c97-1091">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1092">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-1093">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1093">Requirements</span></span>

|<span data-ttu-id="76c97-1094">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1094">Requirement</span></span>|<span data-ttu-id="76c97-1095">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1096">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1097">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-1097">1.0</span></span>|
|[<span data-ttu-id="76c97-1098">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1098">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1099">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1100">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1100">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1101">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1102">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1102">Returns:</span></span>

<span data-ttu-id="76c97-1103">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="76c97-1103">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="76c97-1104">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1104">Example</span></span>

<span data-ttu-id="76c97-1105">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="76c97-1105">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="76c97-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="76c97-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="76c97-1107">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="76c97-1107">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1108">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-1108">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1109">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1109">Parameters</span></span>

|<span data-ttu-id="76c97-1110">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1110">Name</span></span>|<span data-ttu-id="76c97-1111">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1111">Type</span></span>|<span data-ttu-id="76c97-1112">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1112">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="76c97-1113">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="76c97-1113">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="76c97-1114">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="76c97-1114">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1115">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1115">Requirements</span></span>

|<span data-ttu-id="76c97-1116">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1116">Requirement</span></span>|<span data-ttu-id="76c97-1117">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1117">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1118">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1118">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1119">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-1119">1.0</span></span>|
|[<span data-ttu-id="76c97-1120">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1120">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1121">Restreinte</span><span class="sxs-lookup"><span data-stu-id="76c97-1121">Restricted</span></span>|
|[<span data-ttu-id="76c97-1122">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1122">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1123">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1123">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1124">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1124">Returns:</span></span>

<span data-ttu-id="76c97-1125">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="76c97-1125">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="76c97-1126">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="76c97-1126">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="76c97-1127">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="76c97-1127">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="76c97-1128">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="76c97-1128">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="76c97-1129">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="76c97-1129">Value of `entityType`</span></span>|<span data-ttu-id="76c97-1130">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="76c97-1130">Type of objects in returned array</span></span>|<span data-ttu-id="76c97-1131">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="76c97-1131">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="76c97-1132">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1132">String</span></span>|<span data-ttu-id="76c97-1133">**Restreinte**</span><span class="sxs-lookup"><span data-stu-id="76c97-1133">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="76c97-1134">Contact</span><span class="sxs-lookup"><span data-stu-id="76c97-1134">Contact</span></span>|<span data-ttu-id="76c97-1135">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="76c97-1135">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="76c97-1136">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1136">String</span></span>|<span data-ttu-id="76c97-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="76c97-1137">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="76c97-1138">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="76c97-1138">MeetingSuggestion</span></span>|<span data-ttu-id="76c97-1139">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="76c97-1139">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="76c97-1140">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="76c97-1140">PhoneNumber</span></span>|<span data-ttu-id="76c97-1141">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="76c97-1141">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="76c97-1142">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="76c97-1142">TaskSuggestion</span></span>|<span data-ttu-id="76c97-1143">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="76c97-1143">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="76c97-1144">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1144">String</span></span>|<span data-ttu-id="76c97-1145">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="76c97-1145">**Restricted**</span></span>|

<span data-ttu-id="76c97-1146">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="76c97-1146">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="76c97-1147">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1147">Example</span></span>

<span data-ttu-id="76c97-1148">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="76c97-1148">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="76c97-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="76c97-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="76c97-1150">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="76c97-1150">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1151">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-1151">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76c97-1152">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="76c97-1152">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1153">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1153">Parameters</span></span>

|<span data-ttu-id="76c97-1154">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1154">Name</span></span>|<span data-ttu-id="76c97-1155">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1155">Type</span></span>|<span data-ttu-id="76c97-1156">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1156">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="76c97-1157">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1157">String</span></span>|<span data-ttu-id="76c97-1158">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="76c97-1158">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1159">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1159">Requirements</span></span>

|<span data-ttu-id="76c97-1160">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1160">Requirement</span></span>|<span data-ttu-id="76c97-1161">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1162">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1163">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-1163">1.0</span></span>|
|[<span data-ttu-id="76c97-1164">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1164">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1165">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1166">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1166">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1167">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1167">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1168">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1168">Returns:</span></span>

<span data-ttu-id="76c97-p164">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="76c97-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="76c97-1171">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="76c97-1171">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="76c97-1172">getInitializationContextAsync ([options], [Rappel])</span><span class="sxs-lookup"><span data-stu-id="76c97-1172">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="76c97-1173">Obtient les données d'initialisation transmises lorsque le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="76c97-1173">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1174">Cette méthode est uniquement prise en charge par Outlook 2016 ou une version ultérieure pour Windows (versions «démarrer en un clic» ultérieures à 16.0.8413.1000) et Outlook sur le Web pour Office 365.</span><span class="sxs-lookup"><span data-stu-id="76c97-1174">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1175">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1175">Parameters</span></span>
|<span data-ttu-id="76c97-1176">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1176">Name</span></span>|<span data-ttu-id="76c97-1177">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1177">Type</span></span>|<span data-ttu-id="76c97-1178">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1178">Attributes</span></span>|<span data-ttu-id="76c97-1179">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1179">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="76c97-1180">Object</span><span class="sxs-lookup"><span data-stu-id="76c97-1180">Object</span></span>|<span data-ttu-id="76c97-1181">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1182">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-1182">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-1183">Object</span><span class="sxs-lookup"><span data-stu-id="76c97-1183">Object</span></span>|<span data-ttu-id="76c97-1184">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1184">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1185">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1185">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="76c97-1186">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-1186">function</span></span>|<span data-ttu-id="76c97-1187">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1188">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1188">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="76c97-1189">En cas de réussite, les données d'initialisation sont fournies `asyncResult.value` dans la propriété sous la forme d'une chaîne.</span><span class="sxs-lookup"><span data-stu-id="76c97-1189">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="76c97-1190">S'il n'existe pas de contexte d'initialisation `asyncResult` , l'objet contient `Error` un objet dont `code` la propriété est `9020` définie sur `name` et sa propriété `GenericResponseError`est définie sur.</span><span class="sxs-lookup"><span data-stu-id="76c97-1190">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1191">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1191">Requirements</span></span>

|<span data-ttu-id="76c97-1192">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1192">Requirement</span></span>|<span data-ttu-id="76c97-1193">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1194">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1195">Aperçu</span><span class="sxs-lookup"><span data-stu-id="76c97-1195">Preview</span></span>|
|[<span data-ttu-id="76c97-1196">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1197">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1199">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1199">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-1200">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1200">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="76c97-1201">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="76c97-1201">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="76c97-1202">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="76c97-1202">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1203">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76c97-p165">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="76c97-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="76c97-1207">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="76c97-1207">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="76c97-1208">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="76c97-1208">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="76c97-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-1212">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1212">Requirements</span></span>

|<span data-ttu-id="76c97-1213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1213">Requirement</span></span>|<span data-ttu-id="76c97-1214">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1215">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-1216">1.0</span></span>|
|[<span data-ttu-id="76c97-1217">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1217">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1218">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1219">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1219">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1220">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1220">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1221">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1221">Returns:</span></span>

<span data-ttu-id="76c97-p167">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="76c97-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="76c97-1224">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="76c97-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="76c97-1225">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1225">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="76c97-1226">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1226">Example</span></span>

<span data-ttu-id="76c97-1227">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="76c97-1227">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="76c97-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="76c97-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="76c97-1229">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="76c97-1229">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1230">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-1230">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76c97-1231">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="76c97-1231">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="76c97-p168">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="76c97-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1234">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1234">Parameters</span></span>

|<span data-ttu-id="76c97-1235">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1235">Name</span></span>|<span data-ttu-id="76c97-1236">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1236">Type</span></span>|<span data-ttu-id="76c97-1237">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1237">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="76c97-1238">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1238">String</span></span>|<span data-ttu-id="76c97-1239">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="76c97-1239">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1240">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1240">Requirements</span></span>

|<span data-ttu-id="76c97-1241">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1241">Requirement</span></span>|<span data-ttu-id="76c97-1242">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1243">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1243">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-1244">1.0</span></span>|
|[<span data-ttu-id="76c97-1245">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1246">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1247">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1248">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1248">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1249">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1249">Returns:</span></span>

<span data-ttu-id="76c97-1250">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="76c97-1250">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="76c97-1251">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="76c97-1251">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="76c97-1252">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="76c97-1252">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="76c97-1253">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1253">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="76c97-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="76c97-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="76c97-1255">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="76c97-1255">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="76c97-p169">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="76c97-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1258">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1258">Parameters</span></span>

|<span data-ttu-id="76c97-1259">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1259">Name</span></span>|<span data-ttu-id="76c97-1260">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1260">Type</span></span>|<span data-ttu-id="76c97-1261">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1261">Attributes</span></span>|<span data-ttu-id="76c97-1262">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1262">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="76c97-1263">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="76c97-1263">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="76c97-p170">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="76c97-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="76c97-1267">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1267">Object</span></span>|<span data-ttu-id="76c97-1268">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1269">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-1270">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1270">Object</span></span>|<span data-ttu-id="76c97-1271">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1272">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="76c97-1273">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-1273">function</span></span>||<span data-ttu-id="76c97-1274">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="76c97-1275">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="76c97-1275">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="76c97-1276">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="76c97-1276">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1277">Requirements</span></span>

|<span data-ttu-id="76c97-1278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1278">Requirement</span></span>|<span data-ttu-id="76c97-1279">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1279">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1281">1.2</span><span class="sxs-lookup"><span data-stu-id="76c97-1281">1.2</span></span>|
|[<span data-ttu-id="76c97-1282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1283">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1283">ReadWriteItem</span></span>|
|[<span data-ttu-id="76c97-1284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1285">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-1285">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1286">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1286">Returns:</span></span>

<span data-ttu-id="76c97-1287">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="76c97-1287">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="76c97-1288">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="76c97-1288">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="76c97-1289">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1289">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="76c97-1290">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1290">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="76c97-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="76c97-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="76c97-1292">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="76c97-1292">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="76c97-1293">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="76c97-1293">Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1294">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-1294">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-1295">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1295">Requirements</span></span>

|<span data-ttu-id="76c97-1296">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1296">Requirement</span></span>|<span data-ttu-id="76c97-1297">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1298">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1299">1.6</span><span class="sxs-lookup"><span data-stu-id="76c97-1299">1.6</span></span>|
|[<span data-ttu-id="76c97-1300">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1300">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1301">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1302">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1302">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1303">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1303">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1304">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1304">Returns:</span></span>

<span data-ttu-id="76c97-1305">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="76c97-1305">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="76c97-1306">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1306">Example</span></span>

<span data-ttu-id="76c97-1307">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="76c97-1307">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="76c97-1308">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="76c97-1308">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="76c97-p173">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="76c97-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1311">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="76c97-1311">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76c97-p174">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="76c97-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="76c97-1315">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="76c97-1315">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="76c97-1316">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="76c97-1316">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="76c97-p175">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="76c97-1320">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1320">Requirements</span></span>

|<span data-ttu-id="76c97-1321">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1321">Requirement</span></span>|<span data-ttu-id="76c97-1322">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1322">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1323">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1324">1.6</span><span class="sxs-lookup"><span data-stu-id="76c97-1324">1.6</span></span>|
|[<span data-ttu-id="76c97-1325">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1326">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1327">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1328">Lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1328">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76c97-1329">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="76c97-1329">Returns:</span></span>

<span data-ttu-id="76c97-p176">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="76c97-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="76c97-1332">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1332">Example</span></span>

<span data-ttu-id="76c97-1333">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="76c97-1333">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="76c97-1334">getSharedPropertiesAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="76c97-1334">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="76c97-1335">Obtient les propriétés du rendez-vous ou du message sélectionné dans un dossier partagé, un calendrier ou une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="76c97-1335">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1336">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1336">Parameters</span></span>

|<span data-ttu-id="76c97-1337">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1337">Name</span></span>|<span data-ttu-id="76c97-1338">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1338">Type</span></span>|<span data-ttu-id="76c97-1339">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1339">Attributes</span></span>|<span data-ttu-id="76c97-1340">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1340">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="76c97-1341">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1341">Object</span></span>|<span data-ttu-id="76c97-1342">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1342">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1343">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-1343">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-1344">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1344">Object</span></span>|<span data-ttu-id="76c97-1345">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1346">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1346">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="76c97-1347">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-1347">function</span></span>||<span data-ttu-id="76c97-1348">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1348">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="76c97-1349">Les propriétés partagées sont fournies sous [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) la forme d' `asyncResult.value` un objet dans la propriété.</span><span class="sxs-lookup"><span data-stu-id="76c97-1349">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="76c97-1350">Cet objet peut être utilisé pour obtenir les propriétés partagées de l'élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-1350">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1351">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1351">Requirements</span></span>

|<span data-ttu-id="76c97-1352">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1352">Requirement</span></span>|<span data-ttu-id="76c97-1353">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1354">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1355">Aperçu</span><span class="sxs-lookup"><span data-stu-id="76c97-1355">Preview</span></span>|
|[<span data-ttu-id="76c97-1356">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1356">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1357">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1358">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1358">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1359">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1359">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-1360">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1360">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="76c97-1361">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="76c97-1361">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="76c97-1362">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="76c97-1362">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="76c97-p178">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="76c97-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1366">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1366">Parameters</span></span>

|<span data-ttu-id="76c97-1367">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1367">Name</span></span>|<span data-ttu-id="76c97-1368">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1368">Type</span></span>|<span data-ttu-id="76c97-1369">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1369">Attributes</span></span>|<span data-ttu-id="76c97-1370">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1370">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="76c97-1371">function</span><span class="sxs-lookup"><span data-stu-id="76c97-1371">function</span></span>||<span data-ttu-id="76c97-1372">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1372">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="76c97-1373">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="76c97-1373">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="76c97-1374">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="76c97-1374">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="76c97-1375">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1375">Object</span></span>|<span data-ttu-id="76c97-1376">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1376">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1377">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1377">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="76c97-1378">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1378">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1379">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1379">Requirements</span></span>

|<span data-ttu-id="76c97-1380">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1380">Requirement</span></span>|<span data-ttu-id="76c97-1381">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1381">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1382">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1383">1.0</span><span class="sxs-lookup"><span data-stu-id="76c97-1383">1.0</span></span>|
|[<span data-ttu-id="76c97-1384">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1385">ReadItem</span></span>|
|[<span data-ttu-id="76c97-1386">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1387">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1387">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-1388">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1388">Example</span></span>

<span data-ttu-id="76c97-p181">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="76c97-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="76c97-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="76c97-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="76c97-1393">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="76c97-1393">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="76c97-1394">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-1394">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="76c97-1395">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="76c97-1395">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="76c97-1396">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="76c97-1396">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="76c97-1397">Une session est terminée lorsque l'utilisateur ferme l'application, ou si l'utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="76c97-1397">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1398">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1398">Parameters</span></span>

|<span data-ttu-id="76c97-1399">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1399">Name</span></span>|<span data-ttu-id="76c97-1400">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1400">Type</span></span>|<span data-ttu-id="76c97-1401">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1401">Attributes</span></span>|<span data-ttu-id="76c97-1402">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1402">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="76c97-1403">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1403">String</span></span>||<span data-ttu-id="76c97-1404">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="76c97-1404">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="76c97-1405">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1405">Object</span></span>|<span data-ttu-id="76c97-1406">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1407">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-1407">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-1408">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1408">Object</span></span>|<span data-ttu-id="76c97-1409">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1409">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1410">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1410">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="76c97-1411">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-1411">function</span></span>|<span data-ttu-id="76c97-1412">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1412">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1413">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1413">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="76c97-1414">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="76c97-1414">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="76c97-1415">Erreurs</span><span class="sxs-lookup"><span data-stu-id="76c97-1415">Errors</span></span>

|<span data-ttu-id="76c97-1416">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="76c97-1416">Error code</span></span>|<span data-ttu-id="76c97-1417">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1417">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="76c97-1418">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="76c97-1418">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1419">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1419">Requirements</span></span>

|<span data-ttu-id="76c97-1420">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1420">Requirement</span></span>|<span data-ttu-id="76c97-1421">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1421">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1422">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1423">1.1</span><span class="sxs-lookup"><span data-stu-id="76c97-1423">1.1</span></span>|
|[<span data-ttu-id="76c97-1424">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1424">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1425">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1425">ReadWriteItem</span></span>|
|[<span data-ttu-id="76c97-1426">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1426">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1427">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-1427">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-1428">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1428">Example</span></span>

<span data-ttu-id="76c97-1429">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="76c97-1429">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="76c97-1430">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="76c97-1430">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="76c97-1431">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="76c97-1431">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="76c97-1432">Actuellement, les types d'événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="76c97-1432">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1433">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1433">Parameters</span></span>

| <span data-ttu-id="76c97-1434">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1434">Name</span></span> | <span data-ttu-id="76c97-1435">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1435">Type</span></span> | <span data-ttu-id="76c97-1436">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1436">Attributes</span></span> | <span data-ttu-id="76c97-1437">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1437">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="76c97-1438">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="76c97-1438">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="76c97-1439">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="76c97-1439">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="76c97-1440">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1440">Object</span></span> | <span data-ttu-id="76c97-1441">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1441">&lt;optional&gt;</span></span> | <span data-ttu-id="76c97-1442">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-1442">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="76c97-1443">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1443">Object</span></span> | <span data-ttu-id="76c97-1444">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1444">&lt;optional&gt;</span></span> | <span data-ttu-id="76c97-1445">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1445">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="76c97-1446">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-1446">function</span></span>| <span data-ttu-id="76c97-1447">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1447">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1448">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1448">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1449">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1449">Requirements</span></span>

|<span data-ttu-id="76c97-1450">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1450">Requirement</span></span>| <span data-ttu-id="76c97-1451">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1451">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1452">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76c97-1453">1.7</span><span class="sxs-lookup"><span data-stu-id="76c97-1453">1.7</span></span> |
|[<span data-ttu-id="76c97-1454">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1454">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76c97-1455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1455">ReadItem</span></span> |
|[<span data-ttu-id="76c97-1456">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1456">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76c97-1457">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="76c97-1457">Compose or Read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="76c97-1458">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="76c97-1458">saveAsync([options], callback)</span></span>

<span data-ttu-id="76c97-1459">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="76c97-1459">Asynchronously saves an item.</span></span>

<span data-ttu-id="76c97-p183">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="76c97-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1463">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="76c97-1463">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="76c97-1464">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="76c97-1464">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="76c97-p185">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="76c97-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="76c97-1468">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="76c97-1468">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="76c97-1469">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="76c97-1469">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="76c97-1470">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="76c97-1470">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="76c97-1471">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="76c97-1471">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1472">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1472">Parameters</span></span>

|<span data-ttu-id="76c97-1473">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1473">Name</span></span>|<span data-ttu-id="76c97-1474">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1474">Type</span></span>|<span data-ttu-id="76c97-1475">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1475">Attributes</span></span>|<span data-ttu-id="76c97-1476">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1476">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="76c97-1477">Object</span><span class="sxs-lookup"><span data-stu-id="76c97-1477">Object</span></span>|<span data-ttu-id="76c97-1478">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1478">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1479">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-1479">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-1480">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1480">Object</span></span>|<span data-ttu-id="76c97-1481">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1481">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1482">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1482">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="76c97-1483">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-1483">function</span></span>||<span data-ttu-id="76c97-1484">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1484">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="76c97-1485">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="76c97-1485">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1486">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1486">Requirements</span></span>

|<span data-ttu-id="76c97-1487">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1487">Requirement</span></span>|<span data-ttu-id="76c97-1488">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1489">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1490">1.3</span><span class="sxs-lookup"><span data-stu-id="76c97-1490">1.3</span></span>|
|[<span data-ttu-id="76c97-1491">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1492">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1492">ReadWriteItem</span></span>|
|[<span data-ttu-id="76c97-1493">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1494">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-1494">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="76c97-1495">範例</span><span class="sxs-lookup"><span data-stu-id="76c97-1495">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="76c97-p187">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="76c97-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="76c97-1498">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="76c97-1498">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="76c97-1499">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="76c97-1499">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="76c97-p188">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="76c97-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76c97-1503">Paramètres</span><span class="sxs-lookup"><span data-stu-id="76c97-1503">Parameters</span></span>

|<span data-ttu-id="76c97-1504">Nom</span><span class="sxs-lookup"><span data-stu-id="76c97-1504">Name</span></span>|<span data-ttu-id="76c97-1505">Type</span><span class="sxs-lookup"><span data-stu-id="76c97-1505">Type</span></span>|<span data-ttu-id="76c97-1506">Attributs</span><span class="sxs-lookup"><span data-stu-id="76c97-1506">Attributes</span></span>|<span data-ttu-id="76c97-1507">Description</span><span class="sxs-lookup"><span data-stu-id="76c97-1507">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="76c97-1508">String</span><span class="sxs-lookup"><span data-stu-id="76c97-1508">String</span></span>||<span data-ttu-id="76c97-p189">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="76c97-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="76c97-1512">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1512">Object</span></span>|<span data-ttu-id="76c97-1513">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1514">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="76c97-1514">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="76c97-1515">Objet</span><span class="sxs-lookup"><span data-stu-id="76c97-1515">Object</span></span>|<span data-ttu-id="76c97-1516">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1516">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-1517">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="76c97-1517">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="76c97-1518">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="76c97-1518">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="76c97-1519">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76c97-1519">&lt;optional&gt;</span></span>|<span data-ttu-id="76c97-p190">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="76c97-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="76c97-p191">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="76c97-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="76c97-1524">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="76c97-1524">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="76c97-1525">fonction</span><span class="sxs-lookup"><span data-stu-id="76c97-1525">function</span></span>||<span data-ttu-id="76c97-1526">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="76c97-1526">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76c97-1527">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76c97-1527">Requirements</span></span>

|<span data-ttu-id="76c97-1528">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76c97-1528">Requirement</span></span>|<span data-ttu-id="76c97-1529">Valeur</span><span class="sxs-lookup"><span data-stu-id="76c97-1529">Value</span></span>|
|---|---|
|[<span data-ttu-id="76c97-1530">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76c97-1530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="76c97-1531">1.2</span><span class="sxs-lookup"><span data-stu-id="76c97-1531">1.2</span></span>|
|[<span data-ttu-id="76c97-1532">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76c97-1532">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="76c97-1533">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="76c97-1533">ReadWriteItem</span></span>|
|[<span data-ttu-id="76c97-1534">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76c97-1534">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="76c97-1535">Composition</span><span class="sxs-lookup"><span data-stu-id="76c97-1535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="76c97-1536">Exemple</span><span class="sxs-lookup"><span data-stu-id="76c97-1536">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
