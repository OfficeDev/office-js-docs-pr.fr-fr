---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 94fcb090fb7e676ed0c30caef342e1d6cfc6e9f9
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600290"
---
# <a name="item"></a><span data-ttu-id="2e4a3-102">élément</span><span class="sxs-lookup"><span data-stu-id="2e4a3-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="2e4a3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="2e4a3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="2e4a3-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-106">Requirements</span></span>

|<span data-ttu-id="2e4a3-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-107">Requirement</span></span>|<span data-ttu-id="2e4a3-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-110">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-110">1.0</span></span>|
|[<span data-ttu-id="2e4a3-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="2e4a3-112">Restricted</span></span>|
|[<span data-ttu-id="2e4a3-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2e4a3-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="2e4a3-115">Members and methods</span></span>

| <span data-ttu-id="2e4a3-116">Membre	</span><span class="sxs-lookup"><span data-stu-id="2e4a3-116">Member</span></span> | <span data-ttu-id="2e4a3-117">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2e4a3-118">attachments</span><span class="sxs-lookup"><span data-stu-id="2e4a3-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="2e4a3-119">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-119">Member</span></span> |
| [<span data-ttu-id="2e4a3-120">bcc</span><span class="sxs-lookup"><span data-stu-id="2e4a3-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="2e4a3-121">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-121">Member</span></span> |
| [<span data-ttu-id="2e4a3-122">body</span><span class="sxs-lookup"><span data-stu-id="2e4a3-122">body</span></span>](#body-body) | <span data-ttu-id="2e4a3-123">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-123">Member</span></span> |
| [<span data-ttu-id="2e4a3-124">cc</span><span class="sxs-lookup"><span data-stu-id="2e4a3-124">cc</span></span>](#cc-arrayemailaddressdetails) | <span data-ttu-id="2e4a3-125">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-125">Member</span></span> |
| [<span data-ttu-id="2e4a3-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="2e4a3-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="2e4a3-127">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-127">Member</span></span> |
| [<span data-ttu-id="2e4a3-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="2e4a3-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="2e4a3-129">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-129">Member</span></span> |
| [<span data-ttu-id="2e4a3-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="2e4a3-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="2e4a3-131">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-131">Member</span></span> |
| [<span data-ttu-id="2e4a3-132">end</span><span class="sxs-lookup"><span data-stu-id="2e4a3-132">end</span></span>](#end-datetime) | <span data-ttu-id="2e4a3-133">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-133">Member</span></span> |
| [<span data-ttu-id="2e4a3-134">from</span><span class="sxs-lookup"><span data-stu-id="2e4a3-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="2e4a3-135">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-135">Member</span></span> |
| [<span data-ttu-id="2e4a3-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="2e4a3-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="2e4a3-137">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-137">Member</span></span> |
| [<span data-ttu-id="2e4a3-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="2e4a3-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="2e4a3-139">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-139">Member</span></span> |
| [<span data-ttu-id="2e4a3-140">itemId</span><span class="sxs-lookup"><span data-stu-id="2e4a3-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="2e4a3-141">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-141">Member</span></span> |
| [<span data-ttu-id="2e4a3-142">itemType</span><span class="sxs-lookup"><span data-stu-id="2e4a3-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="2e4a3-143">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-143">Member</span></span> |
| [<span data-ttu-id="2e4a3-144">location</span><span class="sxs-lookup"><span data-stu-id="2e4a3-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="2e4a3-145">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-145">Member</span></span> |
| [<span data-ttu-id="2e4a3-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="2e4a3-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="2e4a3-147">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-147">Member</span></span> |
| [<span data-ttu-id="2e4a3-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="2e4a3-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="2e4a3-149">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-149">Member</span></span> |
| [<span data-ttu-id="2e4a3-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="2e4a3-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetails) | <span data-ttu-id="2e4a3-151">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-151">Member</span></span> |
| [<span data-ttu-id="2e4a3-152">organizer</span><span class="sxs-lookup"><span data-stu-id="2e4a3-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="2e4a3-153">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-153">Member</span></span> |
| [<span data-ttu-id="2e4a3-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="2e4a3-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="2e4a3-155">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-155">Member</span></span> |
| [<span data-ttu-id="2e4a3-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="2e4a3-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetails) | <span data-ttu-id="2e4a3-157">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-157">Member</span></span> |
| [<span data-ttu-id="2e4a3-158">sender</span><span class="sxs-lookup"><span data-stu-id="2e4a3-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="2e4a3-159">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-159">Member</span></span> |
| [<span data-ttu-id="2e4a3-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="2e4a3-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="2e4a3-161">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-161">Member</span></span> |
| [<span data-ttu-id="2e4a3-162">start</span><span class="sxs-lookup"><span data-stu-id="2e4a3-162">start</span></span>](#start-datetime) | <span data-ttu-id="2e4a3-163">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-163">Member</span></span> |
| [<span data-ttu-id="2e4a3-164">subject</span><span class="sxs-lookup"><span data-stu-id="2e4a3-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="2e4a3-165">Member</span><span class="sxs-lookup"><span data-stu-id="2e4a3-165">Member</span></span> |
| [<span data-ttu-id="2e4a3-166">to</span><span class="sxs-lookup"><span data-stu-id="2e4a3-166">to</span></span>](#to-arrayemailaddressdetails) | <span data-ttu-id="2e4a3-167">Membre</span><span class="sxs-lookup"><span data-stu-id="2e4a3-167">Member</span></span> |
| [<span data-ttu-id="2e4a3-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2e4a3-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="2e4a3-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-169">Method</span></span> |
| [<span data-ttu-id="2e4a3-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2e4a3-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="2e4a3-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-171">Method</span></span> |
| [<span data-ttu-id="2e4a3-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2e4a3-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="2e4a3-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-173">Method</span></span> |
| [<span data-ttu-id="2e4a3-174">close</span><span class="sxs-lookup"><span data-stu-id="2e4a3-174">close</span></span>](#close) | <span data-ttu-id="2e4a3-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-175">Method</span></span> |
| [<span data-ttu-id="2e4a3-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="2e4a3-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="2e4a3-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-177">Method</span></span> |
| [<span data-ttu-id="2e4a3-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="2e4a3-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="2e4a3-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-179">Method</span></span> |
| [<span data-ttu-id="2e4a3-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="2e4a3-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="2e4a3-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-181">Method</span></span> |
| [<span data-ttu-id="2e4a3-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="2e4a3-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontact) | <span data-ttu-id="2e4a3-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-183">Method</span></span> |
| [<span data-ttu-id="2e4a3-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="2e4a3-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontact) | <span data-ttu-id="2e4a3-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-185">Method</span></span> |
| [<span data-ttu-id="2e4a3-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="2e4a3-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="2e4a3-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-187">Method</span></span> |
| [<span data-ttu-id="2e4a3-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="2e4a3-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="2e4a3-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-189">Method</span></span> |
| [<span data-ttu-id="2e4a3-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2e4a3-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="2e4a3-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-191">Method</span></span> |
| [<span data-ttu-id="2e4a3-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="2e4a3-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="2e4a3-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-193">Method</span></span> |
| [<span data-ttu-id="2e4a3-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="2e4a3-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="2e4a3-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-195">Method</span></span> |
| [<span data-ttu-id="2e4a3-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="2e4a3-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="2e4a3-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-197">Method</span></span> |
| [<span data-ttu-id="2e4a3-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2e4a3-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="2e4a3-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-199">Method</span></span> |
| [<span data-ttu-id="2e4a3-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2e4a3-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="2e4a3-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-201">Method</span></span> |
| [<span data-ttu-id="2e4a3-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="2e4a3-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="2e4a3-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-203">Method</span></span> |
| [<span data-ttu-id="2e4a3-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2e4a3-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="2e4a3-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="2e4a3-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="2e4a3-206">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-206">Example</span></span>

<span data-ttu-id="2e4a3-207">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="2e4a3-208">Membres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="2e4a3-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2e4a3-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="2e4a3-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-212">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="2e4a3-213">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-214">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-214">Type</span></span>

*   <span data-ttu-id="2e4a3-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2e4a3-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-216">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-216">Requirements</span></span>

|<span data-ttu-id="2e4a3-217">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-217">Requirement</span></span>|<span data-ttu-id="2e4a3-218">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-219">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-220">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-220">1.0</span></span>|
|[<span data-ttu-id="2e4a3-221">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-222">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-224">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-225">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-225">Example</span></span>

<span data-ttu-id="2e4a3-226">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2e4a3-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2e4a3-228">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="2e4a3-229">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-230">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-230">Type</span></span>

*   [<span data-ttu-id="2e4a3-231">Destinataires</span><span class="sxs-lookup"><span data-stu-id="2e4a3-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="2e4a3-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-232">Requirements</span></span>

|<span data-ttu-id="2e4a3-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-233">Requirement</span></span>|<span data-ttu-id="2e4a3-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-236">1.1</span><span class="sxs-lookup"><span data-stu-id="2e4a3-236">1.1</span></span>|
|[<span data-ttu-id="2e4a3-237">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-238">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-240">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-241">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="2e4a3-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="2e4a3-243">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-244">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-244">Type</span></span>

*   [<span data-ttu-id="2e4a3-245">Body</span><span class="sxs-lookup"><span data-stu-id="2e4a3-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="2e4a3-246">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-246">Requirements</span></span>

|<span data-ttu-id="2e4a3-247">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-247">Requirement</span></span>|<span data-ttu-id="2e4a3-248">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-249">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-250">1.1</span><span class="sxs-lookup"><span data-stu-id="2e4a3-250">1.1</span></span>|
|[<span data-ttu-id="2e4a3-251">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-252">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-253">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-254">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-255">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-255">Example</span></span>

<span data-ttu-id="2e4a3-256">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-256">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="2e4a3-257">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2e4a3-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2e4a3-259">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="2e4a3-260">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-261">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-261">Read mode</span></span>

<span data-ttu-id="2e4a3-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-264">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-264">Compose mode</span></span>

<span data-ttu-id="2e4a3-265">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2e4a3-266">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-266">Type</span></span>

*   <span data-ttu-id="2e4a3-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-268">Requirements</span></span>

|<span data-ttu-id="2e4a3-269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-269">Requirement</span></span>|<span data-ttu-id="2e4a3-270">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-272">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-272">1.0</span></span>|
|[<span data-ttu-id="2e4a3-273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-274">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-276">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-276">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="2e4a3-277">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-277">(nullable) conversationId :String</span></span>

<span data-ttu-id="2e4a3-278">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="2e4a3-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="2e4a3-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-283">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-283">Type</span></span>

*   <span data-ttu-id="2e4a3-284">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-285">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-285">Requirements</span></span>

|<span data-ttu-id="2e4a3-286">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-286">Requirement</span></span>|<span data-ttu-id="2e4a3-287">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-288">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-289">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-289">1.0</span></span>|
|[<span data-ttu-id="2e4a3-290">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-291">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-293">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-294">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-294">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="2e4a3-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="2e4a3-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="2e4a3-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-298">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-298">Type</span></span>

*   <span data-ttu-id="2e4a3-299">Date</span><span class="sxs-lookup"><span data-stu-id="2e4a3-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-300">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-300">Requirements</span></span>

|<span data-ttu-id="2e4a3-301">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-301">Requirement</span></span>|<span data-ttu-id="2e4a3-302">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-303">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-304">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-304">1.0</span></span>|
|[<span data-ttu-id="2e4a3-305">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-306">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-307">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-308">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-309">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-309">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="2e4a3-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="2e4a3-310">dateTimeModified :Date</span></span>

<span data-ttu-id="2e4a3-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-313">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-314">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-314">Type</span></span>

*   <span data-ttu-id="2e4a3-315">Date</span><span class="sxs-lookup"><span data-stu-id="2e4a3-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-316">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-316">Requirements</span></span>

|<span data-ttu-id="2e4a3-317">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-317">Requirement</span></span>|<span data-ttu-id="2e4a3-318">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-319">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-320">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-320">1.0</span></span>|
|[<span data-ttu-id="2e4a3-321">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-322">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-323">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-324">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-325">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-325">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="2e4a3-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="2e4a3-327">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="2e4a3-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-330">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-330">Read mode</span></span>

<span data-ttu-id="2e4a3-331">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-331">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-332">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-332">Compose mode</span></span>

<span data-ttu-id="2e4a3-333">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="2e4a3-334">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-334">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2e4a3-335">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2e4a3-336">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-336">Type</span></span>

*   <span data-ttu-id="2e4a3-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-338">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-338">Requirements</span></span>

|<span data-ttu-id="2e4a3-339">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-339">Requirement</span></span>|<span data-ttu-id="2e4a3-340">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-341">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-342">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-342">1.0</span></span>|
|[<span data-ttu-id="2e4a3-343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-344">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-346">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-346">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="2e4a3-347">from:[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[from](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-347">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="2e4a3-348">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="2e4a3-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-351">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-352">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-352">Read mode</span></span>

<span data-ttu-id="2e4a3-353">La `from` propriété renvoie un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-354">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-354">Compose mode</span></span>

<span data-ttu-id="2e4a3-355">La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2e4a3-356">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-356">Type</span></span>

*   <span data-ttu-id="2e4a3-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [à partir de](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-358">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-358">Requirements</span></span>

|<span data-ttu-id="2e4a3-359">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="2e4a3-360">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-361">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-361">1.0</span></span>|<span data-ttu-id="2e4a3-362">1.7</span><span class="sxs-lookup"><span data-stu-id="2e4a3-362">1.7</span></span>|
|[<span data-ttu-id="2e4a3-363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-364">ReadItem</span></span>|<span data-ttu-id="2e4a3-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="2e4a3-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-367">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-367">Read</span></span>|<span data-ttu-id="2e4a3-368">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-368">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="2e4a3-369">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-369">internetMessageId :String</span></span>

<span data-ttu-id="2e4a3-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-372">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-372">Type</span></span>

*   <span data-ttu-id="2e4a3-373">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-374">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-374">Requirements</span></span>

|<span data-ttu-id="2e4a3-375">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-375">Requirement</span></span>|<span data-ttu-id="2e4a3-376">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-377">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-378">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-378">1.0</span></span>|
|[<span data-ttu-id="2e4a3-379">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-380">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-381">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-382">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-383">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-383">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="2e4a3-384">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-384">itemClass :String</span></span>

<span data-ttu-id="2e4a3-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="2e4a3-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="2e4a3-389">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-389">Type</span></span>|<span data-ttu-id="2e4a3-390">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-390">Description</span></span>|<span data-ttu-id="2e4a3-391">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="2e4a3-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="2e4a3-392">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="2e4a3-392">Appointment items</span></span>|<span data-ttu-id="2e4a3-393">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="2e4a3-394">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="2e4a3-394">Message items</span></span>|<span data-ttu-id="2e4a3-395">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="2e4a3-396">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-397">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-397">Type</span></span>

*   <span data-ttu-id="2e4a3-398">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-399">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-399">Requirements</span></span>

|<span data-ttu-id="2e4a3-400">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-400">Requirement</span></span>|<span data-ttu-id="2e4a3-401">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-402">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-403">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-403">1.0</span></span>|
|[<span data-ttu-id="2e4a3-404">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-404">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-405">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-406">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-406">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-407">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-408">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-408">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="2e4a3-409">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-409">(nullable) itemId :String</span></span>

<span data-ttu-id="2e4a3-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-412">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2e4a3-413">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="2e4a3-414">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2e4a3-415">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="2e4a3-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-418">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-418">Type</span></span>

*   <span data-ttu-id="2e4a3-419">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-420">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-420">Requirements</span></span>

|<span data-ttu-id="2e4a3-421">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-421">Requirement</span></span>|<span data-ttu-id="2e4a3-422">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-423">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-424">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-424">1.0</span></span>|
|[<span data-ttu-id="2e4a3-425">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-426">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-427">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-428">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-429">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-429">Example</span></span>

<span data-ttu-id="2e4a3-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="2e4a3-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="2e4a3-433">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="2e4a3-434">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-435">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-435">Type</span></span>

*   [<span data-ttu-id="2e4a3-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="2e4a3-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="2e4a3-437">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-437">Requirements</span></span>

|<span data-ttu-id="2e4a3-438">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-438">Requirement</span></span>|<span data-ttu-id="2e4a3-439">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-440">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-441">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-441">1.0</span></span>|
|[<span data-ttu-id="2e4a3-442">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-442">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-443">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-444">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-444">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-445">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-446">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-446">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="2e4a3-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="2e4a3-448">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-449">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-449">Read mode</span></span>

<span data-ttu-id="2e4a3-450">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-451">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-451">Compose mode</span></span>

<span data-ttu-id="2e4a3-452">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2e4a3-453">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-453">Type</span></span>

*   <span data-ttu-id="2e4a3-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-455">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-455">Requirements</span></span>

|<span data-ttu-id="2e4a3-456">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-456">Requirement</span></span>|<span data-ttu-id="2e4a3-457">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-458">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-459">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-459">1.0</span></span>|
|[<span data-ttu-id="2e4a3-460">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-460">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-461">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-462">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-462">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-463">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-463">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="2e4a3-464">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-464">normalizedSubject :String</span></span>

<span data-ttu-id="2e4a3-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="2e4a3-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-469">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-469">Type</span></span>

*   <span data-ttu-id="2e4a3-470">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-471">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-471">Requirements</span></span>

|<span data-ttu-id="2e4a3-472">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-472">Requirement</span></span>|<span data-ttu-id="2e4a3-473">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-474">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-475">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-475">1.0</span></span>|
|[<span data-ttu-id="2e4a3-476">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-477">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-478">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-479">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-480">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-480">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="2e4a3-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="2e4a3-482">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-483">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-483">Type</span></span>

*   [<span data-ttu-id="2e4a3-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="2e4a3-484">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="2e4a3-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-485">Requirements</span></span>

|<span data-ttu-id="2e4a3-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-486">Requirement</span></span>|<span data-ttu-id="2e4a3-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-489">1.3</span><span class="sxs-lookup"><span data-stu-id="2e4a3-489">1.3</span></span>|
|[<span data-ttu-id="2e4a3-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-491">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-493">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-494">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-494">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2e4a3-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2e4a3-496">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="2e4a3-497">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-498">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-498">Read mode</span></span>

<span data-ttu-id="2e4a3-499">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-500">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-500">Compose mode</span></span>

<span data-ttu-id="2e4a3-501">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2e4a3-502">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-502">Type</span></span>

*   <span data-ttu-id="2e4a3-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-504">Requirements</span></span>

|<span data-ttu-id="2e4a3-505">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-505">Requirement</span></span>|<span data-ttu-id="2e4a3-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-507">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-508">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-508">1.0</span></span>|
|[<span data-ttu-id="2e4a3-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-510">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-512">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-512">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="2e4a3-513">Organisateur:[](/javascript/api/outlook_1_7/office.emailaddressdetails)|[organisateur](/javascript/api/outlook_1_7/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2e4a3-513">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="2e4a3-514">Obtient l'adresse de messagerie de l'organisateur d'une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-515">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-515">Read mode</span></span>

<span data-ttu-id="2e4a3-516">La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) qui représente l'organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-517">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-517">Compose mode</span></span>

<span data-ttu-id="2e4a3-518">La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook_1_7/office.organizer) qui fournit une méthode pour obtenir la valeur de l'organisateur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-518">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="2e4a3-519">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-519">Type</span></span>

*   <span data-ttu-id="2e4a3-520">[](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organisateur](/javascript/api/outlook_1_7/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2e4a3-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-521">Requirements</span></span>

|<span data-ttu-id="2e4a3-522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="2e4a3-523">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-524">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-524">1.0</span></span>|<span data-ttu-id="2e4a3-525">1.7</span><span class="sxs-lookup"><span data-stu-id="2e4a3-525">1.7</span></span>|
|[<span data-ttu-id="2e4a3-526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-527">ReadItem</span></span>|<span data-ttu-id="2e4a3-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="2e4a3-529">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-530">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-530">Read</span></span>|<span data-ttu-id="2e4a3-531">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-531">Compose</span></span>|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="2e4a3-532">(Nullable) récurrence:[périodicité](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="2e4a3-533">Obtient ou définit la périodicité d'un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="2e4a3-534">Obtient la périodicité d'une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="2e4a3-535">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="2e4a3-536">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="2e4a3-537">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook_1_7/office.recurrence) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="2e4a3-538">`null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="2e4a3-539">`undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="2e4a3-540">Remarque: les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="2e4a3-541">Remarque: si l'objet de périodicité `null`est, cela indique que l'objet est un rendez-vous unique ou une demande de réunion d'un seul rendez-vous et non d'une série.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-542">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-542">Read mode</span></span>

<span data-ttu-id="2e4a3-543">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook_1_7/office.recurrence) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="2e4a3-544">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-544">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-545">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-545">Compose mode</span></span>

<span data-ttu-id="2e4a3-546">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook_1_7/office.recurrence) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="2e4a3-547">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2e4a3-548">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-548">Type</span></span>

* [<span data-ttu-id="2e4a3-549">Instances</span><span class="sxs-lookup"><span data-stu-id="2e4a3-549">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="2e4a3-550">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-550">Requirement</span></span>|<span data-ttu-id="2e4a3-551">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-552">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-553">1.7</span><span class="sxs-lookup"><span data-stu-id="2e4a3-553">1.7</span></span>|
|[<span data-ttu-id="2e4a3-554">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-554">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-555">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-556">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-556">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-557">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-557">Compose or Read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2e4a3-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2e4a3-559">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="2e4a3-560">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-561">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-561">Read mode</span></span>

<span data-ttu-id="2e4a3-562">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-563">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-563">Compose mode</span></span>

<span data-ttu-id="2e4a3-564">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="2e4a3-565">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-565">Type</span></span>

*   <span data-ttu-id="2e4a3-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-567">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-567">Requirements</span></span>

|<span data-ttu-id="2e4a3-568">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-568">Requirement</span></span>|<span data-ttu-id="2e4a3-569">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-570">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-571">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-571">1.0</span></span>|
|[<span data-ttu-id="2e4a3-572">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-573">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-574">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-575">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-575">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="2e4a3-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="2e4a3-p128">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="2e4a3-p129">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p129">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-581">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-582">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-582">Type</span></span>

*   [<span data-ttu-id="2e4a3-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2e4a3-583">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="2e4a3-584">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-584">Requirements</span></span>

|<span data-ttu-id="2e4a3-585">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-585">Requirement</span></span>|<span data-ttu-id="2e4a3-586">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-587">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-588">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-588">1.0</span></span>|
|[<span data-ttu-id="2e4a3-589">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-590">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-591">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-592">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-593">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-593">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="2e4a3-594">(Nullable) seriesId: chaîne</span><span class="sxs-lookup"><span data-stu-id="2e4a3-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="2e4a3-595">Obtient l'ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="2e4a3-596">Dans OWA et Outlook, le `seriesId` renvoie l'ID des services Web Exchange (EWS) de l'élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="2e4a3-597">Toutefois, dans iOS et Android, le `seriesId` renvoie l'ID REST de l'élément parent.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-598">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2e4a3-599">La `seriesId` propriété n'est pas identique aux ID Outlook utilisés par l'API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="2e4a3-600">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2e4a3-601">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="2e4a3-602">La `seriesId` propriété renvoie `null` pour les éléments qui n'ont pas d'éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="2e4a3-603">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-603">Type</span></span>

* <span data-ttu-id="2e4a3-604">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-605">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-605">Requirements</span></span>

|<span data-ttu-id="2e4a3-606">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-606">Requirement</span></span>|<span data-ttu-id="2e4a3-607">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-608">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-609">1.7</span><span class="sxs-lookup"><span data-stu-id="2e4a3-609">1.7</span></span>|
|[<span data-ttu-id="2e4a3-610">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-611">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-612">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-613">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-614">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-614">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="2e4a3-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="2e4a3-616">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="2e4a3-p132">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-619">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-619">Read mode</span></span>

<span data-ttu-id="2e4a3-620">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-620">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-621">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-621">Compose mode</span></span>

<span data-ttu-id="2e4a3-622">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="2e4a3-623">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-623">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2e4a3-624">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2e4a3-625">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-625">Type</span></span>

*   <span data-ttu-id="2e4a3-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-627">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-627">Requirements</span></span>

|<span data-ttu-id="2e4a3-628">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-628">Requirement</span></span>|<span data-ttu-id="2e4a3-629">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-630">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-631">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-631">1.0</span></span>|
|[<span data-ttu-id="2e4a3-632">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-632">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-633">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-634">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-634">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-635">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-635">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="2e4a3-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="2e4a3-637">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="2e4a3-638">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-639">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-639">Read mode</span></span>

<span data-ttu-id="2e4a3-p133">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="2e4a3-642">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-643">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-643">Compose mode</span></span>

<span data-ttu-id="2e4a3-644">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="2e4a3-645">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-645">Type</span></span>

*   <span data-ttu-id="2e4a3-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-647">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-647">Requirements</span></span>

|<span data-ttu-id="2e4a3-648">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-648">Requirement</span></span>|<span data-ttu-id="2e4a3-649">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-650">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-651">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-651">1.0</span></span>|
|[<span data-ttu-id="2e4a3-652">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-653">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-654">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-655">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-655">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2e4a3-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2e4a3-657">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="2e4a3-658">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2e4a3-659">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-659">Read mode</span></span>

<span data-ttu-id="2e4a3-p135">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="2e4a3-662">Mode composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-662">Compose mode</span></span>

<span data-ttu-id="2e4a3-663">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2e4a3-664">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-664">Type</span></span>

*   <span data-ttu-id="2e4a3-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-666">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-666">Requirements</span></span>

|<span data-ttu-id="2e4a3-667">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-667">Requirement</span></span>|<span data-ttu-id="2e4a3-668">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-669">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-670">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-670">1.0</span></span>|
|[<span data-ttu-id="2e4a3-671">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-672">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-673">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-674">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="2e4a3-675">Méthodes</span><span class="sxs-lookup"><span data-stu-id="2e4a3-675">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="2e4a3-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2e4a3-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2e4a3-677">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2e4a3-678">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="2e4a3-679">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-680">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-680">Parameters</span></span>
|<span data-ttu-id="2e4a3-681">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-681">Name</span></span>|<span data-ttu-id="2e4a3-682">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-682">Type</span></span>|<span data-ttu-id="2e4a3-683">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-683">Attributes</span></span>|<span data-ttu-id="2e4a3-684">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="2e4a3-685">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-685">String</span></span>||<span data-ttu-id="2e4a3-p136">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="2e4a3-688">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-688">String</span></span>||<span data-ttu-id="2e4a3-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2e4a3-691">Object</span><span class="sxs-lookup"><span data-stu-id="2e4a3-691">Object</span></span>|<span data-ttu-id="2e4a3-692">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-692">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-693">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2e4a3-694">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-694">Object</span></span>|<span data-ttu-id="2e4a3-695">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-695">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-696">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="2e4a3-697">Boolean</span><span class="sxs-lookup"><span data-stu-id="2e4a3-697">Boolean</span></span>|<span data-ttu-id="2e4a3-698">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-698">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-699">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="2e4a3-700">fonction</span><span class="sxs-lookup"><span data-stu-id="2e4a3-700">function</span></span>|<span data-ttu-id="2e4a3-701">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-701">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-702">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2e4a3-703">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2e4a3-704">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2e4a3-705">Erreurs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-705">Errors</span></span>

|<span data-ttu-id="2e4a3-706">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-706">Error code</span></span>|<span data-ttu-id="2e4a3-707">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="2e4a3-708">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="2e4a3-709">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2e4a3-710">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-711">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-711">Requirements</span></span>

|<span data-ttu-id="2e4a3-712">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-712">Requirement</span></span>|<span data-ttu-id="2e4a3-713">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-714">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-715">1.1</span><span class="sxs-lookup"><span data-stu-id="2e4a3-715">1.1</span></span>|
|[<span data-ttu-id="2e4a3-716">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-716">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="2e4a3-718">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-718">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-719">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2e4a3-720">Exemples</span><span class="sxs-lookup"><span data-stu-id="2e4a3-720">Examples</span></span>

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

<span data-ttu-id="2e4a3-721">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="2e4a3-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2e4a3-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="2e4a3-723">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="2e4a3-724">Actuellement, les types d'événement `Office.EventType.AppointmentTimeChanged`pris `Office.EventType.RecipientsChanged`en charge sont, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="2e4a3-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-725">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-725">Parameters</span></span>

| <span data-ttu-id="2e4a3-726">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-726">Name</span></span> | <span data-ttu-id="2e4a3-727">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-727">Type</span></span> | <span data-ttu-id="2e4a3-728">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-728">Attributes</span></span> | <span data-ttu-id="2e4a3-729">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2e4a3-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2e4a3-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2e4a3-731">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="2e4a3-732">Fonction</span><span class="sxs-lookup"><span data-stu-id="2e4a3-732">Function</span></span> || <span data-ttu-id="2e4a3-p138">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="2e4a3-736">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-736">Object</span></span> | <span data-ttu-id="2e4a3-737">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-737">&lt;optional&gt;</span></span> | <span data-ttu-id="2e4a3-738">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2e4a3-739">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-739">Object</span></span> | <span data-ttu-id="2e4a3-740">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-740">&lt;optional&gt;</span></span> | <span data-ttu-id="2e4a3-741">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2e4a3-742">fonction</span><span class="sxs-lookup"><span data-stu-id="2e4a3-742">function</span></span>| <span data-ttu-id="2e4a3-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-743">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-744">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-745">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-745">Requirements</span></span>

|<span data-ttu-id="2e4a3-746">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-746">Requirement</span></span>| <span data-ttu-id="2e4a3-747">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-748">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e4a3-749">1.7</span><span class="sxs-lookup"><span data-stu-id="2e4a3-749">1.7</span></span> |
|[<span data-ttu-id="2e4a3-750">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e4a3-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-751">ReadItem</span></span> |
|[<span data-ttu-id="2e4a3-752">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e4a3-753">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="2e4a3-754">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-754">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="2e4a3-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2e4a3-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2e4a3-756">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="2e4a3-p139">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="2e4a3-760">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="2e4a3-761">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-761">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-762">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-762">Parameters</span></span>

|<span data-ttu-id="2e4a3-763">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-763">Name</span></span>|<span data-ttu-id="2e4a3-764">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-764">Type</span></span>|<span data-ttu-id="2e4a3-765">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-765">Attributes</span></span>|<span data-ttu-id="2e4a3-766">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="2e4a3-767">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-767">String</span></span>||<span data-ttu-id="2e4a3-p140">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="2e4a3-770">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-770">String</span></span>||<span data-ttu-id="2e4a3-771">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-771">The subject of the item to be attached.</span></span> <span data-ttu-id="2e4a3-772">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2e4a3-773">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-773">Object</span></span>|<span data-ttu-id="2e4a3-774">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-774">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-775">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2e4a3-776">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-776">Object</span></span>|<span data-ttu-id="2e4a3-777">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-777">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-778">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2e4a3-779">fonction</span><span class="sxs-lookup"><span data-stu-id="2e4a3-779">function</span></span>|<span data-ttu-id="2e4a3-780">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-780">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-781">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2e4a3-782">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2e4a3-783">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2e4a3-784">Erreurs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-784">Errors</span></span>

|<span data-ttu-id="2e4a3-785">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-785">Error code</span></span>|<span data-ttu-id="2e4a3-786">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2e4a3-787">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-788">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-788">Requirements</span></span>

|<span data-ttu-id="2e4a3-789">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-789">Requirement</span></span>|<span data-ttu-id="2e4a3-790">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-791">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-792">1.1</span><span class="sxs-lookup"><span data-stu-id="2e4a3-792">1.1</span></span>|
|[<span data-ttu-id="2e4a3-793">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-793">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="2e4a3-795">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-795">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-796">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-797">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-797">Example</span></span>

<span data-ttu-id="2e4a3-798">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="2e4a3-799">close()</span><span class="sxs-lookup"><span data-stu-id="2e4a3-799">close()</span></span>

<span data-ttu-id="2e4a3-800">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="2e4a3-p142">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-803">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="2e4a3-804">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-805">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-805">Requirements</span></span>

|<span data-ttu-id="2e4a3-806">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-806">Requirement</span></span>|<span data-ttu-id="2e4a3-807">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-808">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-809">1.3</span><span class="sxs-lookup"><span data-stu-id="2e4a3-809">1.3</span></span>|
|[<span data-ttu-id="2e4a3-810">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-811">Restreinte</span><span class="sxs-lookup"><span data-stu-id="2e4a3-811">Restricted</span></span>|
|[<span data-ttu-id="2e4a3-812">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-813">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-813">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="2e4a3-814">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2e4a3-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="2e4a3-815">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-816">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-816">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e4a3-817">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-817">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2e4a3-818">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="2e4a3-p143">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-822">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-822">Parameters</span></span>

|<span data-ttu-id="2e4a3-823">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-823">Name</span></span>|<span data-ttu-id="2e4a3-824">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-824">Type</span></span>|<span data-ttu-id="2e4a3-825">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-825">Attributes</span></span>|<span data-ttu-id="2e4a3-826">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="2e4a3-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2e4a3-827">String &#124; Object</span></span>||<span data-ttu-id="2e4a3-828">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-828">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="2e4a3-829">La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-829">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2e4a3-830">**OU**</span><span class="sxs-lookup"><span data-stu-id="2e4a3-830">**OR**</span></span><br/><span data-ttu-id="2e4a3-p145">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="2e4a3-833">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-833">String</span></span>|<span data-ttu-id="2e4a3-834">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-834">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="2e4a3-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="2e4a3-838">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-838">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-839">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="2e4a3-840">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2e4a3-840">String</span></span>||<span data-ttu-id="2e4a3-p147">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="2e4a3-843">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-843">String</span></span>||<span data-ttu-id="2e4a3-844">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="2e4a3-845">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2e4a3-845">String</span></span>||<span data-ttu-id="2e4a3-p148">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="2e4a3-848">Booléen</span><span class="sxs-lookup"><span data-stu-id="2e4a3-848">Boolean</span></span>||<span data-ttu-id="2e4a3-p149">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="2e4a3-851">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-851">String</span></span>||<span data-ttu-id="2e4a3-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="2e4a3-855">function</span><span class="sxs-lookup"><span data-stu-id="2e4a3-855">function</span></span>|<span data-ttu-id="2e4a3-856">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-856">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-857">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-858">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-858">Requirements</span></span>

|<span data-ttu-id="2e4a3-859">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-859">Requirement</span></span>|<span data-ttu-id="2e4a3-860">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-861">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-862">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-862">1.0</span></span>|
|[<span data-ttu-id="2e4a3-863">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-864">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-865">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-866">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2e4a3-867">Exemples</span><span class="sxs-lookup"><span data-stu-id="2e4a3-867">Examples</span></span>

<span data-ttu-id="2e4a3-868">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="2e4a3-869">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-869">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="2e4a3-870">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-870">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2e4a3-871">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2e4a3-872">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2e4a3-873">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="2e4a3-874">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2e4a3-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="2e4a3-875">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-876">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-876">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e4a3-877">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-877">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2e4a3-878">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="2e4a3-p151">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-882">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-882">Parameters</span></span>

|<span data-ttu-id="2e4a3-883">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-883">Name</span></span>|<span data-ttu-id="2e4a3-884">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-884">Type</span></span>|<span data-ttu-id="2e4a3-885">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-885">Attributes</span></span>|<span data-ttu-id="2e4a3-886">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="2e4a3-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2e4a3-887">String &#124; Object</span></span>||<span data-ttu-id="2e4a3-888">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-888">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="2e4a3-889">La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-889">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2e4a3-890">**OU**</span><span class="sxs-lookup"><span data-stu-id="2e4a3-890">**OR**</span></span><br/><span data-ttu-id="2e4a3-p153">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="2e4a3-893">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-893">String</span></span>|<span data-ttu-id="2e4a3-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-894">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="2e4a3-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="2e4a3-898">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-898">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-899">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="2e4a3-900">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2e4a3-900">String</span></span>||<span data-ttu-id="2e4a3-p155">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="2e4a3-903">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-903">String</span></span>||<span data-ttu-id="2e4a3-904">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="2e4a3-905">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2e4a3-905">String</span></span>||<span data-ttu-id="2e4a3-p156">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="2e4a3-908">Booléen</span><span class="sxs-lookup"><span data-stu-id="2e4a3-908">Boolean</span></span>||<span data-ttu-id="2e4a3-p157">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="2e4a3-911">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-911">String</span></span>||<span data-ttu-id="2e4a3-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="2e4a3-915">function</span><span class="sxs-lookup"><span data-stu-id="2e4a3-915">function</span></span>|<span data-ttu-id="2e4a3-916">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-916">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-917">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-918">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-918">Requirements</span></span>

|<span data-ttu-id="2e4a3-919">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-919">Requirement</span></span>|<span data-ttu-id="2e4a3-920">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-921">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-922">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-922">1.0</span></span>|
|[<span data-ttu-id="2e4a3-923">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-923">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-924">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-925">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-925">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-926">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2e4a3-927">Exemples</span><span class="sxs-lookup"><span data-stu-id="2e4a3-927">Examples</span></span>

<span data-ttu-id="2e4a3-928">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="2e4a3-929">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-929">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="2e4a3-930">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-930">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2e4a3-931">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2e4a3-932">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2e4a3-933">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="2e4a3-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="2e4a3-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="2e4a3-935">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-936">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-936">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-937">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-937">Requirements</span></span>

|<span data-ttu-id="2e4a3-938">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-938">Requirement</span></span>|<span data-ttu-id="2e4a3-939">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-940">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-941">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-941">1.0</span></span>|
|[<span data-ttu-id="2e4a3-942">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-942">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-943">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-944">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-944">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-945">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e4a3-946">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-946">Returns:</span></span>

<span data-ttu-id="2e4a3-947">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-947">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="2e4a3-948">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-948">Example</span></span>

<span data-ttu-id="2e4a3-949">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-949">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="2e4a3-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2e4a3-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2e4a3-951">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-952">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-952">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-953">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-953">Parameters</span></span>

|<span data-ttu-id="2e4a3-954">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-954">Name</span></span>|<span data-ttu-id="2e4a3-955">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-955">Type</span></span>|<span data-ttu-id="2e4a3-956">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="2e4a3-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="2e4a3-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="2e4a3-958">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-959">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-959">Requirements</span></span>

|<span data-ttu-id="2e4a3-960">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-960">Requirement</span></span>|<span data-ttu-id="2e4a3-961">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-962">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-963">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-963">1.0</span></span>|
|[<span data-ttu-id="2e4a3-964">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-965">Restreinte</span><span class="sxs-lookup"><span data-stu-id="2e4a3-965">Restricted</span></span>|
|[<span data-ttu-id="2e4a3-966">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-967">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e4a3-968">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-968">Returns:</span></span>

<span data-ttu-id="2e4a3-969">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="2e4a3-970">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="2e4a3-971">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="2e4a3-972">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="2e4a3-973">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="2e4a3-973">Value of `entityType`</span></span>|<span data-ttu-id="2e4a3-974">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="2e4a3-974">Type of objects in returned array</span></span>|<span data-ttu-id="2e4a3-975">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="2e4a3-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="2e4a3-976">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-976">String</span></span>|<span data-ttu-id="2e4a3-977">**Restreinte**</span><span class="sxs-lookup"><span data-stu-id="2e4a3-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="2e4a3-978">Contact</span><span class="sxs-lookup"><span data-stu-id="2e4a3-978">Contact</span></span>|<span data-ttu-id="2e4a3-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2e4a3-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="2e4a3-980">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-980">String</span></span>|<span data-ttu-id="2e4a3-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2e4a3-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="2e4a3-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="2e4a3-982">MeetingSuggestion</span></span>|<span data-ttu-id="2e4a3-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2e4a3-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="2e4a3-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="2e4a3-984">PhoneNumber</span></span>|<span data-ttu-id="2e4a3-985">**Restreinte**</span><span class="sxs-lookup"><span data-stu-id="2e4a3-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="2e4a3-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="2e4a3-986">TaskSuggestion</span></span>|<span data-ttu-id="2e4a3-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2e4a3-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="2e4a3-988">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-988">String</span></span>|<span data-ttu-id="2e4a3-989">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="2e4a3-989">**Restricted**</span></span>|

<span data-ttu-id="2e4a3-990">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2e4a3-990">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="2e4a3-991">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-991">Example</span></span>

<span data-ttu-id="2e4a3-992">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="2e4a3-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2e4a3-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2e4a3-994">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-995">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-995">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e4a3-996">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-997">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-997">Parameters</span></span>

|<span data-ttu-id="2e4a3-998">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-998">Name</span></span>|<span data-ttu-id="2e4a3-999">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-999">Type</span></span>|<span data-ttu-id="2e4a3-1000">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="2e4a3-1001">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1001">String</span></span>|<span data-ttu-id="2e4a3-1002">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1003">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1003">Requirements</span></span>

|<span data-ttu-id="2e4a3-1004">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1004">Requirement</span></span>|<span data-ttu-id="2e4a3-1005">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1006">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1007">1.0</span></span>|
|[<span data-ttu-id="2e4a3-1008">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1009">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-1010">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1011">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e4a3-1012">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1012">Returns:</span></span>

<span data-ttu-id="2e4a3-p160">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="2e4a3-1015">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2e4a3-1015">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="2e4a3-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="2e4a3-1017">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-1018">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1018">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e4a3-p161">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2e4a3-1022">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2e4a3-1023">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2e4a3-p162">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1027">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1027">Requirements</span></span>

|<span data-ttu-id="2e4a3-1028">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1028">Requirement</span></span>|<span data-ttu-id="2e4a3-1029">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1030">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1031">1.0</span></span>|
|[<span data-ttu-id="2e4a3-1032">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1032">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1033">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-1034">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1034">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1035">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e4a3-1036">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1036">Returns:</span></span>

<span data-ttu-id="2e4a3-p163">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="2e4a3-1039">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1039">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2e4a3-1040">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1040">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2e4a3-1041">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1041">Example</span></span>

<span data-ttu-id="2e4a3-1042">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1042">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="2e4a3-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="2e4a3-1044">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1044">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-1045">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1045">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e4a3-1046">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1046">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="2e4a3-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-1049">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1049">Parameters</span></span>

|<span data-ttu-id="2e4a3-1050">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1050">Name</span></span>|<span data-ttu-id="2e4a3-1051">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1051">Type</span></span>|<span data-ttu-id="2e4a3-1052">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1052">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="2e4a3-1053">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1053">String</span></span>|<span data-ttu-id="2e4a3-1054">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1054">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1055">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1055">Requirements</span></span>

|<span data-ttu-id="2e4a3-1056">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1056">Requirement</span></span>|<span data-ttu-id="2e4a3-1057">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1057">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1058">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1058">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1059">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1059">1.0</span></span>|
|[<span data-ttu-id="2e4a3-1060">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1060">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1061">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1061">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-1062">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1062">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1063">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1063">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e4a3-1064">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1064">Returns:</span></span>

<span data-ttu-id="2e4a3-1065">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1065">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="2e4a3-1066">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1066">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2e4a3-1067">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="2e4a3-1067">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2e4a3-1068">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1068">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="2e4a3-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="2e4a3-1070">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1070">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="2e4a3-p165">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-1073">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1073">Parameters</span></span>

|<span data-ttu-id="2e4a3-1074">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1074">Name</span></span>|<span data-ttu-id="2e4a3-1075">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1075">Type</span></span>|<span data-ttu-id="2e4a3-1076">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1076">Attributes</span></span>|<span data-ttu-id="2e4a3-1077">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1077">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="2e4a3-1078">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1078">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="2e4a3-p166">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="2e4a3-1082">Object</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1082">Object</span></span>|<span data-ttu-id="2e4a3-1083">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1084">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2e4a3-1085">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1085">Object</span></span>|<span data-ttu-id="2e4a3-1086">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1087">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2e4a3-1088">fonction</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1088">function</span></span>||<span data-ttu-id="2e4a3-1089">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2e4a3-1090">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1090">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="2e4a3-1091">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1091">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1092">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1092">Requirements</span></span>

|<span data-ttu-id="2e4a3-1093">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1093">Requirement</span></span>|<span data-ttu-id="2e4a3-1094">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1094">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1095">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1095">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1096">1.2</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1096">1.2</span></span>|
|[<span data-ttu-id="2e4a3-1097">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1097">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1098">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1098">ReadWriteItem</span></span>|
|[<span data-ttu-id="2e4a3-1099">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1099">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1100">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1100">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e4a3-1101">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1101">Returns:</span></span>

<span data-ttu-id="2e4a3-1102">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1102">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="2e4a3-1103">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1103">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2e4a3-1104">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1104">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2e4a3-1105">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1105">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="2e4a3-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="2e4a3-1107">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1107">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="2e4a3-1108">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1108">Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-1109">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1109">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1110">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1110">Requirements</span></span>

|<span data-ttu-id="2e4a3-1111">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1111">Requirement</span></span>|<span data-ttu-id="2e4a3-1112">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1113">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1114">1.6</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1114">1.6</span></span>|
|[<span data-ttu-id="2e4a3-1115">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1115">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1116">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1116">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-1117">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1117">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1118">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1118">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e4a3-1119">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1119">Returns:</span></span>

<span data-ttu-id="2e4a3-1120">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1120">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="2e4a3-1121">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1121">Example</span></span>

<span data-ttu-id="2e4a3-1122">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1122">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="2e4a3-1123">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1123">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="2e4a3-p169">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-1126">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e4a3-p170">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2e4a3-1130">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1130">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2e4a3-1131">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1131">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2e4a3-p171">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1135">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1135">Requirements</span></span>

|<span data-ttu-id="2e4a3-1136">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1136">Requirement</span></span>|<span data-ttu-id="2e4a3-1137">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1138">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1139">1.6</span></span>|
|[<span data-ttu-id="2e4a3-1140">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1140">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1141">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-1142">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1142">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1143">Lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e4a3-1144">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1144">Returns:</span></span>

<span data-ttu-id="2e4a3-p172">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="2e4a3-1147">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1147">Example</span></span>

<span data-ttu-id="2e4a3-1148">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1148">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="2e4a3-1149">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1149">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="2e4a3-1150">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1150">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="2e4a3-p173">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-1154">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1154">Parameters</span></span>

|<span data-ttu-id="2e4a3-1155">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1155">Name</span></span>|<span data-ttu-id="2e4a3-1156">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1156">Type</span></span>|<span data-ttu-id="2e4a3-1157">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1157">Attributes</span></span>|<span data-ttu-id="2e4a3-1158">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1158">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="2e4a3-1159">function</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1159">function</span></span>||<span data-ttu-id="2e4a3-1160">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1160">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2e4a3-1161">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1161">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2e4a3-1162">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1162">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="2e4a3-1163">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1163">Object</span></span>|<span data-ttu-id="2e4a3-1164">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1164">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1165">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1165">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="2e4a3-1166">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1166">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1167">Requirements</span></span>

|<span data-ttu-id="2e4a3-1168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1168">Requirement</span></span>|<span data-ttu-id="2e4a3-1169">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1171">1.0</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1171">1.0</span></span>|
|[<span data-ttu-id="2e4a3-1172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1173">ReadItem</span></span>|
|[<span data-ttu-id="2e4a3-1174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1175">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-1176">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1176">Example</span></span>

<span data-ttu-id="2e4a3-p176">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="2e4a3-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="2e4a3-1181">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1181">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="2e4a3-p177">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-1186">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1186">Parameters</span></span>

|<span data-ttu-id="2e4a3-1187">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1187">Name</span></span>|<span data-ttu-id="2e4a3-1188">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1188">Type</span></span>|<span data-ttu-id="2e4a3-1189">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1189">Attributes</span></span>|<span data-ttu-id="2e4a3-1190">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1190">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="2e4a3-1191">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1191">String</span></span>||<span data-ttu-id="2e4a3-1192">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1192">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="2e4a3-1193">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1193">Object</span></span>|<span data-ttu-id="2e4a3-1194">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1195">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1195">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2e4a3-1196">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1196">Object</span></span>|<span data-ttu-id="2e4a3-1197">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1198">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1198">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2e4a3-1199">fonction</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1199">function</span></span>|<span data-ttu-id="2e4a3-1200">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1201">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1201">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2e4a3-1202">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1202">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2e4a3-1203">Erreurs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1203">Errors</span></span>

|<span data-ttu-id="2e4a3-1204">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1204">Error code</span></span>|<span data-ttu-id="2e4a3-1205">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1205">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="2e4a3-1206">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1206">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1207">Requirements</span></span>

|<span data-ttu-id="2e4a3-1208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1208">Requirement</span></span>|<span data-ttu-id="2e4a3-1209">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1209">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1211">1.1</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1211">1.1</span></span>|
|[<span data-ttu-id="2e4a3-1212">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1213">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1213">ReadWriteItem</span></span>|
|[<span data-ttu-id="2e4a3-1214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1215">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1215">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-1216">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1216">Example</span></span>

<span data-ttu-id="2e4a3-1217">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1217">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="2e4a3-1218">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1218">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="2e4a3-1219">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1219">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="2e4a3-1220">Actuellement, les types d'événement `Office.EventType.AppointmentTimeChanged`pris `Office.EventType.RecipientsChanged`en charge sont, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1220">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-1221">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1221">Parameters</span></span>

| <span data-ttu-id="2e4a3-1222">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1222">Name</span></span> | <span data-ttu-id="2e4a3-1223">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1223">Type</span></span> | <span data-ttu-id="2e4a3-1224">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1224">Attributes</span></span> | <span data-ttu-id="2e4a3-1225">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1225">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2e4a3-1226">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1226">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2e4a3-1227">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1227">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="2e4a3-1228">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1228">Object</span></span> | <span data-ttu-id="2e4a3-1229">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="2e4a3-1230">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2e4a3-1231">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1231">Object</span></span> | <span data-ttu-id="2e4a3-1232">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="2e4a3-1233">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2e4a3-1234">fonction</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1234">function</span></span>| <span data-ttu-id="2e4a3-1235">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1236">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1237">Requirements</span></span>

|<span data-ttu-id="2e4a3-1238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1238">Requirement</span></span>| <span data-ttu-id="2e4a3-1239">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e4a3-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1241">1.7</span></span> |
|[<span data-ttu-id="2e4a3-1242">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e4a3-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1243">ReadItem</span></span> |
|[<span data-ttu-id="2e4a3-1244">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e4a3-1245">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1245">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="2e4a3-1246">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1246">Example</span></span>

```javascript
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="2e4a3-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="2e4a3-1248">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="2e4a3-p178">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-1252">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="2e4a3-1253">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="2e4a3-p180">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="2e4a3-1257">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="2e4a3-1258">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="2e4a3-1259">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="2e4a3-1260">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-1261">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1261">Parameters</span></span>

|<span data-ttu-id="2e4a3-1262">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1262">Name</span></span>|<span data-ttu-id="2e4a3-1263">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1263">Type</span></span>|<span data-ttu-id="2e4a3-1264">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1264">Attributes</span></span>|<span data-ttu-id="2e4a3-1265">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2e4a3-1266">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1266">Object</span></span>|<span data-ttu-id="2e4a3-1267">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1268">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2e4a3-1269">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1269">Object</span></span>|<span data-ttu-id="2e4a3-1270">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1271">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2e4a3-1272">fonction</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1272">function</span></span>||<span data-ttu-id="2e4a3-1273">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2e4a3-1274">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1275">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1275">Requirements</span></span>

|<span data-ttu-id="2e4a3-1276">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1276">Requirement</span></span>|<span data-ttu-id="2e4a3-1277">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1278">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1279">1.3</span></span>|
|[<span data-ttu-id="2e4a3-1280">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="2e4a3-1282">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1283">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2e4a3-1284">範例</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1284">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="2e4a3-p182">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="2e4a3-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="2e4a3-1288">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="2e4a3-p183">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e4a3-1292">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1292">Parameters</span></span>

|<span data-ttu-id="2e4a3-1293">Nom</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1293">Name</span></span>|<span data-ttu-id="2e4a3-1294">Type</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1294">Type</span></span>|<span data-ttu-id="2e4a3-1295">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1295">Attributes</span></span>|<span data-ttu-id="2e4a3-1296">Description</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="2e4a3-1297">String</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1297">String</span></span>||<span data-ttu-id="2e4a3-p184">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="2e4a3-1301">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1301">Object</span></span>|<span data-ttu-id="2e4a3-1302">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1303">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2e4a3-1304">Objet</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1304">Object</span></span>|<span data-ttu-id="2e4a3-1305">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-1306">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="2e4a3-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="2e4a3-1308">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="2e4a3-p185">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="2e4a3-p186">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="2e4a3-1313">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="2e4a3-1314">fonction</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1314">function</span></span>||<span data-ttu-id="2e4a3-1315">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e4a3-1316">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1316">Requirements</span></span>

|<span data-ttu-id="2e4a3-1317">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1317">Requirement</span></span>|<span data-ttu-id="2e4a3-1318">Valeur</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e4a3-1319">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2e4a3-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1320">1.2</span></span>|
|[<span data-ttu-id="2e4a3-1321">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2e4a3-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="2e4a3-1323">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2e4a3-1324">Composition</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2e4a3-1325">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a3-1325">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
