---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,7
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: d400765293449899eb2e26f3d87128bc88b70000
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629678"
---
# <a name="item"></a><span data-ttu-id="d240f-102">élément</span><span class="sxs-lookup"><span data-stu-id="d240f-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d240f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d240f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d240f-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="d240f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-106">Requirements</span></span>

|<span data-ttu-id="d240f-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-107">Requirement</span></span>|<span data-ttu-id="d240f-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-110">1.0</span></span>|
|[<span data-ttu-id="d240f-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="d240f-112">Restricted</span></span>|
|[<span data-ttu-id="d240f-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d240f-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="d240f-115">Members and methods</span></span>

| <span data-ttu-id="d240f-116">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-116">Member</span></span> | <span data-ttu-id="d240f-117">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d240f-118">attachments</span><span class="sxs-lookup"><span data-stu-id="d240f-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="d240f-119">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-119">Member</span></span> |
| [<span data-ttu-id="d240f-120">bcc</span><span class="sxs-lookup"><span data-stu-id="d240f-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="d240f-121">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-121">Member</span></span> |
| [<span data-ttu-id="d240f-122">body</span><span class="sxs-lookup"><span data-stu-id="d240f-122">body</span></span>](#body-body) | <span data-ttu-id="d240f-123">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-123">Member</span></span> |
| [<span data-ttu-id="d240f-124">cc</span><span class="sxs-lookup"><span data-stu-id="d240f-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d240f-125">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-125">Member</span></span> |
| [<span data-ttu-id="d240f-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="d240f-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="d240f-127">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-127">Member</span></span> |
| [<span data-ttu-id="d240f-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="d240f-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="d240f-129">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-129">Member</span></span> |
| [<span data-ttu-id="d240f-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="d240f-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="d240f-131">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-131">Member</span></span> |
| [<span data-ttu-id="d240f-132">end</span><span class="sxs-lookup"><span data-stu-id="d240f-132">end</span></span>](#end-datetime) | <span data-ttu-id="d240f-133">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-133">Member</span></span> |
| [<span data-ttu-id="d240f-134">from</span><span class="sxs-lookup"><span data-stu-id="d240f-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="d240f-135">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-135">Member</span></span> |
| [<span data-ttu-id="d240f-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="d240f-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="d240f-137">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-137">Member</span></span> |
| [<span data-ttu-id="d240f-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="d240f-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="d240f-139">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-139">Member</span></span> |
| [<span data-ttu-id="d240f-140">itemId</span><span class="sxs-lookup"><span data-stu-id="d240f-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="d240f-141">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-141">Member</span></span> |
| [<span data-ttu-id="d240f-142">itemType</span><span class="sxs-lookup"><span data-stu-id="d240f-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="d240f-143">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-143">Member</span></span> |
| [<span data-ttu-id="d240f-144">location</span><span class="sxs-lookup"><span data-stu-id="d240f-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="d240f-145">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-145">Member</span></span> |
| [<span data-ttu-id="d240f-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="d240f-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="d240f-147">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-147">Member</span></span> |
| [<span data-ttu-id="d240f-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="d240f-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="d240f-149">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-149">Member</span></span> |
| [<span data-ttu-id="d240f-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="d240f-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d240f-151">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-151">Member</span></span> |
| [<span data-ttu-id="d240f-152">organizer</span><span class="sxs-lookup"><span data-stu-id="d240f-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="d240f-153">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-153">Member</span></span> |
| [<span data-ttu-id="d240f-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="d240f-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="d240f-155">Member</span><span class="sxs-lookup"><span data-stu-id="d240f-155">Member</span></span> |
| [<span data-ttu-id="d240f-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="d240f-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d240f-157">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-157">Member</span></span> |
| [<span data-ttu-id="d240f-158">sender</span><span class="sxs-lookup"><span data-stu-id="d240f-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="d240f-159">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-159">Member</span></span> |
| [<span data-ttu-id="d240f-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="d240f-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="d240f-161">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-161">Member</span></span> |
| [<span data-ttu-id="d240f-162">start</span><span class="sxs-lookup"><span data-stu-id="d240f-162">start</span></span>](#start-datetime) | <span data-ttu-id="d240f-163">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-163">Member</span></span> |
| [<span data-ttu-id="d240f-164">subject</span><span class="sxs-lookup"><span data-stu-id="d240f-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="d240f-165">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-165">Member</span></span> |
| [<span data-ttu-id="d240f-166">to</span><span class="sxs-lookup"><span data-stu-id="d240f-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d240f-167">Membre</span><span class="sxs-lookup"><span data-stu-id="d240f-167">Member</span></span> |
| [<span data-ttu-id="d240f-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d240f-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="d240f-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-169">Method</span></span> |
| [<span data-ttu-id="d240f-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d240f-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d240f-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-171">Method</span></span> |
| [<span data-ttu-id="d240f-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d240f-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="d240f-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-173">Method</span></span> |
| [<span data-ttu-id="d240f-174">close</span><span class="sxs-lookup"><span data-stu-id="d240f-174">close</span></span>](#close) | <span data-ttu-id="d240f-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-175">Method</span></span> |
| [<span data-ttu-id="d240f-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="d240f-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="d240f-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-177">Method</span></span> |
| [<span data-ttu-id="d240f-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="d240f-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="d240f-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-179">Method</span></span> |
| [<span data-ttu-id="d240f-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="d240f-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="d240f-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-181">Method</span></span> |
| [<span data-ttu-id="d240f-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="d240f-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d240f-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-183">Method</span></span> |
| [<span data-ttu-id="d240f-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="d240f-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d240f-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-185">Method</span></span> |
| [<span data-ttu-id="d240f-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d240f-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="d240f-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-187">Method</span></span> |
| [<span data-ttu-id="d240f-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="d240f-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="d240f-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-189">Method</span></span> |
| [<span data-ttu-id="d240f-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d240f-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="d240f-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-191">Method</span></span> |
| [<span data-ttu-id="d240f-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="d240f-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="d240f-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-193">Method</span></span> |
| [<span data-ttu-id="d240f-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d240f-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="d240f-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-195">Method</span></span> |
| [<span data-ttu-id="d240f-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d240f-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="d240f-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-197">Method</span></span> |
| [<span data-ttu-id="d240f-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d240f-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="d240f-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-199">Method</span></span> |
| [<span data-ttu-id="d240f-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d240f-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d240f-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-201">Method</span></span> |
| [<span data-ttu-id="d240f-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="d240f-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="d240f-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-203">Method</span></span> |
| [<span data-ttu-id="d240f-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d240f-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="d240f-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="d240f-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="d240f-206">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-206">Example</span></span>

<span data-ttu-id="d240f-207">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="d240f-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d240f-208">Members</span><span class="sxs-lookup"><span data-stu-id="d240f-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="d240f-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="d240f-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="d240f-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d240f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-212">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="d240f-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d240f-213">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="d240f-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-214">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-214">Type</span></span>

*   <span data-ttu-id="d240f-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="d240f-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-216">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-216">Requirements</span></span>

|<span data-ttu-id="d240f-217">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-217">Requirement</span></span>|<span data-ttu-id="d240f-218">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-219">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-220">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-220">1.0</span></span>|
|[<span data-ttu-id="d240f-221">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-222">ReadItem</span></span>|
|[<span data-ttu-id="d240f-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-224">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-225">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-225">Example</span></span>

<span data-ttu-id="d240f-226">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d240f-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d240f-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-228">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="d240f-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d240f-229">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="d240f-229">Compose mode only.</span></span>

<span data-ttu-id="d240f-230">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d240f-230">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d240f-231">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="d240f-231">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d240f-232">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="d240f-232">Get 500 members maximum.</span></span>
- <span data-ttu-id="d240f-233">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="d240f-233">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-234">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-234">Type</span></span>

*   [<span data-ttu-id="d240f-235">Destinataires</span><span class="sxs-lookup"><span data-stu-id="d240f-235">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d240f-236">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-236">Requirements</span></span>

|<span data-ttu-id="d240f-237">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-237">Requirement</span></span>|<span data-ttu-id="d240f-238">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-239">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-240">1.1</span><span class="sxs-lookup"><span data-stu-id="d240f-240">1.1</span></span>|
|[<span data-ttu-id="d240f-241">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-242">ReadItem</span></span>|
|[<span data-ttu-id="d240f-243">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-244">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-244">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-245">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-245">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="d240f-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-247">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-247">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-248">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-248">Type</span></span>

*   [<span data-ttu-id="d240f-249">Body</span><span class="sxs-lookup"><span data-stu-id="d240f-249">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d240f-250">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-250">Requirements</span></span>

|<span data-ttu-id="d240f-251">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-251">Requirement</span></span>|<span data-ttu-id="d240f-252">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-253">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-254">1.1</span><span class="sxs-lookup"><span data-stu-id="d240f-254">1.1</span></span>|
|[<span data-ttu-id="d240f-255">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-256">ReadItem</span></span>|
|[<span data-ttu-id="d240f-257">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-258">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-259">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-259">Example</span></span>

<span data-ttu-id="d240f-260">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="d240f-260">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="d240f-261">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-261">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d240f-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-263">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="d240f-263">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d240f-264">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d240f-264">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-265">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-265">Read mode</span></span>

<span data-ttu-id="d240f-266">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="d240f-266">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="d240f-267">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d240f-267">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d240f-268">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="d240f-268">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-269">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-269">Compose mode</span></span>

<span data-ttu-id="d240f-270">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="d240f-270">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="d240f-271">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d240f-271">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d240f-272">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="d240f-272">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d240f-273">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="d240f-273">Get 500 members maximum.</span></span>
- <span data-ttu-id="d240f-274">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="d240f-274">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d240f-275">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-275">Type</span></span>

*   <span data-ttu-id="d240f-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-277">Requirements</span></span>

|<span data-ttu-id="d240f-278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-278">Requirement</span></span>|<span data-ttu-id="d240f-279">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-281">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-281">1.0</span></span>|
|[<span data-ttu-id="d240f-282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-283">ReadItem</span></span>|
|[<span data-ttu-id="d240f-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-285">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-285">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="d240f-286">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="d240f-286">(nullable) conversationId: String</span></span>

<span data-ttu-id="d240f-287">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="d240f-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d240f-p109">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="d240f-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d240f-p110">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="d240f-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-292">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-292">Type</span></span>

*   <span data-ttu-id="d240f-293">String</span><span class="sxs-lookup"><span data-stu-id="d240f-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-294">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-294">Requirements</span></span>

|<span data-ttu-id="d240f-295">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-295">Requirement</span></span>|<span data-ttu-id="d240f-296">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-297">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-298">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-298">1.0</span></span>|
|[<span data-ttu-id="d240f-299">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-300">ReadItem</span></span>|
|[<span data-ttu-id="d240f-301">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-302">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-302">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-303">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-303">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="d240f-304">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="d240f-304">dateTimeCreated: Date</span></span>

<span data-ttu-id="d240f-p111">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d240f-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-307">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-307">Type</span></span>

*   <span data-ttu-id="d240f-308">Date</span><span class="sxs-lookup"><span data-stu-id="d240f-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-309">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-309">Requirements</span></span>

|<span data-ttu-id="d240f-310">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-310">Requirement</span></span>|<span data-ttu-id="d240f-311">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-312">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-313">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-313">1.0</span></span>|
|[<span data-ttu-id="d240f-314">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-315">ReadItem</span></span>|
|[<span data-ttu-id="d240f-316">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-317">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-318">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-318">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="d240f-319">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="d240f-319">dateTimeModified: Date</span></span>

<span data-ttu-id="d240f-p112">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d240f-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-322">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-322">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-323">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-323">Type</span></span>

*   <span data-ttu-id="d240f-324">Date</span><span class="sxs-lookup"><span data-stu-id="d240f-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-325">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-325">Requirements</span></span>

|<span data-ttu-id="d240f-326">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-326">Requirement</span></span>|<span data-ttu-id="d240f-327">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-328">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-329">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-329">1.0</span></span>|
|[<span data-ttu-id="d240f-330">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-331">ReadItem</span></span>|
|[<span data-ttu-id="d240f-332">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-333">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-334">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-334">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="d240f-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-336">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d240f-p113">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="d240f-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-339">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-339">Read mode</span></span>

<span data-ttu-id="d240f-340">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="d240f-340">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-341">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-341">Compose mode</span></span>

<span data-ttu-id="d240f-342">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="d240f-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d240f-343">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="d240f-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d240f-344">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="d240f-344">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d240f-345">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-345">Type</span></span>

*   <span data-ttu-id="d240f-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-347">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-347">Requirements</span></span>

|<span data-ttu-id="d240f-348">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-348">Requirement</span></span>|<span data-ttu-id="d240f-349">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-350">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-351">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-351">1.0</span></span>|
|[<span data-ttu-id="d240f-352">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-353">ReadItem</span></span>|
|[<span data-ttu-id="d240f-354">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-355">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-355">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="d240f-356">from : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[from](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-356">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-357">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="d240f-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="d240f-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="d240f-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-360">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d240f-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-361">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-361">Read mode</span></span>

<span data-ttu-id="d240f-362">La `from` propriété renvoie un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="d240f-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-363">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-363">Compose mode</span></span>

<span data-ttu-id="d240f-364">La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.</span><span class="sxs-lookup"><span data-stu-id="d240f-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d240f-365">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-365">Type</span></span>

*   <span data-ttu-id="d240f-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [à partir de](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-367">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-367">Requirements</span></span>

|<span data-ttu-id="d240f-368">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d240f-369">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-370">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-370">1.0</span></span>|<span data-ttu-id="d240f-371">1.7</span><span class="sxs-lookup"><span data-stu-id="d240f-371">1.7</span></span>|
|[<span data-ttu-id="d240f-372">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-373">ReadItem</span></span>|<span data-ttu-id="d240f-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d240f-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="d240f-375">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-375">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-376">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-376">Read</span></span>|<span data-ttu-id="d240f-377">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-377">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="d240f-378">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="d240f-378">internetMessageId: String</span></span>

<span data-ttu-id="d240f-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d240f-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-381">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-381">Type</span></span>

*   <span data-ttu-id="d240f-382">String</span><span class="sxs-lookup"><span data-stu-id="d240f-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-383">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-383">Requirements</span></span>

|<span data-ttu-id="d240f-384">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-384">Requirement</span></span>|<span data-ttu-id="d240f-385">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-386">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-387">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-387">1.0</span></span>|
|[<span data-ttu-id="d240f-388">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-389">ReadItem</span></span>|
|[<span data-ttu-id="d240f-390">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-391">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-392">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-392">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="d240f-393">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="d240f-393">itemClass: String</span></span>

<span data-ttu-id="d240f-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d240f-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d240f-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="d240f-398">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-398">Type</span></span>|<span data-ttu-id="d240f-399">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-399">Description</span></span>|<span data-ttu-id="d240f-400">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="d240f-400">item class</span></span>|
|---|---|---|
|<span data-ttu-id="d240f-401">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="d240f-401">Appointment items</span></span>|<span data-ttu-id="d240f-402">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="d240f-402">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="d240f-403">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="d240f-403">Message items</span></span>|<span data-ttu-id="d240f-404">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="d240f-404">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="d240f-405">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="d240f-405">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-406">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-406">Type</span></span>

*   <span data-ttu-id="d240f-407">String</span><span class="sxs-lookup"><span data-stu-id="d240f-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-408">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-408">Requirements</span></span>

|<span data-ttu-id="d240f-409">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-409">Requirement</span></span>|<span data-ttu-id="d240f-410">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-411">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-412">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-412">1.0</span></span>|
|[<span data-ttu-id="d240f-413">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-414">ReadItem</span></span>|
|[<span data-ttu-id="d240f-415">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-416">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-417">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-417">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d240f-418">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="d240f-418">(nullable) itemId: String</span></span>

<span data-ttu-id="d240f-p118">Permet d’obtenir l’[identificateur de l’élément des services web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d240f-p118">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-421">L’identificateur renvoyé par la propriété `itemId` est identique à l’[identificateur d’élément des services web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="d240f-421">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="d240f-422">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="d240f-422">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d240f-423">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="d240f-423">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d240f-424">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="d240f-424">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d240f-p120">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-p120">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-427">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-427">Type</span></span>

*   <span data-ttu-id="d240f-428">String</span><span class="sxs-lookup"><span data-stu-id="d240f-428">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-429">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-429">Requirements</span></span>

|<span data-ttu-id="d240f-430">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-430">Requirement</span></span>|<span data-ttu-id="d240f-431">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-431">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-432">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-432">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-433">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-433">1.0</span></span>|
|[<span data-ttu-id="d240f-434">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-434">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-435">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-435">ReadItem</span></span>|
|[<span data-ttu-id="d240f-436">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-436">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-437">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-437">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-438">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-438">Example</span></span>

<span data-ttu-id="d240f-p121">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d240f-p121">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="d240f-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-442">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="d240f-442">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d240f-443">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-443">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-444">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-444">Type</span></span>

*   [<span data-ttu-id="d240f-445">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d240f-445">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d240f-446">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-446">Requirements</span></span>

|<span data-ttu-id="d240f-447">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-447">Requirement</span></span>|<span data-ttu-id="d240f-448">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-449">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-450">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-450">1.0</span></span>|
|[<span data-ttu-id="d240f-451">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-452">ReadItem</span></span>|
|[<span data-ttu-id="d240f-453">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-454">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-454">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-455">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-455">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="d240f-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-457">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-457">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-458">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-458">Read mode</span></span>

<span data-ttu-id="d240f-459">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-459">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-460">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-460">Compose mode</span></span>

<span data-ttu-id="d240f-461">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-461">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d240f-462">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-462">Type</span></span>

*   <span data-ttu-id="d240f-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-464">Requirements</span></span>

|<span data-ttu-id="d240f-465">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-465">Requirement</span></span>|<span data-ttu-id="d240f-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-468">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-468">1.0</span></span>|
|[<span data-ttu-id="d240f-469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-470">ReadItem</span></span>|
|[<span data-ttu-id="d240f-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-472">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-472">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d240f-473">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="d240f-473">normalizedSubject: String</span></span>

<span data-ttu-id="d240f-p122">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d240f-p122">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d240f-p123">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="d240f-p123">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-478">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-478">Type</span></span>

*   <span data-ttu-id="d240f-479">String</span><span class="sxs-lookup"><span data-stu-id="d240f-479">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-480">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-480">Requirements</span></span>

|<span data-ttu-id="d240f-481">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-481">Requirement</span></span>|<span data-ttu-id="d240f-482">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-482">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-483">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-484">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-484">1.0</span></span>|
|[<span data-ttu-id="d240f-485">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-485">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-486">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-486">ReadItem</span></span>|
|[<span data-ttu-id="d240f-487">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-487">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-488">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-488">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-489">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-489">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="d240f-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-491">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-491">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-492">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-492">Type</span></span>

*   [<span data-ttu-id="d240f-493">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d240f-493">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d240f-494">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-494">Requirements</span></span>

|<span data-ttu-id="d240f-495">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-495">Requirement</span></span>|<span data-ttu-id="d240f-496">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-497">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-498">1.3</span><span class="sxs-lookup"><span data-stu-id="d240f-498">1.3</span></span>|
|[<span data-ttu-id="d240f-499">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-499">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-500">ReadItem</span></span>|
|[<span data-ttu-id="d240f-501">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-501">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-502">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-502">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-503">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-503">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d240f-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-505">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="d240f-505">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d240f-506">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d240f-506">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-507">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-507">Read mode</span></span>

<span data-ttu-id="d240f-508">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-508">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="d240f-509">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d240f-509">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d240f-510">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="d240f-510">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-511">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-511">Compose mode</span></span>

<span data-ttu-id="d240f-512">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-512">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="d240f-513">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d240f-513">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d240f-514">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="d240f-514">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d240f-515">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="d240f-515">Get 500 members maximum.</span></span>
- <span data-ttu-id="d240f-516">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="d240f-516">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d240f-517">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-517">Type</span></span>

*   <span data-ttu-id="d240f-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-519">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-519">Requirements</span></span>

|<span data-ttu-id="d240f-520">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-520">Requirement</span></span>|<span data-ttu-id="d240f-521">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-522">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-523">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-523">1.0</span></span>|
|[<span data-ttu-id="d240f-524">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-525">ReadItem</span></span>|
|[<span data-ttu-id="d240f-526">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-527">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-527">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="d240f-528">Organisateur : [](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[organisateur](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d240f-528">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-529">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="d240f-529">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-530">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-530">Read mode</span></span>

<span data-ttu-id="d240f-531">La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-531">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-532">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-532">Compose mode</span></span>

<span data-ttu-id="d240f-533">La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="d240f-533">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="d240f-534">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-534">Type</span></span>

*   <span data-ttu-id="d240f-535">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organisateur](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d240f-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-536">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-536">Requirements</span></span>

|<span data-ttu-id="d240f-537">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-537">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d240f-538">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-539">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-539">1.0</span></span>|<span data-ttu-id="d240f-540">1.7</span><span class="sxs-lookup"><span data-stu-id="d240f-540">1.7</span></span>|
|[<span data-ttu-id="d240f-541">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-542">ReadItem</span></span>|<span data-ttu-id="d240f-543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d240f-543">ReadWriteItem</span></span>|
|[<span data-ttu-id="d240f-544">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-545">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-545">Read</span></span>|<span data-ttu-id="d240f-546">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-546">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="d240f-547">(Nullable) récurrence : [périodicité](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-547">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-548">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-548">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="d240f-549">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-549">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="d240f-550">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-550">Read and compose modes for appointment items.</span></span> <span data-ttu-id="d240f-551">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-551">Read mode for meeting request items.</span></span>

<span data-ttu-id="d240f-552">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="d240f-552">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="d240f-553">`null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="d240f-553">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="d240f-554">`undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-554">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="d240f-555">Remarque : les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="d240f-555">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="d240f-556">Remarque : si l’objet de périodicité `null`est, cela indique que l’objet est un rendez-vous unique ou une demande de réunion d’un seul rendez-vous et non d’une série.</span><span class="sxs-lookup"><span data-stu-id="d240f-556">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-557">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-557">Read mode</span></span>

<span data-ttu-id="d240f-558">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-558">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="d240f-559">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-559">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-560">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-560">Compose mode</span></span>

<span data-ttu-id="d240f-561">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-561">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="d240f-562">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-562">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d240f-563">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-563">Type</span></span>

* [<span data-ttu-id="d240f-564">Instances</span><span class="sxs-lookup"><span data-stu-id="d240f-564">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="d240f-565">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-565">Requirement</span></span>|<span data-ttu-id="d240f-566">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-567">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-568">1.7</span><span class="sxs-lookup"><span data-stu-id="d240f-568">1.7</span></span>|
|[<span data-ttu-id="d240f-569">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-570">ReadItem</span></span>|
|[<span data-ttu-id="d240f-571">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-572">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d240f-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-574">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="d240f-574">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d240f-575">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d240f-575">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-576">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-576">Read mode</span></span>

<span data-ttu-id="d240f-577">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-577">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="d240f-578">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d240f-578">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d240f-579">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="d240f-579">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-580">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-580">Compose mode</span></span>

<span data-ttu-id="d240f-581">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-581">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="d240f-582">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d240f-582">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d240f-583">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="d240f-583">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d240f-584">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="d240f-584">Get 500 members maximum.</span></span>
- <span data-ttu-id="d240f-585">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="d240f-585">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="d240f-586">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-586">Type</span></span>

*   <span data-ttu-id="d240f-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-588">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-588">Requirements</span></span>

|<span data-ttu-id="d240f-589">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-589">Requirement</span></span>|<span data-ttu-id="d240f-590">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-590">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-591">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-591">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-592">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-592">1.0</span></span>|
|[<span data-ttu-id="d240f-593">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-594">ReadItem</span></span>|
|[<span data-ttu-id="d240f-595">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-595">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-596">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-596">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="d240f-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-p134">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d240f-p134">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d240f-p135">Les propriétés [`from`](#from-emailaddressdetailsfrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="d240f-p135">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-602">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d240f-602">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-603">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-603">Type</span></span>

*   [<span data-ttu-id="d240f-604">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d240f-604">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d240f-605">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-605">Requirements</span></span>

|<span data-ttu-id="d240f-606">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-606">Requirement</span></span>|<span data-ttu-id="d240f-607">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-608">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-609">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-609">1.0</span></span>|
|[<span data-ttu-id="d240f-610">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-611">ReadItem</span></span>|
|[<span data-ttu-id="d240f-612">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-613">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-613">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-614">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-614">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="d240f-615">(Nullable) seriesId : chaîne</span><span class="sxs-lookup"><span data-stu-id="d240f-615">(nullable) seriesId: String</span></span>

<span data-ttu-id="d240f-616">Obtient l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="d240f-616">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="d240f-617">Dans Outlook sur le Web et les clients de bureau `seriesId` , le renvoie l’ID des services Web Exchange (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="d240f-617">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="d240f-618">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="d240f-618">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-619">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="d240f-619">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d240f-620">La `seriesId` propriété n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="d240f-620">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="d240f-621">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="d240f-621">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d240f-622">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="d240f-622">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="d240f-623">La `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-623">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="d240f-624">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-624">Type</span></span>

* <span data-ttu-id="d240f-625">String</span><span class="sxs-lookup"><span data-stu-id="d240f-625">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-626">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-626">Requirements</span></span>

|<span data-ttu-id="d240f-627">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-627">Requirement</span></span>|<span data-ttu-id="d240f-628">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-629">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-630">1.7</span><span class="sxs-lookup"><span data-stu-id="d240f-630">1.7</span></span>|
|[<span data-ttu-id="d240f-631">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-631">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-632">ReadItem</span></span>|
|[<span data-ttu-id="d240f-633">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-633">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-634">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-634">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-635">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-635">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="d240f-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-637">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-637">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d240f-p138">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="d240f-p138">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-640">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-640">Read mode</span></span>

<span data-ttu-id="d240f-641">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="d240f-641">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-642">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-642">Compose mode</span></span>

<span data-ttu-id="d240f-643">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="d240f-643">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d240f-644">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="d240f-644">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d240f-645">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="d240f-645">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d240f-646">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-646">Type</span></span>

*   <span data-ttu-id="d240f-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-648">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-648">Requirements</span></span>

|<span data-ttu-id="d240f-649">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-649">Requirement</span></span>|<span data-ttu-id="d240f-650">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-651">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-652">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-652">1.0</span></span>|
|[<span data-ttu-id="d240f-653">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-653">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-654">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-654">ReadItem</span></span>|
|[<span data-ttu-id="d240f-655">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-655">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-656">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-656">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="d240f-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-658">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d240f-659">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="d240f-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-660">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-660">Read mode</span></span>

<span data-ttu-id="d240f-p139">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="d240f-p139">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="d240f-663">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="d240f-663">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-664">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-664">Compose mode</span></span>

<span data-ttu-id="d240f-665">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="d240f-665">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="d240f-666">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-666">Type</span></span>

*   <span data-ttu-id="d240f-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-668">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-668">Requirements</span></span>

|<span data-ttu-id="d240f-669">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-669">Requirement</span></span>|<span data-ttu-id="d240f-670">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-671">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-672">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-672">1.0</span></span>|
|[<span data-ttu-id="d240f-673">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-674">ReadItem</span></span>|
|[<span data-ttu-id="d240f-675">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-676">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-676">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d240f-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d240f-678">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="d240f-678">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d240f-679">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d240f-679">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d240f-680">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-680">Read mode</span></span>

<span data-ttu-id="d240f-681">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="d240f-681">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="d240f-682">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d240f-682">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d240f-683">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="d240f-683">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="d240f-684">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d240f-684">Compose mode</span></span>

<span data-ttu-id="d240f-685">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="d240f-685">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="d240f-686">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d240f-686">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d240f-687">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="d240f-687">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d240f-688">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="d240f-688">Get 500 members maximum.</span></span>
- <span data-ttu-id="d240f-689">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="d240f-689">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d240f-690">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-690">Type</span></span>

*   <span data-ttu-id="d240f-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-692">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-692">Requirements</span></span>

|<span data-ttu-id="d240f-693">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-693">Requirement</span></span>|<span data-ttu-id="d240f-694">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-695">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-696">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-696">1.0</span></span>|
|[<span data-ttu-id="d240f-697">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-698">ReadItem</span></span>|
|[<span data-ttu-id="d240f-699">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-700">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-700">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d240f-701">Méthodes</span><span class="sxs-lookup"><span data-stu-id="d240f-701">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d240f-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d240f-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d240f-703">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="d240f-703">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d240f-704">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="d240f-704">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d240f-705">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="d240f-705">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-706">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d240f-706">Parameters</span></span>
|<span data-ttu-id="d240f-707">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-707">Name</span></span>|<span data-ttu-id="d240f-708">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-708">Type</span></span>|<span data-ttu-id="d240f-709">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-709">Attributes</span></span>|<span data-ttu-id="d240f-710">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-710">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="d240f-711">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d240f-711">String</span></span>||<span data-ttu-id="d240f-p143">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="d240f-p143">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d240f-714">String</span><span class="sxs-lookup"><span data-stu-id="d240f-714">String</span></span>||<span data-ttu-id="d240f-p144">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d240f-p144">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d240f-717">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-717">Object</span></span>|<span data-ttu-id="d240f-718">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-718">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-719">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d240f-719">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d240f-720">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-720">Object</span></span>|<span data-ttu-id="d240f-721">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-721">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-722">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-722">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="d240f-723">Boolean</span><span class="sxs-lookup"><span data-stu-id="d240f-723">Boolean</span></span>|<span data-ttu-id="d240f-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-724">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-725">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="d240f-725">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="d240f-726">fonction</span><span class="sxs-lookup"><span data-stu-id="d240f-726">function</span></span>|<span data-ttu-id="d240f-727">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-727">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-728">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-728">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d240f-729">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d240f-729">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d240f-730">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="d240f-730">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d240f-731">Erreurs</span><span class="sxs-lookup"><span data-stu-id="d240f-731">Errors</span></span>

|<span data-ttu-id="d240f-732">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="d240f-732">Error code</span></span>|<span data-ttu-id="d240f-733">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-733">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="d240f-734">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="d240f-734">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="d240f-735">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="d240f-735">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d240f-736">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="d240f-736">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-737">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-737">Requirements</span></span>

|<span data-ttu-id="d240f-738">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-738">Requirement</span></span>|<span data-ttu-id="d240f-739">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-740">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-741">1.1</span><span class="sxs-lookup"><span data-stu-id="d240f-741">1.1</span></span>|
|[<span data-ttu-id="d240f-742">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-743">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d240f-743">ReadWriteItem</span></span>|
|[<span data-ttu-id="d240f-744">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-745">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-745">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d240f-746">Exemples</span><span class="sxs-lookup"><span data-stu-id="d240f-746">Examples</span></span>

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

<span data-ttu-id="d240f-747">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="d240f-747">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d240f-748">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d240f-748">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d240f-749">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="d240f-749">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d240f-750">Actuellement, les types d’événement `Office.EventType.AppointmentTimeChanged`pris `Office.EventType.RecipientsChanged`en charge sont, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="d240f-750">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-751">Parameters</span><span class="sxs-lookup"><span data-stu-id="d240f-751">Parameters</span></span>

| <span data-ttu-id="d240f-752">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-752">Name</span></span> | <span data-ttu-id="d240f-753">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-753">Type</span></span> | <span data-ttu-id="d240f-754">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-754">Attributes</span></span> | <span data-ttu-id="d240f-755">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-755">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d240f-756">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d240f-756">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d240f-757">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="d240f-757">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d240f-758">Fonction</span><span class="sxs-lookup"><span data-stu-id="d240f-758">Function</span></span> || <span data-ttu-id="d240f-p145">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="d240f-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d240f-762">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-762">Object</span></span> | <span data-ttu-id="d240f-763">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-763">&lt;optional&gt;</span></span> | <span data-ttu-id="d240f-764">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d240f-764">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d240f-765">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-765">Object</span></span> | <span data-ttu-id="d240f-766">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-766">&lt;optional&gt;</span></span> | <span data-ttu-id="d240f-767">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-767">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d240f-768">fonction</span><span class="sxs-lookup"><span data-stu-id="d240f-768">function</span></span>| <span data-ttu-id="d240f-769">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-769">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-770">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-770">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-771">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-771">Requirements</span></span>

|<span data-ttu-id="d240f-772">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-772">Requirement</span></span>| <span data-ttu-id="d240f-773">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-773">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-774">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-774">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d240f-775">1.7</span><span class="sxs-lookup"><span data-stu-id="d240f-775">1.7</span></span> |
|[<span data-ttu-id="d240f-776">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-776">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d240f-777">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-777">ReadItem</span></span> |
|[<span data-ttu-id="d240f-778">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-778">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d240f-779">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-779">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="d240f-780">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-780">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d240f-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d240f-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d240f-782">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-782">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d240f-p146">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-p146">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d240f-786">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="d240f-786">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d240f-787">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="d240f-787">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-788">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d240f-788">Parameters</span></span>

|<span data-ttu-id="d240f-789">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-789">Name</span></span>|<span data-ttu-id="d240f-790">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-790">Type</span></span>|<span data-ttu-id="d240f-791">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-791">Attributes</span></span>|<span data-ttu-id="d240f-792">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-792">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="d240f-793">String</span><span class="sxs-lookup"><span data-stu-id="d240f-793">String</span></span>||<span data-ttu-id="d240f-p147">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="d240f-p147">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d240f-796">String</span><span class="sxs-lookup"><span data-stu-id="d240f-796">String</span></span>||<span data-ttu-id="d240f-797">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="d240f-797">The subject of the item to be attached.</span></span> <span data-ttu-id="d240f-798">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d240f-798">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d240f-799">Object</span><span class="sxs-lookup"><span data-stu-id="d240f-799">Object</span></span>|<span data-ttu-id="d240f-800">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-800">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-801">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d240f-801">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d240f-802">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-802">Object</span></span>|<span data-ttu-id="d240f-803">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-803">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-804">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-804">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d240f-805">fonction</span><span class="sxs-lookup"><span data-stu-id="d240f-805">function</span></span>|<span data-ttu-id="d240f-806">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-806">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-807">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-807">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d240f-808">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d240f-808">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d240f-809">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="d240f-809">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d240f-810">Erreurs</span><span class="sxs-lookup"><span data-stu-id="d240f-810">Errors</span></span>

|<span data-ttu-id="d240f-811">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="d240f-811">Error code</span></span>|<span data-ttu-id="d240f-812">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-812">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d240f-813">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="d240f-813">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-814">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-814">Requirements</span></span>

|<span data-ttu-id="d240f-815">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-815">Requirement</span></span>|<span data-ttu-id="d240f-816">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-817">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-818">1.1</span><span class="sxs-lookup"><span data-stu-id="d240f-818">1.1</span></span>|
|[<span data-ttu-id="d240f-819">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-820">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d240f-820">ReadWriteItem</span></span>|
|[<span data-ttu-id="d240f-821">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-822">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-822">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-823">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-823">Example</span></span>

<span data-ttu-id="d240f-824">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="d240f-824">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="d240f-825">close()</span><span class="sxs-lookup"><span data-stu-id="d240f-825">close()</span></span>

<span data-ttu-id="d240f-826">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="d240f-826">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d240f-p149">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="d240f-p149">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-829">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-829">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d240f-830">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="d240f-830">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-831">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-831">Requirements</span></span>

|<span data-ttu-id="d240f-832">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-832">Requirement</span></span>|<span data-ttu-id="d240f-833">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-834">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-835">1.3</span><span class="sxs-lookup"><span data-stu-id="d240f-835">1.3</span></span>|
|[<span data-ttu-id="d240f-836">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-837">Restreinte</span><span class="sxs-lookup"><span data-stu-id="d240f-837">Restricted</span></span>|
|[<span data-ttu-id="d240f-838">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-839">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-839">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="d240f-840">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d240f-840">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="d240f-841">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d240f-841">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-842">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-842">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d240f-843">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="d240f-843">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d240f-844">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="d240f-844">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d240f-p150">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="d240f-p150">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-848">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d240f-848">Parameters</span></span>

|<span data-ttu-id="d240f-849">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-849">Name</span></span>|<span data-ttu-id="d240f-850">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-850">Type</span></span>|<span data-ttu-id="d240f-851">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-851">Attributes</span></span>|<span data-ttu-id="d240f-852">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-852">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d240f-853">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d240f-853">String &#124; Object</span></span>||<span data-ttu-id="d240f-p151">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d240f-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d240f-856">**OU**</span><span class="sxs-lookup"><span data-stu-id="d240f-856">**OR**</span></span><br/><span data-ttu-id="d240f-p152">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="d240f-p152">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d240f-859">String</span><span class="sxs-lookup"><span data-stu-id="d240f-859">String</span></span>|<span data-ttu-id="d240f-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-860">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-p153">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d240f-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d240f-863">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-863">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d240f-864">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-864">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-865">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-865">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d240f-866">String</span><span class="sxs-lookup"><span data-stu-id="d240f-866">String</span></span>||<span data-ttu-id="d240f-p154">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-p154">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d240f-869">String</span><span class="sxs-lookup"><span data-stu-id="d240f-869">String</span></span>||<span data-ttu-id="d240f-870">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d240f-870">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d240f-871">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d240f-871">String</span></span>||<span data-ttu-id="d240f-p155">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="d240f-p155">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d240f-874">Booléen</span><span class="sxs-lookup"><span data-stu-id="d240f-874">Boolean</span></span>||<span data-ttu-id="d240f-p156">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="d240f-p156">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d240f-877">String</span><span class="sxs-lookup"><span data-stu-id="d240f-877">String</span></span>||<span data-ttu-id="d240f-p157">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="d240f-p157">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d240f-881">function</span><span class="sxs-lookup"><span data-stu-id="d240f-881">function</span></span>|<span data-ttu-id="d240f-882">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-882">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-883">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-884">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-884">Requirements</span></span>

|<span data-ttu-id="d240f-885">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-885">Requirement</span></span>|<span data-ttu-id="d240f-886">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-887">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-888">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-888">1.0</span></span>|
|[<span data-ttu-id="d240f-889">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-890">ReadItem</span></span>|
|[<span data-ttu-id="d240f-891">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-892">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-892">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d240f-893">Exemples</span><span class="sxs-lookup"><span data-stu-id="d240f-893">Examples</span></span>

<span data-ttu-id="d240f-894">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="d240f-894">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d240f-895">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="d240f-895">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d240f-896">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="d240f-896">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d240f-897">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="d240f-897">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d240f-898">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-898">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d240f-899">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-899">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="d240f-900">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d240f-900">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="d240f-901">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d240f-901">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-902">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-902">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d240f-903">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="d240f-903">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d240f-904">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="d240f-904">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d240f-p158">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="d240f-p158">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-908">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d240f-908">Parameters</span></span>

|<span data-ttu-id="d240f-909">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-909">Name</span></span>|<span data-ttu-id="d240f-910">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-910">Type</span></span>|<span data-ttu-id="d240f-911">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-911">Attributes</span></span>|<span data-ttu-id="d240f-912">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d240f-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d240f-913">String &#124; Object</span></span>||<span data-ttu-id="d240f-p159">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d240f-p159">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d240f-916">**OU**</span><span class="sxs-lookup"><span data-stu-id="d240f-916">**OR**</span></span><br/><span data-ttu-id="d240f-p160">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="d240f-p160">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d240f-919">String</span><span class="sxs-lookup"><span data-stu-id="d240f-919">String</span></span>|<span data-ttu-id="d240f-920">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-920">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-p161">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d240f-p161">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d240f-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d240f-924">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-924">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-925">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d240f-926">String</span><span class="sxs-lookup"><span data-stu-id="d240f-926">String</span></span>||<span data-ttu-id="d240f-p162">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-p162">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d240f-929">String</span><span class="sxs-lookup"><span data-stu-id="d240f-929">String</span></span>||<span data-ttu-id="d240f-930">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d240f-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d240f-931">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d240f-931">String</span></span>||<span data-ttu-id="d240f-p163">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="d240f-p163">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d240f-934">Booléen</span><span class="sxs-lookup"><span data-stu-id="d240f-934">Boolean</span></span>||<span data-ttu-id="d240f-p164">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="d240f-p164">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d240f-937">String</span><span class="sxs-lookup"><span data-stu-id="d240f-937">String</span></span>||<span data-ttu-id="d240f-p165">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="d240f-p165">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d240f-941">function</span><span class="sxs-lookup"><span data-stu-id="d240f-941">function</span></span>|<span data-ttu-id="d240f-942">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-942">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-943">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-944">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-944">Requirements</span></span>

|<span data-ttu-id="d240f-945">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-945">Requirement</span></span>|<span data-ttu-id="d240f-946">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-947">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-948">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-948">1.0</span></span>|
|[<span data-ttu-id="d240f-949">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-950">ReadItem</span></span>|
|[<span data-ttu-id="d240f-951">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-952">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d240f-953">Exemples</span><span class="sxs-lookup"><span data-stu-id="d240f-953">Examples</span></span>

<span data-ttu-id="d240f-954">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="d240f-954">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d240f-955">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="d240f-955">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d240f-956">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="d240f-956">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d240f-957">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="d240f-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d240f-958">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d240f-959">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="d240f-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="d240f-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="d240f-961">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d240f-961">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-962">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-962">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-963">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-963">Requirements</span></span>

|<span data-ttu-id="d240f-964">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-964">Requirement</span></span>|<span data-ttu-id="d240f-965">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-966">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-967">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-967">1.0</span></span>|
|[<span data-ttu-id="d240f-968">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-968">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-969">ReadItem</span></span>|
|[<span data-ttu-id="d240f-970">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-970">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-971">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-971">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d240f-972">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d240f-972">Returns:</span></span>

<span data-ttu-id="d240f-973">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-973">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="d240f-974">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-974">Example</span></span>

<span data-ttu-id="d240f-975">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d240f-975">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="d240f-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="d240f-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="d240f-977">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d240f-977">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-978">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-978">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-979">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d240f-979">Parameters</span></span>

|<span data-ttu-id="d240f-980">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-980">Name</span></span>|<span data-ttu-id="d240f-981">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-981">Type</span></span>|<span data-ttu-id="d240f-982">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-982">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="d240f-983">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d240f-983">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="d240f-984">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="d240f-984">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-985">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-985">Requirements</span></span>

|<span data-ttu-id="d240f-986">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-986">Requirement</span></span>|<span data-ttu-id="d240f-987">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-987">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-988">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-988">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-989">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-989">1.0</span></span>|
|[<span data-ttu-id="d240f-990">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-990">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-991">Restreinte</span><span class="sxs-lookup"><span data-stu-id="d240f-991">Restricted</span></span>|
|[<span data-ttu-id="d240f-992">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-992">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-993">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-993">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d240f-994">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d240f-994">Returns:</span></span>

<span data-ttu-id="d240f-995">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="d240f-995">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d240f-996">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="d240f-996">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d240f-997">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="d240f-997">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d240f-998">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="d240f-998">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="d240f-999">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="d240f-999">Value of `entityType`</span></span>|<span data-ttu-id="d240f-1000">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="d240f-1000">Type of objects in returned array</span></span>|<span data-ttu-id="d240f-1001">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="d240f-1001">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="d240f-1002">String</span><span class="sxs-lookup"><span data-stu-id="d240f-1002">String</span></span>|<span data-ttu-id="d240f-1003">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="d240f-1003">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="d240f-1004">Contact</span><span class="sxs-lookup"><span data-stu-id="d240f-1004">Contact</span></span>|<span data-ttu-id="d240f-1005">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d240f-1005">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="d240f-1006">String</span><span class="sxs-lookup"><span data-stu-id="d240f-1006">String</span></span>|<span data-ttu-id="d240f-1007">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d240f-1007">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="d240f-1008">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d240f-1008">MeetingSuggestion</span></span>|<span data-ttu-id="d240f-1009">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d240f-1009">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="d240f-1010">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d240f-1010">PhoneNumber</span></span>|<span data-ttu-id="d240f-1011">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="d240f-1011">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="d240f-1012">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d240f-1012">TaskSuggestion</span></span>|<span data-ttu-id="d240f-1013">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d240f-1013">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="d240f-1014">String</span><span class="sxs-lookup"><span data-stu-id="d240f-1014">String</span></span>|<span data-ttu-id="d240f-1015">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="d240f-1015">**Restricted**</span></span>|

<span data-ttu-id="d240f-1016">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="d240f-1016">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="d240f-1017">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1017">Example</span></span>

<span data-ttu-id="d240f-1018">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d240f-1018">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="d240f-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="d240f-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="d240f-1020">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="d240f-1020">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-1021">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-1021">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d240f-1022">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="d240f-1022">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-1023">Parameters</span><span class="sxs-lookup"><span data-stu-id="d240f-1023">Parameters</span></span>

|<span data-ttu-id="d240f-1024">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-1024">Name</span></span>|<span data-ttu-id="d240f-1025">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-1025">Type</span></span>|<span data-ttu-id="d240f-1026">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-1026">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d240f-1027">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d240f-1027">String</span></span>|<span data-ttu-id="d240f-1028">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="d240f-1028">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-1029">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1029">Requirements</span></span>

|<span data-ttu-id="d240f-1030">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1030">Requirement</span></span>|<span data-ttu-id="d240f-1031">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1031">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1032">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1032">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1033">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-1033">1.0</span></span>|
|[<span data-ttu-id="d240f-1034">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1034">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1035">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1035">ReadItem</span></span>|
|[<span data-ttu-id="d240f-1036">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1036">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1037">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-1037">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d240f-1038">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d240f-1038">Returns:</span></span>

<span data-ttu-id="d240f-p167">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="d240f-p167">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d240f-1041">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="d240f-1041">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="d240f-1042">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d240f-1042">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d240f-1043">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="d240f-1043">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-1044">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d240f-p168">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="d240f-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d240f-1048">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="d240f-1048">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d240f-1049">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="d240f-1049">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d240f-p169">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-1053">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1053">Requirements</span></span>

|<span data-ttu-id="d240f-1054">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1054">Requirement</span></span>|<span data-ttu-id="d240f-1055">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1056">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1057">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-1057">1.0</span></span>|
|[<span data-ttu-id="d240f-1058">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1059">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1059">ReadItem</span></span>|
|[<span data-ttu-id="d240f-1060">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1061">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-1061">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d240f-1062">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d240f-1062">Returns:</span></span>

<span data-ttu-id="d240f-p170">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="d240f-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="d240f-1065">Type : Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1065">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="d240f-1066">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1066">Example</span></span>

<span data-ttu-id="d240f-1067">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="d240f-1067">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d240f-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="d240f-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d240f-1069">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="d240f-1069">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-1070">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-1070">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d240f-1071">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="d240f-1071">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d240f-p171">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="d240f-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-1074">Parameters</span><span class="sxs-lookup"><span data-stu-id="d240f-1074">Parameters</span></span>

|<span data-ttu-id="d240f-1075">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-1075">Name</span></span>|<span data-ttu-id="d240f-1076">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-1076">Type</span></span>|<span data-ttu-id="d240f-1077">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-1077">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d240f-1078">String</span><span class="sxs-lookup"><span data-stu-id="d240f-1078">String</span></span>|<span data-ttu-id="d240f-1079">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="d240f-1079">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-1080">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1080">Requirements</span></span>

|<span data-ttu-id="d240f-1081">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1081">Requirement</span></span>|<span data-ttu-id="d240f-1082">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1083">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1084">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-1084">1.0</span></span>|
|[<span data-ttu-id="d240f-1085">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1085">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1086">ReadItem</span></span>|
|[<span data-ttu-id="d240f-1087">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1087">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1088">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-1088">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d240f-1089">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d240f-1089">Returns:</span></span>

<span data-ttu-id="d240f-1090">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="d240f-1090">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="d240f-1091">Type : Array.< String ></span><span class="sxs-lookup"><span data-stu-id="d240f-1091">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="d240f-1092">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1092">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d240f-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d240f-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d240f-1094">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="d240f-1094">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d240f-p172">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie une chaîne vide pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="d240f-p172">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-1097">Parameters</span><span class="sxs-lookup"><span data-stu-id="d240f-1097">Parameters</span></span>

|<span data-ttu-id="d240f-1098">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-1098">Name</span></span>|<span data-ttu-id="d240f-1099">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-1099">Type</span></span>|<span data-ttu-id="d240f-1100">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-1100">Attributes</span></span>|<span data-ttu-id="d240f-1101">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-1101">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="d240f-1102">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d240f-1102">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d240f-p173">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="d240f-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="d240f-1106">Object</span><span class="sxs-lookup"><span data-stu-id="d240f-1106">Object</span></span>|<span data-ttu-id="d240f-1107">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1107">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1108">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d240f-1108">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d240f-1109">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1109">Object</span></span>|<span data-ttu-id="d240f-1110">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1110">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1111">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-1111">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d240f-1112">fonction</span><span class="sxs-lookup"><span data-stu-id="d240f-1112">function</span></span>||<span data-ttu-id="d240f-1113">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-1113">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d240f-1114">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="d240f-1114">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d240f-1115">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="d240f-1115">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-1116">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1116">Requirements</span></span>

|<span data-ttu-id="d240f-1117">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1117">Requirement</span></span>|<span data-ttu-id="d240f-1118">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1118">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1119">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1119">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1120">1.2</span><span class="sxs-lookup"><span data-stu-id="d240f-1120">1.2</span></span>|
|[<span data-ttu-id="d240f-1121">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1121">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1122">ReadItem</span></span>|
|[<span data-ttu-id="d240f-1123">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1123">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1124">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-1124">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d240f-1125">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d240f-1125">Returns:</span></span>

<span data-ttu-id="d240f-1126">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="d240f-1126">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="d240f-1127">Type : String</span><span class="sxs-lookup"><span data-stu-id="d240f-1127">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d240f-1128">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1128">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="d240f-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="d240f-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="d240f-1130">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d240f-1130">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="d240f-1131">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="d240f-1131">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-1132">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-1132">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-1133">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1133">Requirements</span></span>

|<span data-ttu-id="d240f-1134">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1134">Requirement</span></span>|<span data-ttu-id="d240f-1135">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1136">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1137">1.6</span><span class="sxs-lookup"><span data-stu-id="d240f-1137">1.6</span></span>|
|[<span data-ttu-id="d240f-1138">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1139">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1139">ReadItem</span></span>|
|[<span data-ttu-id="d240f-1140">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1141">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d240f-1142">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d240f-1142">Returns:</span></span>

<span data-ttu-id="d240f-1143">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d240f-1143">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="d240f-1144">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1144">Example</span></span>

<span data-ttu-id="d240f-1145">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d240f-1145">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="d240f-1146">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d240f-1146">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="d240f-p176">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="d240f-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-1149">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d240f-1149">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d240f-p177">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="d240f-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d240f-1153">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="d240f-1153">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d240f-1154">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="d240f-1154">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d240f-p178">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d240f-1158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1158">Requirements</span></span>

|<span data-ttu-id="d240f-1159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1159">Requirement</span></span>|<span data-ttu-id="d240f-1160">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1160">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1162">1.6</span><span class="sxs-lookup"><span data-stu-id="d240f-1162">1.6</span></span>|
|[<span data-ttu-id="d240f-1163">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1164">ReadItem</span></span>|
|[<span data-ttu-id="d240f-1165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1166">Lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-1166">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d240f-1167">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d240f-1167">Returns:</span></span>

<span data-ttu-id="d240f-p179">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="d240f-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="d240f-1170">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1170">Example</span></span>

<span data-ttu-id="d240f-1171">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="d240f-1171">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d240f-1172">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d240f-1172">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d240f-1173">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d240f-1173">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d240f-p180">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="d240f-p180">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-1177">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d240f-1177">Parameters</span></span>

|<span data-ttu-id="d240f-1178">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-1178">Name</span></span>|<span data-ttu-id="d240f-1179">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-1179">Type</span></span>|<span data-ttu-id="d240f-1180">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-1180">Attributes</span></span>|<span data-ttu-id="d240f-1181">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-1181">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="d240f-1182">function</span><span class="sxs-lookup"><span data-stu-id="d240f-1182">function</span></span>||<span data-ttu-id="d240f-1183">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-1183">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d240f-1184">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d240f-1184">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d240f-1185">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="d240f-1185">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="d240f-1186">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1186">Object</span></span>|<span data-ttu-id="d240f-1187">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1188">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-1188">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d240f-1189">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-1189">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-1190">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1190">Requirements</span></span>

|<span data-ttu-id="d240f-1191">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1191">Requirement</span></span>|<span data-ttu-id="d240f-1192">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1193">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1194">1.0</span><span class="sxs-lookup"><span data-stu-id="d240f-1194">1.0</span></span>|
|[<span data-ttu-id="d240f-1195">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1196">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1196">ReadItem</span></span>|
|[<span data-ttu-id="d240f-1197">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1198">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-1198">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-1199">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1199">Example</span></span>

<span data-ttu-id="d240f-p183">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="d240f-p183">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d240f-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d240f-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d240f-1204">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d240f-1204">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d240f-1205">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-1205">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d240f-1206">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="d240f-1206">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="d240f-1207">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="d240f-1207">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d240f-1208">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="d240f-1208">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-1209">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d240f-1209">Parameters</span></span>

|<span data-ttu-id="d240f-1210">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-1210">Name</span></span>|<span data-ttu-id="d240f-1211">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-1211">Type</span></span>|<span data-ttu-id="d240f-1212">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-1212">Attributes</span></span>|<span data-ttu-id="d240f-1213">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-1213">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="d240f-1214">String</span><span class="sxs-lookup"><span data-stu-id="d240f-1214">String</span></span>||<span data-ttu-id="d240f-1215">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="d240f-1215">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="d240f-1216">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1216">Object</span></span>|<span data-ttu-id="d240f-1217">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1217">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1218">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d240f-1218">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d240f-1219">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1219">Object</span></span>|<span data-ttu-id="d240f-1220">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1220">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1221">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-1221">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d240f-1222">fonction</span><span class="sxs-lookup"><span data-stu-id="d240f-1222">function</span></span>|<span data-ttu-id="d240f-1223">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1223">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1224">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-1224">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d240f-1225">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="d240f-1225">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d240f-1226">Erreurs</span><span class="sxs-lookup"><span data-stu-id="d240f-1226">Errors</span></span>

|<span data-ttu-id="d240f-1227">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="d240f-1227">Error code</span></span>|<span data-ttu-id="d240f-1228">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-1228">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="d240f-1229">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="d240f-1229">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-1230">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1230">Requirements</span></span>

|<span data-ttu-id="d240f-1231">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1231">Requirement</span></span>|<span data-ttu-id="d240f-1232">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1233">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1234">1.1</span><span class="sxs-lookup"><span data-stu-id="d240f-1234">1.1</span></span>|
|[<span data-ttu-id="d240f-1235">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1236">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1236">ReadWriteItem</span></span>|
|[<span data-ttu-id="d240f-1237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1238">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-1238">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-1239">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1239">Example</span></span>

<span data-ttu-id="d240f-1240">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="d240f-1240">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d240f-1241">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d240f-1241">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d240f-1242">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="d240f-1242">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d240f-1243">Actuellement, les types d’événement `Office.EventType.AppointmentTimeChanged`pris `Office.EventType.RecipientsChanged`en charge sont, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="d240f-1243">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-1244">Parameters</span><span class="sxs-lookup"><span data-stu-id="d240f-1244">Parameters</span></span>

| <span data-ttu-id="d240f-1245">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-1245">Name</span></span> | <span data-ttu-id="d240f-1246">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-1246">Type</span></span> | <span data-ttu-id="d240f-1247">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-1247">Attributes</span></span> | <span data-ttu-id="d240f-1248">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-1248">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d240f-1249">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d240f-1249">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d240f-1250">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="d240f-1250">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="d240f-1251">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1251">Object</span></span> | <span data-ttu-id="d240f-1252">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1252">&lt;optional&gt;</span></span> | <span data-ttu-id="d240f-1253">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d240f-1253">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d240f-1254">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1254">Object</span></span> | <span data-ttu-id="d240f-1255">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1255">&lt;optional&gt;</span></span> | <span data-ttu-id="d240f-1256">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-1256">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d240f-1257">fonction</span><span class="sxs-lookup"><span data-stu-id="d240f-1257">function</span></span>| <span data-ttu-id="d240f-1258">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1258">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1259">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-1259">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-1260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1260">Requirements</span></span>

|<span data-ttu-id="d240f-1261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1261">Requirement</span></span>| <span data-ttu-id="d240f-1262">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1262">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d240f-1264">1.7</span><span class="sxs-lookup"><span data-stu-id="d240f-1264">1.7</span></span> |
|[<span data-ttu-id="d240f-1265">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d240f-1266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1266">ReadItem</span></span> |
|[<span data-ttu-id="d240f-1267">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d240f-1268">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d240f-1268">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="d240f-1269">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1269">Example</span></span>

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

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="d240f-1270">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d240f-1270">saveAsync([options], callback)</span></span>

<span data-ttu-id="d240f-1271">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d240f-1271">Asynchronously saves an item.</span></span>

<span data-ttu-id="d240f-1272">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-1272">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="d240f-1273">Dans Outlook sur le web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="d240f-1273">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="d240f-1274">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="d240f-1274">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-1275">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="d240f-1275">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d240f-1276">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="d240f-1276">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d240f-p187">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="d240f-p187">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d240f-1280">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="d240f-1280">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d240f-1281">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="d240f-1281">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="d240f-1282">La méthode `saveAsync` échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="d240f-1282">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="d240f-1283">Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="d240f-1283">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="d240f-1284">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="d240f-1284">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-1285">Parameters</span><span class="sxs-lookup"><span data-stu-id="d240f-1285">Parameters</span></span>

|<span data-ttu-id="d240f-1286">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-1286">Name</span></span>|<span data-ttu-id="d240f-1287">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-1287">Type</span></span>|<span data-ttu-id="d240f-1288">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-1288">Attributes</span></span>|<span data-ttu-id="d240f-1289">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-1289">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d240f-1290">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1290">Object</span></span>|<span data-ttu-id="d240f-1291">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1291">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1292">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d240f-1292">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d240f-1293">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1293">Object</span></span>|<span data-ttu-id="d240f-1294">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1294">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1295">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-1295">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d240f-1296">fonction</span><span class="sxs-lookup"><span data-stu-id="d240f-1296">function</span></span>||<span data-ttu-id="d240f-1297">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-1297">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d240f-1298">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d240f-1298">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-1299">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1299">Requirements</span></span>

|<span data-ttu-id="d240f-1300">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1300">Requirement</span></span>|<span data-ttu-id="d240f-1301">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1302">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1303">1.3</span><span class="sxs-lookup"><span data-stu-id="d240f-1303">1.3</span></span>|
|[<span data-ttu-id="d240f-1304">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1305">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1305">ReadWriteItem</span></span>|
|[<span data-ttu-id="d240f-1306">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1307">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-1307">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d240f-1308">範例</span><span class="sxs-lookup"><span data-stu-id="d240f-1308">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="d240f-p189">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="d240f-p189">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d240f-1311">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d240f-1311">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d240f-1312">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d240f-1312">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d240f-p190">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="d240f-p190">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d240f-1316">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d240f-1316">Parameters</span></span>

|<span data-ttu-id="d240f-1317">Nom</span><span class="sxs-lookup"><span data-stu-id="d240f-1317">Name</span></span>|<span data-ttu-id="d240f-1318">Type</span><span class="sxs-lookup"><span data-stu-id="d240f-1318">Type</span></span>|<span data-ttu-id="d240f-1319">Attributs</span><span class="sxs-lookup"><span data-stu-id="d240f-1319">Attributes</span></span>|<span data-ttu-id="d240f-1320">Description</span><span class="sxs-lookup"><span data-stu-id="d240f-1320">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="d240f-1321">String</span><span class="sxs-lookup"><span data-stu-id="d240f-1321">String</span></span>||<span data-ttu-id="d240f-p191">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="d240f-p191">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="d240f-1325">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1325">Object</span></span>|<span data-ttu-id="d240f-1326">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1326">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1327">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d240f-1327">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d240f-1328">Objet</span><span class="sxs-lookup"><span data-stu-id="d240f-1328">Object</span></span>|<span data-ttu-id="d240f-1329">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1329">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1330">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d240f-1330">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="d240f-1331">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d240f-1331">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="d240f-1332">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d240f-1332">&lt;optional&gt;</span></span>|<span data-ttu-id="d240f-1333">Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau.</span><span class="sxs-lookup"><span data-stu-id="d240f-1333">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="d240f-1334">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="d240f-1334">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d240f-1335">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="d240f-1335">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="d240f-1336">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="d240f-1336">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d240f-1337">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="d240f-1337">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="d240f-1338">fonction</span><span class="sxs-lookup"><span data-stu-id="d240f-1338">function</span></span>||<span data-ttu-id="d240f-1339">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d240f-1339">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d240f-1340">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d240f-1340">Requirements</span></span>

|<span data-ttu-id="d240f-1341">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d240f-1341">Requirement</span></span>|<span data-ttu-id="d240f-1342">Valeur</span><span class="sxs-lookup"><span data-stu-id="d240f-1342">Value</span></span>|
|---|---|
|[<span data-ttu-id="d240f-1343">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d240f-1343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d240f-1344">1.2</span><span class="sxs-lookup"><span data-stu-id="d240f-1344">1.2</span></span>|
|[<span data-ttu-id="d240f-1345">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d240f-1345">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d240f-1346">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d240f-1346">ReadWriteItem</span></span>|
|[<span data-ttu-id="d240f-1347">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d240f-1347">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d240f-1348">Composition</span><span class="sxs-lookup"><span data-stu-id="d240f-1348">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d240f-1349">Exemple</span><span class="sxs-lookup"><span data-stu-id="d240f-1349">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
