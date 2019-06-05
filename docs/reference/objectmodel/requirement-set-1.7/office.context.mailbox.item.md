---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,7
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: fd618b766a519c522f323e0a9d43105b3258c421
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706312"
---
# <a name="item"></a><span data-ttu-id="21f7f-102">élément</span><span class="sxs-lookup"><span data-stu-id="21f7f-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="21f7f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="21f7f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="21f7f-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="21f7f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-106">Requirements</span></span>

|<span data-ttu-id="21f7f-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-107">Requirement</span></span>|<span data-ttu-id="21f7f-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-110">1.0</span></span>|
|[<span data-ttu-id="21f7f-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="21f7f-112">Restricted</span></span>|
|[<span data-ttu-id="21f7f-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="21f7f-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="21f7f-115">Members and methods</span></span>

| <span data-ttu-id="21f7f-116">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-116">Member</span></span> | <span data-ttu-id="21f7f-117">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="21f7f-118">attachments</span><span class="sxs-lookup"><span data-stu-id="21f7f-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="21f7f-119">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-119">Member</span></span> |
| [<span data-ttu-id="21f7f-120">bcc</span><span class="sxs-lookup"><span data-stu-id="21f7f-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="21f7f-121">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-121">Member</span></span> |
| [<span data-ttu-id="21f7f-122">body</span><span class="sxs-lookup"><span data-stu-id="21f7f-122">body</span></span>](#body-body) | <span data-ttu-id="21f7f-123">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-123">Member</span></span> |
| [<span data-ttu-id="21f7f-124">cc</span><span class="sxs-lookup"><span data-stu-id="21f7f-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="21f7f-125">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-125">Member</span></span> |
| [<span data-ttu-id="21f7f-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="21f7f-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="21f7f-127">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-127">Member</span></span> |
| [<span data-ttu-id="21f7f-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="21f7f-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="21f7f-129">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-129">Member</span></span> |
| [<span data-ttu-id="21f7f-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="21f7f-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="21f7f-131">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-131">Member</span></span> |
| [<span data-ttu-id="21f7f-132">end</span><span class="sxs-lookup"><span data-stu-id="21f7f-132">end</span></span>](#end-datetime) | <span data-ttu-id="21f7f-133">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-133">Member</span></span> |
| [<span data-ttu-id="21f7f-134">from</span><span class="sxs-lookup"><span data-stu-id="21f7f-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="21f7f-135">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-135">Member</span></span> |
| [<span data-ttu-id="21f7f-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="21f7f-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="21f7f-137">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-137">Member</span></span> |
| [<span data-ttu-id="21f7f-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="21f7f-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="21f7f-139">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-139">Member</span></span> |
| [<span data-ttu-id="21f7f-140">itemId</span><span class="sxs-lookup"><span data-stu-id="21f7f-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="21f7f-141">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-141">Member</span></span> |
| [<span data-ttu-id="21f7f-142">itemType</span><span class="sxs-lookup"><span data-stu-id="21f7f-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="21f7f-143">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-143">Member</span></span> |
| [<span data-ttu-id="21f7f-144">location</span><span class="sxs-lookup"><span data-stu-id="21f7f-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="21f7f-145">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-145">Member</span></span> |
| [<span data-ttu-id="21f7f-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="21f7f-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="21f7f-147">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-147">Member</span></span> |
| [<span data-ttu-id="21f7f-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="21f7f-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="21f7f-149">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-149">Member</span></span> |
| [<span data-ttu-id="21f7f-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="21f7f-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="21f7f-151">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-151">Member</span></span> |
| [<span data-ttu-id="21f7f-152">organizer</span><span class="sxs-lookup"><span data-stu-id="21f7f-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="21f7f-153">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-153">Member</span></span> |
| [<span data-ttu-id="21f7f-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="21f7f-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="21f7f-155">Member</span><span class="sxs-lookup"><span data-stu-id="21f7f-155">Member</span></span> |
| [<span data-ttu-id="21f7f-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="21f7f-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="21f7f-157">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-157">Member</span></span> |
| [<span data-ttu-id="21f7f-158">sender</span><span class="sxs-lookup"><span data-stu-id="21f7f-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="21f7f-159">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-159">Member</span></span> |
| [<span data-ttu-id="21f7f-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="21f7f-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="21f7f-161">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-161">Member</span></span> |
| [<span data-ttu-id="21f7f-162">start</span><span class="sxs-lookup"><span data-stu-id="21f7f-162">start</span></span>](#start-datetime) | <span data-ttu-id="21f7f-163">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-163">Member</span></span> |
| [<span data-ttu-id="21f7f-164">subject</span><span class="sxs-lookup"><span data-stu-id="21f7f-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="21f7f-165">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-165">Member</span></span> |
| [<span data-ttu-id="21f7f-166">to</span><span class="sxs-lookup"><span data-stu-id="21f7f-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="21f7f-167">Membre</span><span class="sxs-lookup"><span data-stu-id="21f7f-167">Member</span></span> |
| [<span data-ttu-id="21f7f-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="21f7f-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="21f7f-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-169">Method</span></span> |
| [<span data-ttu-id="21f7f-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="21f7f-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="21f7f-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-171">Method</span></span> |
| [<span data-ttu-id="21f7f-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="21f7f-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="21f7f-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-173">Method</span></span> |
| [<span data-ttu-id="21f7f-174">close</span><span class="sxs-lookup"><span data-stu-id="21f7f-174">close</span></span>](#close) | <span data-ttu-id="21f7f-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-175">Method</span></span> |
| [<span data-ttu-id="21f7f-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="21f7f-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="21f7f-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-177">Method</span></span> |
| [<span data-ttu-id="21f7f-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="21f7f-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="21f7f-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-179">Method</span></span> |
| [<span data-ttu-id="21f7f-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="21f7f-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="21f7f-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-181">Method</span></span> |
| [<span data-ttu-id="21f7f-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="21f7f-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="21f7f-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-183">Method</span></span> |
| [<span data-ttu-id="21f7f-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="21f7f-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="21f7f-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-185">Method</span></span> |
| [<span data-ttu-id="21f7f-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="21f7f-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="21f7f-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-187">Method</span></span> |
| [<span data-ttu-id="21f7f-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="21f7f-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="21f7f-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-189">Method</span></span> |
| [<span data-ttu-id="21f7f-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="21f7f-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="21f7f-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-191">Method</span></span> |
| [<span data-ttu-id="21f7f-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="21f7f-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="21f7f-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-193">Method</span></span> |
| [<span data-ttu-id="21f7f-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="21f7f-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="21f7f-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-195">Method</span></span> |
| [<span data-ttu-id="21f7f-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="21f7f-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="21f7f-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-197">Method</span></span> |
| [<span data-ttu-id="21f7f-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="21f7f-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="21f7f-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-199">Method</span></span> |
| [<span data-ttu-id="21f7f-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="21f7f-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="21f7f-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-201">Method</span></span> |
| [<span data-ttu-id="21f7f-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="21f7f-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="21f7f-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-203">Method</span></span> |
| [<span data-ttu-id="21f7f-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="21f7f-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="21f7f-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="21f7f-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="21f7f-206">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-206">Example</span></span>

<span data-ttu-id="21f7f-207">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="21f7f-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="21f7f-208">Membres</span><span class="sxs-lookup"><span data-stu-id="21f7f-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="21f7f-209">pièces jointes: tableau. <[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="21f7f-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="21f7f-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-212">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="21f7f-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="21f7f-213">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="21f7f-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-214">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-214">Type</span></span>

*   <span data-ttu-id="21f7f-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="21f7f-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-216">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-216">Requirements</span></span>

|<span data-ttu-id="21f7f-217">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-217">Requirement</span></span>|<span data-ttu-id="21f7f-218">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-219">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-220">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-220">1.0</span></span>|
|[<span data-ttu-id="21f7f-221">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-222">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-224">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-225">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-225">Example</span></span>

<span data-ttu-id="21f7f-226">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="21f7f-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="21f7f-227">CCI: [destinataires](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="21f7f-227">bcc: [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="21f7f-228">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="21f7f-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="21f7f-229">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-230">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-230">Type</span></span>

*   [<span data-ttu-id="21f7f-231">Destinataires</span><span class="sxs-lookup"><span data-stu-id="21f7f-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="21f7f-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-232">Requirements</span></span>

|<span data-ttu-id="21f7f-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-233">Requirement</span></span>|<span data-ttu-id="21f7f-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-236">1.1</span><span class="sxs-lookup"><span data-stu-id="21f7f-236">1.1</span></span>|
|[<span data-ttu-id="21f7f-237">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-238">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-240">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-241">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="21f7f-242">Body: [Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="21f7f-242">body: [Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="21f7f-243">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-244">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-244">Type</span></span>

*   [<span data-ttu-id="21f7f-245">Body</span><span class="sxs-lookup"><span data-stu-id="21f7f-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="21f7f-246">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-246">Requirements</span></span>

|<span data-ttu-id="21f7f-247">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-247">Requirement</span></span>|<span data-ttu-id="21f7f-248">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-249">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-250">1.1</span><span class="sxs-lookup"><span data-stu-id="21f7f-250">1.1</span></span>|
|[<span data-ttu-id="21f7f-251">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-251">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-252">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-253">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-253">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-254">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-255">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-255">Example</span></span>

<span data-ttu-id="21f7f-256">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="21f7f-256">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="21f7f-257">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="21f7f-258">CC: Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="21f7f-258">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="21f7f-259">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="21f7f-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="21f7f-260">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="21f7f-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-261">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-261">Read mode</span></span>

<span data-ttu-id="21f7f-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-264">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-264">Compose mode</span></span>

<span data-ttu-id="21f7f-265">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="21f7f-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="21f7f-266">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-266">Type</span></span>

*   <span data-ttu-id="21f7f-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="21f7f-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-268">Requirements</span></span>

|<span data-ttu-id="21f7f-269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-269">Requirement</span></span>|<span data-ttu-id="21f7f-270">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-272">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-272">1.0</span></span>|
|[<span data-ttu-id="21f7f-273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-274">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-276">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-276">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="21f7f-277">(Nullable) conversationId: chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-277">(nullable) conversationId: String</span></span>

<span data-ttu-id="21f7f-278">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="21f7f-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="21f7f-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="21f7f-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-283">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-283">Type</span></span>

*   <span data-ttu-id="21f7f-284">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-285">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-285">Requirements</span></span>

|<span data-ttu-id="21f7f-286">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-286">Requirement</span></span>|<span data-ttu-id="21f7f-287">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-288">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-289">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-289">1.0</span></span>|
|[<span data-ttu-id="21f7f-290">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-291">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-293">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-294">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-294">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="21f7f-295">dateTimeCreated: date</span><span class="sxs-lookup"><span data-stu-id="21f7f-295">dateTimeCreated: Date</span></span>

<span data-ttu-id="21f7f-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-298">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-298">Type</span></span>

*   <span data-ttu-id="21f7f-299">Date</span><span class="sxs-lookup"><span data-stu-id="21f7f-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-300">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-300">Requirements</span></span>

|<span data-ttu-id="21f7f-301">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-301">Requirement</span></span>|<span data-ttu-id="21f7f-302">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-303">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-304">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-304">1.0</span></span>|
|[<span data-ttu-id="21f7f-305">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-306">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-307">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-308">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-309">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-309">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="21f7f-310">dateTimeModified: date</span><span class="sxs-lookup"><span data-stu-id="21f7f-310">dateTimeModified: Date</span></span>

<span data-ttu-id="21f7f-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-313">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-314">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-314">Type</span></span>

*   <span data-ttu-id="21f7f-315">Date</span><span class="sxs-lookup"><span data-stu-id="21f7f-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-316">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-316">Requirements</span></span>

|<span data-ttu-id="21f7f-317">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-317">Requirement</span></span>|<span data-ttu-id="21f7f-318">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-319">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-320">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-320">1.0</span></span>|
|[<span data-ttu-id="21f7f-321">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-321">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-322">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-323">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-323">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-324">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-325">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-325">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="21f7f-326">fin: date | [Fois](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="21f7f-326">end: Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="21f7f-327">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="21f7f-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-330">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-330">Read mode</span></span>

<span data-ttu-id="21f7f-331">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-331">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-332">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-332">Compose mode</span></span>

<span data-ttu-id="21f7f-333">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="21f7f-334">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="21f7f-334">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="21f7f-335">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="21f7f-336">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-336">Type</span></span>

*   <span data-ttu-id="21f7f-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="21f7f-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-338">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-338">Requirements</span></span>

|<span data-ttu-id="21f7f-339">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-339">Requirement</span></span>|<span data-ttu-id="21f7f-340">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-341">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-342">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-342">1.0</span></span>|
|[<span data-ttu-id="21f7f-343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-344">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-346">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-346">Compose or Read</span></span>|

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="21f7f-347">from: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[from](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="21f7f-347">from: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="21f7f-348">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="21f7f-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="21f7f-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-351">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-352">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-352">Read mode</span></span>

<span data-ttu-id="21f7f-353">La `from` propriété renvoie un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="21f7f-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-354">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-354">Compose mode</span></span>

<span data-ttu-id="21f7f-355">La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.</span><span class="sxs-lookup"><span data-stu-id="21f7f-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="21f7f-356">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-356">Type</span></span>

*   <span data-ttu-id="21f7f-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [à partir de](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="21f7f-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-358">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-358">Requirements</span></span>

|<span data-ttu-id="21f7f-359">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="21f7f-360">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-361">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-361">1.0</span></span>|<span data-ttu-id="21f7f-362">1.7</span><span class="sxs-lookup"><span data-stu-id="21f7f-362">1.7</span></span>|
|[<span data-ttu-id="21f7f-363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-364">ReadItem</span></span>|<span data-ttu-id="21f7f-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="21f7f-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-367">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-367">Read</span></span>|<span data-ttu-id="21f7f-368">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-368">Compose</span></span>|

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="21f7f-369">internetMessageId: chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-369">internetMessageId: String</span></span>

<span data-ttu-id="21f7f-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-372">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-372">Type</span></span>

*   <span data-ttu-id="21f7f-373">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-374">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-374">Requirements</span></span>

|<span data-ttu-id="21f7f-375">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-375">Requirement</span></span>|<span data-ttu-id="21f7f-376">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-377">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-378">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-378">1.0</span></span>|
|[<span data-ttu-id="21f7f-379">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-380">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-381">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-382">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-383">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-383">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="21f7f-384">itemClass: chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-384">itemClass: String</span></span>

<span data-ttu-id="21f7f-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="21f7f-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="21f7f-389">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-389">Type</span></span>|<span data-ttu-id="21f7f-390">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-390">Description</span></span>|<span data-ttu-id="21f7f-391">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="21f7f-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="21f7f-392">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="21f7f-392">Appointment items</span></span>|<span data-ttu-id="21f7f-393">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="21f7f-394">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="21f7f-394">Message items</span></span>|<span data-ttu-id="21f7f-395">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="21f7f-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="21f7f-396">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-397">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-397">Type</span></span>

*   <span data-ttu-id="21f7f-398">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-399">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-399">Requirements</span></span>

|<span data-ttu-id="21f7f-400">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-400">Requirement</span></span>|<span data-ttu-id="21f7f-401">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-402">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-403">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-403">1.0</span></span>|
|[<span data-ttu-id="21f7f-404">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-405">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-406">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-407">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-408">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-408">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="21f7f-409">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="21f7f-409">(nullable) itemId: String</span></span>

<span data-ttu-id="21f7f-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-412">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="21f7f-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="21f7f-413">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="21f7f-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="21f7f-414">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="21f7f-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="21f7f-415">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="21f7f-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="21f7f-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-418">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-418">Type</span></span>

*   <span data-ttu-id="21f7f-419">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-420">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-420">Requirements</span></span>

|<span data-ttu-id="21f7f-421">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-421">Requirement</span></span>|<span data-ttu-id="21f7f-422">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-423">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-424">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-424">1.0</span></span>|
|[<span data-ttu-id="21f7f-425">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-426">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-427">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-428">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-429">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-429">Example</span></span>

<span data-ttu-id="21f7f-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="21f7f-432">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="21f7f-432">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="21f7f-433">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="21f7f-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="21f7f-434">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-435">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-435">Type</span></span>

*   [<span data-ttu-id="21f7f-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="21f7f-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="21f7f-437">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-437">Requirements</span></span>

|<span data-ttu-id="21f7f-438">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-438">Requirement</span></span>|<span data-ttu-id="21f7f-439">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-440">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-441">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-441">1.0</span></span>|
|[<span data-ttu-id="21f7f-442">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-443">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-444">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-445">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-446">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-446">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="21f7f-447">Location: String | [Emplacement](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="21f7f-447">location: String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="21f7f-448">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-449">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-449">Read mode</span></span>

<span data-ttu-id="21f7f-450">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-451">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-451">Compose mode</span></span>

<span data-ttu-id="21f7f-452">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="21f7f-453">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-453">Type</span></span>

*   <span data-ttu-id="21f7f-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="21f7f-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-455">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-455">Requirements</span></span>

|<span data-ttu-id="21f7f-456">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-456">Requirement</span></span>|<span data-ttu-id="21f7f-457">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-458">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-459">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-459">1.0</span></span>|
|[<span data-ttu-id="21f7f-460">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-461">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-462">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-463">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-463">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="21f7f-464">normalizedSubject: chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-464">normalizedSubject: String</span></span>

<span data-ttu-id="21f7f-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="21f7f-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="21f7f-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-469">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-469">Type</span></span>

*   <span data-ttu-id="21f7f-470">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-471">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-471">Requirements</span></span>

|<span data-ttu-id="21f7f-472">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-472">Requirement</span></span>|<span data-ttu-id="21f7f-473">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-474">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-475">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-475">1.0</span></span>|
|[<span data-ttu-id="21f7f-476">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-476">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-477">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-478">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-478">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-479">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-480">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-480">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="21f7f-481">notificationMessages: [notificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="21f7f-481">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="21f7f-482">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-483">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-483">Type</span></span>

*   [<span data-ttu-id="21f7f-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="21f7f-484">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="21f7f-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-485">Requirements</span></span>

|<span data-ttu-id="21f7f-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-486">Requirement</span></span>|<span data-ttu-id="21f7f-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-489">1.3</span><span class="sxs-lookup"><span data-stu-id="21f7f-489">1.3</span></span>|
|[<span data-ttu-id="21f7f-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-491">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-493">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-494">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-494">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="21f7f-495">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[](/javascript/api/outlook_1_7/office.recipients) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="21f7f-495">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="21f7f-496">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="21f7f-497">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="21f7f-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-498">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-498">Read mode</span></span>

<span data-ttu-id="21f7f-499">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-500">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-500">Compose mode</span></span>

<span data-ttu-id="21f7f-501">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="21f7f-502">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-502">Type</span></span>

*   <span data-ttu-id="21f7f-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="21f7f-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-504">Requirements</span></span>

|<span data-ttu-id="21f7f-505">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-505">Requirement</span></span>|<span data-ttu-id="21f7f-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-507">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-508">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-508">1.0</span></span>|
|[<span data-ttu-id="21f7f-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-510">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-512">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-512">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="21f7f-513">Organisateur: [](/javascript/api/outlook_1_7/office.emailaddressdetails)|[organisateur](/javascript/api/outlook_1_7/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="21f7f-513">organizer: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="21f7f-514">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="21f7f-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-515">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-515">Read mode</span></span>

<span data-ttu-id="21f7f-516">La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-517">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-517">Compose mode</span></span>

<span data-ttu-id="21f7f-518">La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook_1_7/office.organizer) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="21f7f-518">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="21f7f-519">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-519">Type</span></span>

*   <span data-ttu-id="21f7f-520">[](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organisateur](/javascript/api/outlook_1_7/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="21f7f-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-521">Requirements</span></span>

|<span data-ttu-id="21f7f-522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="21f7f-523">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-524">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-524">1.0</span></span>|<span data-ttu-id="21f7f-525">1.7</span><span class="sxs-lookup"><span data-stu-id="21f7f-525">1.7</span></span>|
|[<span data-ttu-id="21f7f-526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-527">ReadItem</span></span>|<span data-ttu-id="21f7f-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="21f7f-529">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-530">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-530">Read</span></span>|<span data-ttu-id="21f7f-531">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-531">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="21f7f-532">(Nullable) récurrence: [périodicité](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="21f7f-532">(nullable) recurrence: [Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="21f7f-533">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="21f7f-534">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="21f7f-535">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="21f7f-536">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="21f7f-537">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook_1_7/office.recurrence) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="21f7f-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="21f7f-538">`null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="21f7f-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="21f7f-539">`undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="21f7f-540">Remarque: les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="21f7f-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="21f7f-541">Remarque: si l’objet de périodicité `null`est, cela indique que l’objet est un rendez-vous unique ou une demande de réunion d’un seul rendez-vous et non d’une série.</span><span class="sxs-lookup"><span data-stu-id="21f7f-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-542">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-542">Read mode</span></span>

<span data-ttu-id="21f7f-543">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook_1_7/office.recurrence) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="21f7f-544">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-544">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-545">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-545">Compose mode</span></span>

<span data-ttu-id="21f7f-546">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook_1_7/office.recurrence) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="21f7f-547">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="21f7f-548">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-548">Type</span></span>

* [<span data-ttu-id="21f7f-549">Instances</span><span class="sxs-lookup"><span data-stu-id="21f7f-549">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="21f7f-550">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-550">Requirement</span></span>|<span data-ttu-id="21f7f-551">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-552">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-553">1.7</span><span class="sxs-lookup"><span data-stu-id="21f7f-553">1.7</span></span>|
|[<span data-ttu-id="21f7f-554">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-554">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-555">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-556">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-556">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-557">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-557">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="21f7f-558">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[](/javascript/api/outlook_1_7/office.recipients) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="21f7f-558">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="21f7f-559">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="21f7f-560">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="21f7f-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-561">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-561">Read mode</span></span>

<span data-ttu-id="21f7f-562">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-563">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-563">Compose mode</span></span>

<span data-ttu-id="21f7f-564">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="21f7f-565">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-565">Type</span></span>

*   <span data-ttu-id="21f7f-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="21f7f-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-567">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-567">Requirements</span></span>

|<span data-ttu-id="21f7f-568">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-568">Requirement</span></span>|<span data-ttu-id="21f7f-569">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-570">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-571">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-571">1.0</span></span>|
|[<span data-ttu-id="21f7f-572">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-572">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-573">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-574">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-574">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-575">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-575">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="21f7f-576">expéditeur: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="21f7f-576">sender: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="21f7f-p128">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="21f7f-p129">Les propriétés [`from`](#from-emailaddressdetailsfrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-581">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-582">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-582">Type</span></span>

*   [<span data-ttu-id="21f7f-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="21f7f-583">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="21f7f-584">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-584">Requirements</span></span>

|<span data-ttu-id="21f7f-585">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-585">Requirement</span></span>|<span data-ttu-id="21f7f-586">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-587">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-588">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-588">1.0</span></span>|
|[<span data-ttu-id="21f7f-589">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-590">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-591">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-592">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-593">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-593">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="21f7f-594">(Nullable) seriesId: chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-594">(nullable) seriesId: String</span></span>

<span data-ttu-id="21f7f-595">Obtient l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="21f7f-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="21f7f-596">Dans OWA et Outlook, le `seriesId` renvoie l’ID des services Web Exchange (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="21f7f-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="21f7f-597">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="21f7f-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-598">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="21f7f-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="21f7f-599">La `seriesId` propriété n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="21f7f-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="21f7f-600">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="21f7f-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="21f7f-601">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="21f7f-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="21f7f-602">La `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="21f7f-603">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-603">Type</span></span>

* <span data-ttu-id="21f7f-604">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-605">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-605">Requirements</span></span>

|<span data-ttu-id="21f7f-606">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-606">Requirement</span></span>|<span data-ttu-id="21f7f-607">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-608">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-609">1.7</span><span class="sxs-lookup"><span data-stu-id="21f7f-609">1.7</span></span>|
|[<span data-ttu-id="21f7f-610">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-611">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-612">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-613">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-614">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-614">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="21f7f-615">début: date | [Fois](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="21f7f-615">start: Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="21f7f-616">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="21f7f-p132">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-619">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-619">Read mode</span></span>

<span data-ttu-id="21f7f-620">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-620">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-621">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-621">Compose mode</span></span>

<span data-ttu-id="21f7f-622">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="21f7f-623">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="21f7f-623">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="21f7f-624">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="21f7f-625">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-625">Type</span></span>

*   <span data-ttu-id="21f7f-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="21f7f-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-627">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-627">Requirements</span></span>

|<span data-ttu-id="21f7f-628">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-628">Requirement</span></span>|<span data-ttu-id="21f7f-629">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-630">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-631">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-631">1.0</span></span>|
|[<span data-ttu-id="21f7f-632">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-632">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-633">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-634">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-634">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-635">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-635">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="21f7f-636">Subject: String | [Objet](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="21f7f-636">subject: String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="21f7f-637">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="21f7f-638">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="21f7f-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-639">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-639">Read mode</span></span>

<span data-ttu-id="21f7f-p133">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="21f7f-642">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="21f7f-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-643">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-643">Compose mode</span></span>

<span data-ttu-id="21f7f-644">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="21f7f-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="21f7f-645">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-645">Type</span></span>

*   <span data-ttu-id="21f7f-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="21f7f-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-647">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-647">Requirements</span></span>

|<span data-ttu-id="21f7f-648">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-648">Requirement</span></span>|<span data-ttu-id="21f7f-649">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-650">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-651">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-651">1.0</span></span>|
|[<span data-ttu-id="21f7f-652">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-653">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-654">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-655">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-655">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="21f7f-656">to: Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="21f7f-656">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="21f7f-657">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="21f7f-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="21f7f-658">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="21f7f-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="21f7f-659">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-659">Read mode</span></span>

<span data-ttu-id="21f7f-p135">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="21f7f-662">Mode composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-662">Compose mode</span></span>

<span data-ttu-id="21f7f-663">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="21f7f-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="21f7f-664">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-664">Type</span></span>

*   <span data-ttu-id="21f7f-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="21f7f-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-666">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-666">Requirements</span></span>

|<span data-ttu-id="21f7f-667">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-667">Requirement</span></span>|<span data-ttu-id="21f7f-668">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-669">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-670">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-670">1.0</span></span>|
|[<span data-ttu-id="21f7f-671">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-672">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-673">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-674">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="21f7f-675">Méthodes</span><span class="sxs-lookup"><span data-stu-id="21f7f-675">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="21f7f-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="21f7f-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="21f7f-677">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="21f7f-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="21f7f-678">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="21f7f-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="21f7f-679">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="21f7f-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-680">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-680">Parameters</span></span>
|<span data-ttu-id="21f7f-681">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-681">Name</span></span>|<span data-ttu-id="21f7f-682">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-682">Type</span></span>|<span data-ttu-id="21f7f-683">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-683">Attributes</span></span>|<span data-ttu-id="21f7f-684">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="21f7f-685">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-685">String</span></span>||<span data-ttu-id="21f7f-p136">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="21f7f-688">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-688">String</span></span>||<span data-ttu-id="21f7f-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="21f7f-691">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-691">Object</span></span>|<span data-ttu-id="21f7f-692">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-692">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-693">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="21f7f-694">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-694">Object</span></span>|<span data-ttu-id="21f7f-695">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-695">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-696">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="21f7f-697">Boolean</span><span class="sxs-lookup"><span data-stu-id="21f7f-697">Boolean</span></span>|<span data-ttu-id="21f7f-698">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-698">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-699">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="21f7f-700">fonction</span><span class="sxs-lookup"><span data-stu-id="21f7f-700">function</span></span>|<span data-ttu-id="21f7f-701">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-701">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-702">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="21f7f-703">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="21f7f-704">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="21f7f-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="21f7f-705">Erreurs</span><span class="sxs-lookup"><span data-stu-id="21f7f-705">Errors</span></span>

|<span data-ttu-id="21f7f-706">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="21f7f-706">Error code</span></span>|<span data-ttu-id="21f7f-707">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="21f7f-708">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="21f7f-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="21f7f-709">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="21f7f-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="21f7f-710">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-711">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-711">Requirements</span></span>

|<span data-ttu-id="21f7f-712">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-712">Requirement</span></span>|<span data-ttu-id="21f7f-713">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-714">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-715">1.1</span><span class="sxs-lookup"><span data-stu-id="21f7f-715">1.1</span></span>|
|[<span data-ttu-id="21f7f-716">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-716">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="21f7f-718">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-718">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-719">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="21f7f-720">Exemples</span><span class="sxs-lookup"><span data-stu-id="21f7f-720">Examples</span></span>

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

<span data-ttu-id="21f7f-721">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="21f7f-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="21f7f-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="21f7f-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="21f7f-723">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="21f7f-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="21f7f-724">Actuellement, les types d’événement `Office.EventType.AppointmentTimeChanged`pris `Office.EventType.RecipientsChanged`en charge sont, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="21f7f-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-725">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-725">Parameters</span></span>

| <span data-ttu-id="21f7f-726">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-726">Name</span></span> | <span data-ttu-id="21f7f-727">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-727">Type</span></span> | <span data-ttu-id="21f7f-728">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-728">Attributes</span></span> | <span data-ttu-id="21f7f-729">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="21f7f-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="21f7f-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="21f7f-731">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="21f7f-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="21f7f-732">Fonction</span><span class="sxs-lookup"><span data-stu-id="21f7f-732">Function</span></span> || <span data-ttu-id="21f7f-p138">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="21f7f-736">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-736">Object</span></span> | <span data-ttu-id="21f7f-737">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-737">&lt;optional&gt;</span></span> | <span data-ttu-id="21f7f-738">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="21f7f-739">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-739">Object</span></span> | <span data-ttu-id="21f7f-740">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-740">&lt;optional&gt;</span></span> | <span data-ttu-id="21f7f-741">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="21f7f-742">fonction</span><span class="sxs-lookup"><span data-stu-id="21f7f-742">function</span></span>| <span data-ttu-id="21f7f-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-743">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-744">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-745">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-745">Requirements</span></span>

|<span data-ttu-id="21f7f-746">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-746">Requirement</span></span>| <span data-ttu-id="21f7f-747">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-748">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21f7f-749">1.7</span><span class="sxs-lookup"><span data-stu-id="21f7f-749">1.7</span></span> |
|[<span data-ttu-id="21f7f-750">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21f7f-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-751">ReadItem</span></span> |
|[<span data-ttu-id="21f7f-752">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="21f7f-753">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="21f7f-754">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-754">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="21f7f-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="21f7f-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="21f7f-756">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="21f7f-p139">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="21f7f-760">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="21f7f-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="21f7f-761">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="21f7f-761">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-762">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-762">Parameters</span></span>

|<span data-ttu-id="21f7f-763">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-763">Name</span></span>|<span data-ttu-id="21f7f-764">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-764">Type</span></span>|<span data-ttu-id="21f7f-765">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-765">Attributes</span></span>|<span data-ttu-id="21f7f-766">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="21f7f-767">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-767">String</span></span>||<span data-ttu-id="21f7f-p140">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="21f7f-770">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-770">String</span></span>||<span data-ttu-id="21f7f-771">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="21f7f-771">The subject of the item to be attached.</span></span> <span data-ttu-id="21f7f-772">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="21f7f-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="21f7f-773">Object</span><span class="sxs-lookup"><span data-stu-id="21f7f-773">Object</span></span>|<span data-ttu-id="21f7f-774">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-774">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-775">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="21f7f-776">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-776">Object</span></span>|<span data-ttu-id="21f7f-777">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-777">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-778">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="21f7f-779">fonction</span><span class="sxs-lookup"><span data-stu-id="21f7f-779">function</span></span>|<span data-ttu-id="21f7f-780">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-780">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-781">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="21f7f-782">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="21f7f-783">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="21f7f-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="21f7f-784">Erreurs</span><span class="sxs-lookup"><span data-stu-id="21f7f-784">Errors</span></span>

|<span data-ttu-id="21f7f-785">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="21f7f-785">Error code</span></span>|<span data-ttu-id="21f7f-786">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="21f7f-787">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-788">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-788">Requirements</span></span>

|<span data-ttu-id="21f7f-789">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-789">Requirement</span></span>|<span data-ttu-id="21f7f-790">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-791">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-792">1.1</span><span class="sxs-lookup"><span data-stu-id="21f7f-792">1.1</span></span>|
|[<span data-ttu-id="21f7f-793">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="21f7f-795">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-796">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-797">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-797">Example</span></span>

<span data-ttu-id="21f7f-798">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="21f7f-799">close()</span><span class="sxs-lookup"><span data-stu-id="21f7f-799">close()</span></span>

<span data-ttu-id="21f7f-800">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="21f7f-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="21f7f-p142">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-803">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="21f7f-804">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="21f7f-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-805">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-805">Requirements</span></span>

|<span data-ttu-id="21f7f-806">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-806">Requirement</span></span>|<span data-ttu-id="21f7f-807">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-808">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-809">1.3</span><span class="sxs-lookup"><span data-stu-id="21f7f-809">1.3</span></span>|
|[<span data-ttu-id="21f7f-810">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-811">Restreinte</span><span class="sxs-lookup"><span data-stu-id="21f7f-811">Restricted</span></span>|
|[<span data-ttu-id="21f7f-812">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-813">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-813">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="21f7f-814">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="21f7f-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="21f7f-815">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="21f7f-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-816">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-816">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="21f7f-817">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-817">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="21f7f-818">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="21f7f-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="21f7f-p143">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-822">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-822">Parameters</span></span>

|<span data-ttu-id="21f7f-823">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-823">Name</span></span>|<span data-ttu-id="21f7f-824">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-824">Type</span></span>|<span data-ttu-id="21f7f-825">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-825">Attributes</span></span>|<span data-ttu-id="21f7f-826">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="21f7f-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="21f7f-827">String &#124; Object</span></span>||<span data-ttu-id="21f7f-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="21f7f-830">**OU**</span><span class="sxs-lookup"><span data-stu-id="21f7f-830">**OR**</span></span><br/><span data-ttu-id="21f7f-p145">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="21f7f-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="21f7f-833">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-833">String</span></span>|<span data-ttu-id="21f7f-834">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-834">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="21f7f-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="21f7f-838">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-838">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-839">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="21f7f-840">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-840">String</span></span>||<span data-ttu-id="21f7f-p147">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="21f7f-843">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-843">String</span></span>||<span data-ttu-id="21f7f-844">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="21f7f-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="21f7f-845">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-845">String</span></span>||<span data-ttu-id="21f7f-p148">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="21f7f-848">Booléen</span><span class="sxs-lookup"><span data-stu-id="21f7f-848">Boolean</span></span>||<span data-ttu-id="21f7f-p149">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="21f7f-851">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-851">String</span></span>||<span data-ttu-id="21f7f-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="21f7f-855">function</span><span class="sxs-lookup"><span data-stu-id="21f7f-855">function</span></span>|<span data-ttu-id="21f7f-856">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-856">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-857">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-858">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-858">Requirements</span></span>

|<span data-ttu-id="21f7f-859">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-859">Requirement</span></span>|<span data-ttu-id="21f7f-860">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-861">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-862">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-862">1.0</span></span>|
|[<span data-ttu-id="21f7f-863">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-864">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-865">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-866">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="21f7f-867">Exemples</span><span class="sxs-lookup"><span data-stu-id="21f7f-867">Examples</span></span>

<span data-ttu-id="21f7f-868">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="21f7f-869">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="21f7f-869">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="21f7f-870">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="21f7f-870">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="21f7f-871">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="21f7f-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="21f7f-872">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="21f7f-873">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="21f7f-874">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="21f7f-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="21f7f-875">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="21f7f-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-876">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-876">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="21f7f-877">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-877">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="21f7f-878">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="21f7f-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="21f7f-p151">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-882">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-882">Parameters</span></span>

|<span data-ttu-id="21f7f-883">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-883">Name</span></span>|<span data-ttu-id="21f7f-884">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-884">Type</span></span>|<span data-ttu-id="21f7f-885">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-885">Attributes</span></span>|<span data-ttu-id="21f7f-886">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="21f7f-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="21f7f-887">String &#124; Object</span></span>||<span data-ttu-id="21f7f-p152">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="21f7f-890">**OU**</span><span class="sxs-lookup"><span data-stu-id="21f7f-890">**OR**</span></span><br/><span data-ttu-id="21f7f-p153">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="21f7f-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="21f7f-893">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-893">String</span></span>|<span data-ttu-id="21f7f-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-894">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="21f7f-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="21f7f-898">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-898">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-899">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="21f7f-900">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-900">String</span></span>||<span data-ttu-id="21f7f-p155">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="21f7f-903">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-903">String</span></span>||<span data-ttu-id="21f7f-904">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="21f7f-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="21f7f-905">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-905">String</span></span>||<span data-ttu-id="21f7f-p156">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="21f7f-908">Booléen</span><span class="sxs-lookup"><span data-stu-id="21f7f-908">Boolean</span></span>||<span data-ttu-id="21f7f-p157">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="21f7f-911">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-911">String</span></span>||<span data-ttu-id="21f7f-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="21f7f-915">function</span><span class="sxs-lookup"><span data-stu-id="21f7f-915">function</span></span>|<span data-ttu-id="21f7f-916">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-916">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-917">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-918">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-918">Requirements</span></span>

|<span data-ttu-id="21f7f-919">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-919">Requirement</span></span>|<span data-ttu-id="21f7f-920">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-921">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-922">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-922">1.0</span></span>|
|[<span data-ttu-id="21f7f-923">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-923">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-924">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-925">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-925">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-926">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="21f7f-927">Exemples</span><span class="sxs-lookup"><span data-stu-id="21f7f-927">Examples</span></span>

<span data-ttu-id="21f7f-928">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="21f7f-929">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="21f7f-929">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="21f7f-930">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="21f7f-930">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="21f7f-931">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="21f7f-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="21f7f-932">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="21f7f-933">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="21f7f-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="21f7f-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="21f7f-935">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="21f7f-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-936">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-936">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-937">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-937">Requirements</span></span>

|<span data-ttu-id="21f7f-938">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-938">Requirement</span></span>|<span data-ttu-id="21f7f-939">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-940">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-941">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-941">1.0</span></span>|
|[<span data-ttu-id="21f7f-942">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-942">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-943">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-944">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-944">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-945">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="21f7f-946">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="21f7f-946">Returns:</span></span>

<span data-ttu-id="21f7f-947">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="21f7f-947">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="21f7f-948">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-948">Example</span></span>

<span data-ttu-id="21f7f-949">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="21f7f-949">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="21f7f-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="21f7f-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="21f7f-951">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="21f7f-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-952">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-952">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-953">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-953">Parameters</span></span>

|<span data-ttu-id="21f7f-954">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-954">Name</span></span>|<span data-ttu-id="21f7f-955">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-955">Type</span></span>|<span data-ttu-id="21f7f-956">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="21f7f-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="21f7f-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="21f7f-958">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="21f7f-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-959">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-959">Requirements</span></span>

|<span data-ttu-id="21f7f-960">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-960">Requirement</span></span>|<span data-ttu-id="21f7f-961">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-962">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-963">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-963">1.0</span></span>|
|[<span data-ttu-id="21f7f-964">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-964">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-965">Restreinte</span><span class="sxs-lookup"><span data-stu-id="21f7f-965">Restricted</span></span>|
|[<span data-ttu-id="21f7f-966">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-966">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-967">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="21f7f-968">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="21f7f-968">Returns:</span></span>

<span data-ttu-id="21f7f-969">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="21f7f-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="21f7f-970">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="21f7f-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="21f7f-971">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="21f7f-972">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="21f7f-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="21f7f-973">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="21f7f-973">Value of `entityType`</span></span>|<span data-ttu-id="21f7f-974">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="21f7f-974">Type of objects in returned array</span></span>|<span data-ttu-id="21f7f-975">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="21f7f-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="21f7f-976">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-976">String</span></span>|<span data-ttu-id="21f7f-977">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="21f7f-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="21f7f-978">Contact</span><span class="sxs-lookup"><span data-stu-id="21f7f-978">Contact</span></span>|<span data-ttu-id="21f7f-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="21f7f-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="21f7f-980">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-980">String</span></span>|<span data-ttu-id="21f7f-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="21f7f-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="21f7f-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="21f7f-982">MeetingSuggestion</span></span>|<span data-ttu-id="21f7f-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="21f7f-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="21f7f-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="21f7f-984">PhoneNumber</span></span>|<span data-ttu-id="21f7f-985">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="21f7f-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="21f7f-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="21f7f-986">TaskSuggestion</span></span>|<span data-ttu-id="21f7f-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="21f7f-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="21f7f-988">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-988">String</span></span>|<span data-ttu-id="21f7f-989">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="21f7f-989">**Restricted**</span></span>|

<span data-ttu-id="21f7f-990">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="21f7f-990">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="21f7f-991">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-991">Example</span></span>

<span data-ttu-id="21f7f-992">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="21f7f-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="21f7f-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="21f7f-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="21f7f-994">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="21f7f-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-995">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-995">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="21f7f-996">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="21f7f-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-997">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-997">Parameters</span></span>

|<span data-ttu-id="21f7f-998">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-998">Name</span></span>|<span data-ttu-id="21f7f-999">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-999">Type</span></span>|<span data-ttu-id="21f7f-1000">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="21f7f-1001">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-1001">String</span></span>|<span data-ttu-id="21f7f-1002">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-1003">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1003">Requirements</span></span>

|<span data-ttu-id="21f7f-1004">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1004">Requirement</span></span>|<span data-ttu-id="21f7f-1005">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1006">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-1007">1.0</span></span>|
|[<span data-ttu-id="21f7f-1008">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1008">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1009">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-1010">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1010">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1011">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="21f7f-1012">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="21f7f-1012">Returns:</span></span>

<span data-ttu-id="21f7f-p160">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="21f7f-1015">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="21f7f-1015">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="21f7f-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="21f7f-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="21f7f-1017">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-1018">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1018">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="21f7f-p161">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="21f7f-1022">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="21f7f-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="21f7f-1023">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="21f7f-p162">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-1027">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1027">Requirements</span></span>

|<span data-ttu-id="21f7f-1028">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1028">Requirement</span></span>|<span data-ttu-id="21f7f-1029">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1030">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-1031">1.0</span></span>|
|[<span data-ttu-id="21f7f-1032">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1032">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1033">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-1034">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1034">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1035">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="21f7f-1036">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="21f7f-1036">Returns:</span></span>

<span data-ttu-id="21f7f-p163">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="21f7f-1039">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="21f7f-1039">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="21f7f-1040">Object</span><span class="sxs-lookup"><span data-stu-id="21f7f-1040">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="21f7f-1041">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-1041">Example</span></span>

<span data-ttu-id="21f7f-1042">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1042">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="21f7f-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="21f7f-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="21f7f-1044">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1044">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-1045">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1045">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="21f7f-1046">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1046">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="21f7f-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-1049">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1049">Parameters</span></span>

|<span data-ttu-id="21f7f-1050">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-1050">Name</span></span>|<span data-ttu-id="21f7f-1051">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-1051">Type</span></span>|<span data-ttu-id="21f7f-1052">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-1052">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="21f7f-1053">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-1053">String</span></span>|<span data-ttu-id="21f7f-1054">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1054">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-1055">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1055">Requirements</span></span>

|<span data-ttu-id="21f7f-1056">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1056">Requirement</span></span>|<span data-ttu-id="21f7f-1057">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1057">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1058">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1058">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1059">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-1059">1.0</span></span>|
|[<span data-ttu-id="21f7f-1060">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1060">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1061">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1061">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-1062">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1062">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1063">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-1063">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="21f7f-1064">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="21f7f-1064">Returns:</span></span>

<span data-ttu-id="21f7f-1065">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1065">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="21f7f-1066">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="21f7f-1066">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="21f7f-1067">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="21f7f-1067">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="21f7f-1068">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-1068">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="21f7f-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="21f7f-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="21f7f-1070">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1070">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="21f7f-p165">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-1073">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1073">Parameters</span></span>

|<span data-ttu-id="21f7f-1074">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-1074">Name</span></span>|<span data-ttu-id="21f7f-1075">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-1075">Type</span></span>|<span data-ttu-id="21f7f-1076">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-1076">Attributes</span></span>|<span data-ttu-id="21f7f-1077">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-1077">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="21f7f-1078">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="21f7f-1078">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="21f7f-p166">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="21f7f-1082">Object</span><span class="sxs-lookup"><span data-stu-id="21f7f-1082">Object</span></span>|<span data-ttu-id="21f7f-1083">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1084">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="21f7f-1085">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-1085">Object</span></span>|<span data-ttu-id="21f7f-1086">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1087">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="21f7f-1088">fonction</span><span class="sxs-lookup"><span data-stu-id="21f7f-1088">function</span></span>||<span data-ttu-id="21f7f-1089">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="21f7f-1090">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1090">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="21f7f-1091">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1091">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-1092">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1092">Requirements</span></span>

|<span data-ttu-id="21f7f-1093">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1093">Requirement</span></span>|<span data-ttu-id="21f7f-1094">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1094">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1095">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1095">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1096">1.2</span><span class="sxs-lookup"><span data-stu-id="21f7f-1096">1.2</span></span>|
|[<span data-ttu-id="21f7f-1097">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1097">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1098">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1098">ReadWriteItem</span></span>|
|[<span data-ttu-id="21f7f-1099">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1099">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1100">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-1100">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="21f7f-1101">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="21f7f-1101">Returns:</span></span>

<span data-ttu-id="21f7f-1102">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1102">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="21f7f-1103">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="21f7f-1103">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="21f7f-1104">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-1104">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="21f7f-1105">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-1105">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="21f7f-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="21f7f-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="21f7f-1107">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1107">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="21f7f-1108">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="21f7f-1108">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-1109">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1109">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-1110">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1110">Requirements</span></span>

|<span data-ttu-id="21f7f-1111">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1111">Requirement</span></span>|<span data-ttu-id="21f7f-1112">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1113">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1114">1.6</span><span class="sxs-lookup"><span data-stu-id="21f7f-1114">1.6</span></span>|
|[<span data-ttu-id="21f7f-1115">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1116">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1116">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-1117">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1118">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-1118">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="21f7f-1119">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="21f7f-1119">Returns:</span></span>

<span data-ttu-id="21f7f-1120">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="21f7f-1120">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="21f7f-1121">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-1121">Example</span></span>

<span data-ttu-id="21f7f-1122">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1122">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="21f7f-1123">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="21f7f-1123">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="21f7f-p169">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="21f7f-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-1126">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="21f7f-p170">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="21f7f-1130">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="21f7f-1130">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="21f7f-1131">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1131">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="21f7f-p171">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="21f7f-1135">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1135">Requirements</span></span>

|<span data-ttu-id="21f7f-1136">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1136">Requirement</span></span>|<span data-ttu-id="21f7f-1137">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1138">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="21f7f-1139">1.6</span></span>|
|[<span data-ttu-id="21f7f-1140">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1141">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-1142">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1143">Lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="21f7f-1144">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="21f7f-1144">Returns:</span></span>

<span data-ttu-id="21f7f-p172">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="21f7f-1147">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-1147">Example</span></span>

<span data-ttu-id="21f7f-1148">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1148">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="21f7f-1149">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="21f7f-1149">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="21f7f-1150">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1150">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="21f7f-p173">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-1154">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1154">Parameters</span></span>

|<span data-ttu-id="21f7f-1155">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-1155">Name</span></span>|<span data-ttu-id="21f7f-1156">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-1156">Type</span></span>|<span data-ttu-id="21f7f-1157">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-1157">Attributes</span></span>|<span data-ttu-id="21f7f-1158">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-1158">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="21f7f-1159">function</span><span class="sxs-lookup"><span data-stu-id="21f7f-1159">function</span></span>||<span data-ttu-id="21f7f-1160">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-1160">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="21f7f-1161">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1161">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="21f7f-1162">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1162">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="21f7f-1163">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-1163">Object</span></span>|<span data-ttu-id="21f7f-1164">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1164">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1165">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1165">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="21f7f-1166">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1166">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-1167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1167">Requirements</span></span>

|<span data-ttu-id="21f7f-1168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1168">Requirement</span></span>|<span data-ttu-id="21f7f-1169">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1171">1.0</span><span class="sxs-lookup"><span data-stu-id="21f7f-1171">1.0</span></span>|
|[<span data-ttu-id="21f7f-1172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1173">ReadItem</span></span>|
|[<span data-ttu-id="21f7f-1174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-1175">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-1176">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-1176">Example</span></span>

<span data-ttu-id="21f7f-p176">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="21f7f-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="21f7f-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="21f7f-1181">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1181">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="21f7f-p177">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-1186">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1186">Parameters</span></span>

|<span data-ttu-id="21f7f-1187">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-1187">Name</span></span>|<span data-ttu-id="21f7f-1188">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-1188">Type</span></span>|<span data-ttu-id="21f7f-1189">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-1189">Attributes</span></span>|<span data-ttu-id="21f7f-1190">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-1190">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="21f7f-1191">Chaîne</span><span class="sxs-lookup"><span data-stu-id="21f7f-1191">String</span></span>||<span data-ttu-id="21f7f-1192">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1192">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="21f7f-1193">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-1193">Object</span></span>|<span data-ttu-id="21f7f-1194">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1195">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1195">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="21f7f-1196">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-1196">Object</span></span>|<span data-ttu-id="21f7f-1197">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1198">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1198">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="21f7f-1199">fonction</span><span class="sxs-lookup"><span data-stu-id="21f7f-1199">function</span></span>|<span data-ttu-id="21f7f-1200">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1201">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-1201">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="21f7f-1202">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1202">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="21f7f-1203">Erreurs</span><span class="sxs-lookup"><span data-stu-id="21f7f-1203">Errors</span></span>

|<span data-ttu-id="21f7f-1204">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1204">Error code</span></span>|<span data-ttu-id="21f7f-1205">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-1205">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="21f7f-1206">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1206">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-1207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1207">Requirements</span></span>

|<span data-ttu-id="21f7f-1208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1208">Requirement</span></span>|<span data-ttu-id="21f7f-1209">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1209">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1211">1.1</span><span class="sxs-lookup"><span data-stu-id="21f7f-1211">1.1</span></span>|
|[<span data-ttu-id="21f7f-1212">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1212">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1213">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1213">ReadWriteItem</span></span>|
|[<span data-ttu-id="21f7f-1214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1214">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1215">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-1215">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-1216">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-1216">Example</span></span>

<span data-ttu-id="21f7f-1217">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="21f7f-1217">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="21f7f-1218">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="21f7f-1218">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="21f7f-1219">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1219">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="21f7f-1220">Actuellement, les types d’événement `Office.EventType.AppointmentTimeChanged`pris `Office.EventType.RecipientsChanged`en charge sont, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="21f7f-1220">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-1221">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1221">Parameters</span></span>

| <span data-ttu-id="21f7f-1222">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-1222">Name</span></span> | <span data-ttu-id="21f7f-1223">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-1223">Type</span></span> | <span data-ttu-id="21f7f-1224">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-1224">Attributes</span></span> | <span data-ttu-id="21f7f-1225">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-1225">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="21f7f-1226">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="21f7f-1226">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="21f7f-1227">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1227">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="21f7f-1228">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-1228">Object</span></span> | <span data-ttu-id="21f7f-1229">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="21f7f-1230">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="21f7f-1231">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-1231">Object</span></span> | <span data-ttu-id="21f7f-1232">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="21f7f-1233">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="21f7f-1234">fonction</span><span class="sxs-lookup"><span data-stu-id="21f7f-1234">function</span></span>| <span data-ttu-id="21f7f-1235">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1236">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-1237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1237">Requirements</span></span>

|<span data-ttu-id="21f7f-1238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1238">Requirement</span></span>| <span data-ttu-id="21f7f-1239">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21f7f-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="21f7f-1241">1.7</span></span> |
|[<span data-ttu-id="21f7f-1242">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1242">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21f7f-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1243">ReadItem</span></span> |
|[<span data-ttu-id="21f7f-1244">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1244">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="21f7f-1245">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="21f7f-1245">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="21f7f-1246">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-1246">Example</span></span>

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

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="21f7f-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="21f7f-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="21f7f-1248">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="21f7f-p178">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-1252">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="21f7f-1253">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="21f7f-p180">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="21f7f-1257">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="21f7f-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="21f7f-1258">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1258">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="21f7f-1259">La `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1259">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="21f7f-1260">Consultez la rubrique [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide de l’API Office js](https://support.microsoft.com/help/4505745) pour obtenir une solution de contournement.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1260">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="21f7f-1261">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1261">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-1262">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1262">Parameters</span></span>

|<span data-ttu-id="21f7f-1263">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-1263">Name</span></span>|<span data-ttu-id="21f7f-1264">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-1264">Type</span></span>|<span data-ttu-id="21f7f-1265">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-1265">Attributes</span></span>|<span data-ttu-id="21f7f-1266">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-1266">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="21f7f-1267">Object</span><span class="sxs-lookup"><span data-stu-id="21f7f-1267">Object</span></span>|<span data-ttu-id="21f7f-1268">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1269">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="21f7f-1270">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-1270">Object</span></span>|<span data-ttu-id="21f7f-1271">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1272">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="21f7f-1273">fonction</span><span class="sxs-lookup"><span data-stu-id="21f7f-1273">function</span></span>||<span data-ttu-id="21f7f-1274">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="21f7f-1275">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1275">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-1276">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1276">Requirements</span></span>

|<span data-ttu-id="21f7f-1277">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1277">Requirement</span></span>|<span data-ttu-id="21f7f-1278">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1279">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1280">1.3</span><span class="sxs-lookup"><span data-stu-id="21f7f-1280">1.3</span></span>|
|[<span data-ttu-id="21f7f-1281">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1282">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1282">ReadWriteItem</span></span>|
|[<span data-ttu-id="21f7f-1283">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1284">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="21f7f-1285">範例</span><span class="sxs-lookup"><span data-stu-id="21f7f-1285">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="21f7f-p182">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="21f7f-1288">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="21f7f-1288">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="21f7f-1289">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1289">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="21f7f-p183">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="21f7f-1293">Paramètres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1293">Parameters</span></span>

|<span data-ttu-id="21f7f-1294">Nom</span><span class="sxs-lookup"><span data-stu-id="21f7f-1294">Name</span></span>|<span data-ttu-id="21f7f-1295">Type</span><span class="sxs-lookup"><span data-stu-id="21f7f-1295">Type</span></span>|<span data-ttu-id="21f7f-1296">Attributs</span><span class="sxs-lookup"><span data-stu-id="21f7f-1296">Attributes</span></span>|<span data-ttu-id="21f7f-1297">Description</span><span class="sxs-lookup"><span data-stu-id="21f7f-1297">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="21f7f-1298">String</span><span class="sxs-lookup"><span data-stu-id="21f7f-1298">String</span></span>||<span data-ttu-id="21f7f-p184">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="21f7f-1302">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-1302">Object</span></span>|<span data-ttu-id="21f7f-1303">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1304">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1304">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="21f7f-1305">Objet</span><span class="sxs-lookup"><span data-stu-id="21f7f-1305">Object</span></span>|<span data-ttu-id="21f7f-1306">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1306">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-1307">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1307">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="21f7f-1308">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="21f7f-1308">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="21f7f-1309">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="21f7f-1309">&lt;optional&gt;</span></span>|<span data-ttu-id="21f7f-p185">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="21f7f-p186">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="21f7f-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="21f7f-1314">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="21f7f-1314">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="21f7f-1315">fonction</span><span class="sxs-lookup"><span data-stu-id="21f7f-1315">function</span></span>||<span data-ttu-id="21f7f-1316">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="21f7f-1316">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21f7f-1317">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="21f7f-1317">Requirements</span></span>

|<span data-ttu-id="21f7f-1318">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="21f7f-1318">Requirement</span></span>|<span data-ttu-id="21f7f-1319">Valeur</span><span class="sxs-lookup"><span data-stu-id="21f7f-1319">Value</span></span>|
|---|---|
|[<span data-ttu-id="21f7f-1320">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="21f7f-1320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="21f7f-1321">1.2</span><span class="sxs-lookup"><span data-stu-id="21f7f-1321">1.2</span></span>|
|[<span data-ttu-id="21f7f-1322">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="21f7f-1322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="21f7f-1323">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="21f7f-1323">ReadWriteItem</span></span>|
|[<span data-ttu-id="21f7f-1324">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="21f7f-1324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="21f7f-1325">Composition</span><span class="sxs-lookup"><span data-stu-id="21f7f-1325">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="21f7f-1326">Exemple</span><span class="sxs-lookup"><span data-stu-id="21f7f-1326">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
