---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 87ac31deaabda5fe1d2ca024972e326cfa168ea7
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068251"
---
# <a name="item"></a><span data-ttu-id="4d6f4-102">élément</span><span class="sxs-lookup"><span data-stu-id="4d6f4-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4d6f4-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4d6f4-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4d6f4-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-106">Requirements</span></span>

|<span data-ttu-id="4d6f4-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-107">Requirement</span></span>|<span data-ttu-id="4d6f4-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-110">1.0</span></span>|
|[<span data-ttu-id="4d6f4-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4d6f4-112">Restricted</span></span>|
|[<span data-ttu-id="4d6f4-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4d6f4-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4d6f4-115">Members and methods</span></span>

| <span data-ttu-id="4d6f4-116">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-116">Member</span></span> | <span data-ttu-id="4d6f4-117">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4d6f4-118">attachments</span><span class="sxs-lookup"><span data-stu-id="4d6f4-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="4d6f4-119">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-119">Member</span></span> |
| [<span data-ttu-id="4d6f4-120">bcc</span><span class="sxs-lookup"><span data-stu-id="4d6f4-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="4d6f4-121">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-121">Member</span></span> |
| [<span data-ttu-id="4d6f4-122">body</span><span class="sxs-lookup"><span data-stu-id="4d6f4-122">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="4d6f4-123">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-123">Member</span></span> |
| [<span data-ttu-id="4d6f4-124">cc</span><span class="sxs-lookup"><span data-stu-id="4d6f4-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="4d6f4-125">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-125">Member</span></span> |
| [<span data-ttu-id="4d6f4-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="4d6f4-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4d6f4-127">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-127">Member</span></span> |
| [<span data-ttu-id="4d6f4-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4d6f4-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4d6f4-129">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-129">Member</span></span> |
| [<span data-ttu-id="4d6f4-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4d6f4-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4d6f4-131">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-131">Member</span></span> |
| [<span data-ttu-id="4d6f4-132">end</span><span class="sxs-lookup"><span data-stu-id="4d6f4-132">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="4d6f4-133">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-133">Member</span></span> |
| [<span data-ttu-id="4d6f4-134">from</span><span class="sxs-lookup"><span data-stu-id="4d6f4-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="4d6f4-135">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-135">Member</span></span> |
| [<span data-ttu-id="4d6f4-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4d6f4-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4d6f4-137">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-137">Member</span></span> |
| [<span data-ttu-id="4d6f4-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="4d6f4-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4d6f4-139">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-139">Member</span></span> |
| [<span data-ttu-id="4d6f4-140">itemId</span><span class="sxs-lookup"><span data-stu-id="4d6f4-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4d6f4-141">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-141">Member</span></span> |
| [<span data-ttu-id="4d6f4-142">itemType</span><span class="sxs-lookup"><span data-stu-id="4d6f4-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="4d6f4-143">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-143">Member</span></span> |
| [<span data-ttu-id="4d6f4-144">location</span><span class="sxs-lookup"><span data-stu-id="4d6f4-144">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="4d6f4-145">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-145">Member</span></span> |
| [<span data-ttu-id="4d6f4-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4d6f4-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4d6f4-147">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-147">Member</span></span> |
| [<span data-ttu-id="4d6f4-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4d6f4-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="4d6f4-149">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-149">Member</span></span> |
| [<span data-ttu-id="4d6f4-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4d6f4-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="4d6f4-151">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-151">Member</span></span> |
| [<span data-ttu-id="4d6f4-152">organizer</span><span class="sxs-lookup"><span data-stu-id="4d6f4-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="4d6f4-153">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-153">Member</span></span> |
| [<span data-ttu-id="4d6f4-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="4d6f4-154">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="4d6f4-155">Member</span><span class="sxs-lookup"><span data-stu-id="4d6f4-155">Member</span></span> |
| [<span data-ttu-id="4d6f4-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4d6f4-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="4d6f4-157">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-157">Member</span></span> |
| [<span data-ttu-id="4d6f4-158">sender</span><span class="sxs-lookup"><span data-stu-id="4d6f4-158">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="4d6f4-159">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-159">Member</span></span> |
| [<span data-ttu-id="4d6f4-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="4d6f4-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="4d6f4-161">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-161">Member</span></span> |
| [<span data-ttu-id="4d6f4-162">start</span><span class="sxs-lookup"><span data-stu-id="4d6f4-162">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="4d6f4-163">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-163">Member</span></span> |
| [<span data-ttu-id="4d6f4-164">subject</span><span class="sxs-lookup"><span data-stu-id="4d6f4-164">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="4d6f4-165">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-165">Member</span></span> |
| [<span data-ttu-id="4d6f4-166">to</span><span class="sxs-lookup"><span data-stu-id="4d6f4-166">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="4d6f4-167">Membre</span><span class="sxs-lookup"><span data-stu-id="4d6f4-167">Member</span></span> |
| [<span data-ttu-id="4d6f4-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4d6f4-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4d6f4-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-169">Method</span></span> |
| [<span data-ttu-id="4d6f4-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4d6f4-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4d6f4-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-171">Method</span></span> |
| [<span data-ttu-id="4d6f4-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4d6f4-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4d6f4-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-173">Method</span></span> |
| [<span data-ttu-id="4d6f4-174">close</span><span class="sxs-lookup"><span data-stu-id="4d6f4-174">close</span></span>](#close) | <span data-ttu-id="4d6f4-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-175">Method</span></span> |
| [<span data-ttu-id="4d6f4-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4d6f4-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="4d6f4-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-177">Method</span></span> |
| [<span data-ttu-id="4d6f4-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4d6f4-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="4d6f4-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-179">Method</span></span> |
| [<span data-ttu-id="4d6f4-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="4d6f4-180">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="4d6f4-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-181">Method</span></span> |
| [<span data-ttu-id="4d6f4-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4d6f4-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="4d6f4-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-183">Method</span></span> |
| [<span data-ttu-id="4d6f4-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4d6f4-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="4d6f4-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-185">Method</span></span> |
| [<span data-ttu-id="4d6f4-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4d6f4-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4d6f4-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-187">Method</span></span> |
| [<span data-ttu-id="4d6f4-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4d6f4-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4d6f4-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-189">Method</span></span> |
| [<span data-ttu-id="4d6f4-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4d6f4-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4d6f4-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-191">Method</span></span> |
| [<span data-ttu-id="4d6f4-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="4d6f4-192">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="4d6f4-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-193">Method</span></span> |
| [<span data-ttu-id="4d6f4-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4d6f4-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="4d6f4-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-195">Method</span></span> |
| [<span data-ttu-id="4d6f4-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4d6f4-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4d6f4-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-197">Method</span></span> |
| [<span data-ttu-id="4d6f4-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4d6f4-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4d6f4-199">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-199">Method</span></span> |
| [<span data-ttu-id="4d6f4-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4d6f4-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4d6f4-201">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-201">Method</span></span> |
| [<span data-ttu-id="4d6f4-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4d6f4-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4d6f4-203">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-203">Method</span></span> |
| [<span data-ttu-id="4d6f4-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4d6f4-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4d6f4-205">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d6f4-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="4d6f4-206">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-206">Example</span></span>

<span data-ttu-id="4d6f4-207">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="4d6f4-208">Membres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="4d6f4-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4d6f4-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="4d6f4-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-212">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4d6f4-213">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-214">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-214">Type</span></span>

*   <span data-ttu-id="4d6f4-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4d6f4-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-216">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-216">Requirements</span></span>

|<span data-ttu-id="4d6f4-217">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-217">Requirement</span></span>|<span data-ttu-id="4d6f4-218">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-219">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-220">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-220">1.0</span></span>|
|[<span data-ttu-id="4d6f4-221">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-222">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-224">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-225">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-225">Example</span></span>

<span data-ttu-id="4d6f4-226">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="4d6f4-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="4d6f4-228">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4d6f4-229">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-230">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-230">Type</span></span>

*   [<span data-ttu-id="4d6f4-231">Destinataires</span><span class="sxs-lookup"><span data-stu-id="4d6f4-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="4d6f4-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-232">Requirements</span></span>

|<span data-ttu-id="4d6f4-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-233">Requirement</span></span>|<span data-ttu-id="4d6f4-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-236">1.1</span><span class="sxs-lookup"><span data-stu-id="4d6f4-236">1.1</span></span>|
|[<span data-ttu-id="4d6f4-237">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-238">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-240">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-241">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="4d6f4-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="4d6f4-243">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-244">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-244">Type</span></span>

*   [<span data-ttu-id="4d6f4-245">Corps</span><span class="sxs-lookup"><span data-stu-id="4d6f4-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="4d6f4-246">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-246">Requirements</span></span>

|<span data-ttu-id="4d6f4-247">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-247">Requirement</span></span>|<span data-ttu-id="4d6f4-248">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-249">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-250">1.1</span><span class="sxs-lookup"><span data-stu-id="4d6f4-250">1.1</span></span>|
|[<span data-ttu-id="4d6f4-251">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-252">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-253">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-254">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-255">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-255">Example</span></span>

<span data-ttu-id="4d6f4-256">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-256">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="4d6f4-257">Voici un exemple du paramètre result transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="4d6f4-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="4d6f4-259">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4d6f4-260">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-261">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-261">Read mode</span></span>

<span data-ttu-id="4d6f4-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-264">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-264">Compose mode</span></span>

<span data-ttu-id="4d6f4-265">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4d6f4-266">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-266">Type</span></span>

*   <span data-ttu-id="4d6f4-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-268">Requirements</span></span>

|<span data-ttu-id="4d6f4-269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-269">Requirement</span></span>|<span data-ttu-id="4d6f4-270">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-272">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-272">1.0</span></span>|
|[<span data-ttu-id="4d6f4-273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-274">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-276">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-276">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="4d6f4-277">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-277">(nullable) conversationId :String</span></span>

<span data-ttu-id="4d6f4-278">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4d6f4-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4d6f4-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-283">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-283">Type</span></span>

*   <span data-ttu-id="4d6f4-284">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4d6f4-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-285">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-285">Requirements</span></span>

|<span data-ttu-id="4d6f4-286">Requirement</span><span class="sxs-lookup"><span data-stu-id="4d6f4-286">Requirement</span></span>|<span data-ttu-id="4d6f4-287">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-288">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-289">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-289">1.0</span></span>|
|[<span data-ttu-id="4d6f4-290">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-291">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-293">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-294">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-294">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="4d6f4-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="4d6f4-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="4d6f4-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-298">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-298">Type</span></span>

*   <span data-ttu-id="4d6f4-299">Date</span><span class="sxs-lookup"><span data-stu-id="4d6f4-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-300">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-300">Requirements</span></span>

|<span data-ttu-id="4d6f4-301">Requirement</span><span class="sxs-lookup"><span data-stu-id="4d6f4-301">Requirement</span></span>|<span data-ttu-id="4d6f4-302">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-303">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-304">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-304">1.0</span></span>|
|[<span data-ttu-id="4d6f4-305">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-306">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-307">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-308">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-309">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-309">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="4d6f4-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="4d6f4-310">dateTimeModified :Date</span></span>

<span data-ttu-id="4d6f4-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-313">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-314">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-314">Type</span></span>

*   <span data-ttu-id="4d6f4-315">Date</span><span class="sxs-lookup"><span data-stu-id="4d6f4-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-316">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-316">Requirements</span></span>

|<span data-ttu-id="4d6f4-317">Requirement</span><span class="sxs-lookup"><span data-stu-id="4d6f4-317">Requirement</span></span>|<span data-ttu-id="4d6f4-318">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-319">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-320">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-320">1.0</span></span>|
|[<span data-ttu-id="4d6f4-321">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-322">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-323">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-324">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-325">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-325">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="4d6f4-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="4d6f4-327">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4d6f4-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-330">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-330">Read mode</span></span>

<span data-ttu-id="4d6f4-331">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-331">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-332">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-332">Compose mode</span></span>

<span data-ttu-id="4d6f4-333">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4d6f4-334">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-334">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4d6f4-335">L'exemple suivant définit l'heure de fin d'un rendez- [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) vous à l' `Time` aide de la méthode de l'objet.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4d6f4-336">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-336">Type</span></span>

*   <span data-ttu-id="4d6f4-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-338">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-338">Requirements</span></span>

|<span data-ttu-id="4d6f4-339">Requirement</span><span class="sxs-lookup"><span data-stu-id="4d6f4-339">Requirement</span></span>|<span data-ttu-id="4d6f4-340">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-341">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-342">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-342">1.0</span></span>|
|[<span data-ttu-id="4d6f4-343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-344">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-346">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-346">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="4d6f4-347">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-347">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="4d6f4-348">Permet d’obtenir l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="4d6f4-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-351">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-352">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-352">Read mode</span></span>

<span data-ttu-id="4d6f4-353">La propriété `from` renvoie un objet `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-354">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-354">Compose mode</span></span>

<span data-ttu-id="4d6f4-355">La propriété `from` renvoie un objet `From` qui fournit une méthode pour obtenir la valeur from.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4d6f4-356">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-356">Type</span></span>

*   <span data-ttu-id="4d6f4-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-358">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-358">Requirements</span></span>

|<span data-ttu-id="4d6f4-359">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4d6f4-360">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-361">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-361">1.0</span></span>|<span data-ttu-id="4d6f4-362">1.7</span><span class="sxs-lookup"><span data-stu-id="4d6f4-362">1.7</span></span>|
|[<span data-ttu-id="4d6f4-363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-364">ReadItem</span></span>|<span data-ttu-id="4d6f4-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="4d6f4-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-367">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-367">Read</span></span>|<span data-ttu-id="4d6f4-368">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-368">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="4d6f4-369">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-369">internetMessageId :String</span></span>

<span data-ttu-id="4d6f4-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-372">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-372">Type</span></span>

*   <span data-ttu-id="4d6f4-373">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4d6f4-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-374">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-374">Requirements</span></span>

|<span data-ttu-id="4d6f4-375">Requirement</span><span class="sxs-lookup"><span data-stu-id="4d6f4-375">Requirement</span></span>|<span data-ttu-id="4d6f4-376">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-377">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-378">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-378">1.0</span></span>|
|[<span data-ttu-id="4d6f4-379">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-380">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-381">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-382">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-383">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-383">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="4d6f4-384">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-384">itemClass :String</span></span>

<span data-ttu-id="4d6f4-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4d6f4-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="4d6f4-389">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-389">Type</span></span>|<span data-ttu-id="4d6f4-390">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-390">Description</span></span>|<span data-ttu-id="4d6f4-391">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="4d6f4-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="4d6f4-392">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="4d6f4-392">Appointment items</span></span>|<span data-ttu-id="4d6f4-393">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="4d6f4-394">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="4d6f4-394">Message items</span></span>|<span data-ttu-id="4d6f4-395">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="4d6f4-396">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-397">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-397">Type</span></span>

*   <span data-ttu-id="4d6f4-398">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4d6f4-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-399">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-399">Requirements</span></span>

|<span data-ttu-id="4d6f4-400">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-400">Requirement</span></span>|<span data-ttu-id="4d6f4-401">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-402">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-403">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-403">1.0</span></span>|
|[<span data-ttu-id="4d6f4-404">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-404">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-405">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-406">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-406">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-407">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-408">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-408">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4d6f4-409">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-409">(nullable) itemId :String</span></span>

<span data-ttu-id="4d6f4-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-412">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4d6f4-413">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4d6f4-414">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4d6f4-415">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4d6f4-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-418">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-418">Type</span></span>

*   <span data-ttu-id="4d6f4-419">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4d6f4-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-420">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-420">Requirements</span></span>

|<span data-ttu-id="4d6f4-421">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-421">Requirement</span></span>|<span data-ttu-id="4d6f4-422">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-423">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-424">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-424">1.0</span></span>|
|[<span data-ttu-id="4d6f4-425">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-426">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-427">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-428">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-429">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-429">Example</span></span>

<span data-ttu-id="4d6f4-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="4d6f4-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="4d6f4-433">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4d6f4-434">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-435">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-435">Type</span></span>

*   [<span data-ttu-id="4d6f4-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4d6f4-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="4d6f4-437">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-437">Requirements</span></span>

|<span data-ttu-id="4d6f4-438">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-438">Requirement</span></span>|<span data-ttu-id="4d6f4-439">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-440">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-441">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-441">1.0</span></span>|
|[<span data-ttu-id="4d6f4-442">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-442">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-443">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-444">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-444">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-445">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-446">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-446">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="4d6f4-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="4d6f4-448">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-449">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-449">Read mode</span></span>

<span data-ttu-id="4d6f4-450">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-451">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-451">Compose mode</span></span>

<span data-ttu-id="4d6f4-452">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4d6f4-453">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-453">Type</span></span>

*   <span data-ttu-id="4d6f4-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-455">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-455">Requirements</span></span>

|<span data-ttu-id="4d6f4-456">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-456">Requirement</span></span>|<span data-ttu-id="4d6f4-457">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-458">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-459">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-459">1.0</span></span>|
|[<span data-ttu-id="4d6f4-460">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-460">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-461">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-462">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-462">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-463">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-463">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4d6f4-464">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-464">normalizedSubject :String</span></span>

<span data-ttu-id="4d6f4-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4d6f4-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-469">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-469">Type</span></span>

*   <span data-ttu-id="4d6f4-470">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4d6f4-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-471">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-471">Requirements</span></span>

|<span data-ttu-id="4d6f4-472">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-472">Requirement</span></span>|<span data-ttu-id="4d6f4-473">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-474">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-475">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-475">1.0</span></span>|
|[<span data-ttu-id="4d6f4-476">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-477">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-478">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-479">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-480">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-480">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="4d6f4-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="4d6f4-482">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-483">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-483">Type</span></span>

*   [<span data-ttu-id="4d6f4-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4d6f4-484">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="4d6f4-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-485">Requirements</span></span>

|<span data-ttu-id="4d6f4-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-486">Requirement</span></span>|<span data-ttu-id="4d6f4-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-489">1.3</span><span class="sxs-lookup"><span data-stu-id="4d6f4-489">1.3</span></span>|
|[<span data-ttu-id="4d6f4-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-491">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-493">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-494">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-494">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="4d6f4-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="4d6f4-496">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4d6f4-497">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-498">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-498">Read mode</span></span>

<span data-ttu-id="4d6f4-499">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-500">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-500">Compose mode</span></span>

<span data-ttu-id="4d6f4-501">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4d6f4-502">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-502">Type</span></span>

*   <span data-ttu-id="4d6f4-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-504">Requirements</span></span>

|<span data-ttu-id="4d6f4-505">Requirement</span><span class="sxs-lookup"><span data-stu-id="4d6f4-505">Requirement</span></span>|<span data-ttu-id="4d6f4-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-507">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-508">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-508">1.0</span></span>|
|[<span data-ttu-id="4d6f4-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-510">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-512">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-512">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="4d6f4-513">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-513">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="4d6f4-514">Permet d’obtenir l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-515">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-515">Read mode</span></span>

<span data-ttu-id="4d6f4-516">La propriété `organizer` renvoie un objet [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-517">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-517">Compose mode</span></span>

<span data-ttu-id="4d6f4-518">La propriété `organizer` renvoie un objet [Organizer](/javascript/api/outlook_1_7/office.organizer) qui fournit une méthode pour obtenir la valeur organizer.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-518">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="4d6f4-519">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-519">Type</span></span>

*   <span data-ttu-id="4d6f4-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-521">Requirements</span></span>

|<span data-ttu-id="4d6f4-522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4d6f4-523">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-524">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-524">1.0</span></span>|<span data-ttu-id="4d6f4-525">1.7</span><span class="sxs-lookup"><span data-stu-id="4d6f4-525">1.7</span></span>|
|[<span data-ttu-id="4d6f4-526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-527">ReadItem</span></span>|<span data-ttu-id="4d6f4-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="4d6f4-529">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-530">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-530">Read</span></span>|<span data-ttu-id="4d6f4-531">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-531">Compose</span></span>|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="4d6f4-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="4d6f4-533">Permet d’obtenir ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="4d6f4-534">Permet d’obtenir la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="4d6f4-535">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="4d6f4-536">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="4d6f4-537">La propriété `recurrence` renvoie un objet [périodicité](/javascript/api/outlook_1_7/office.recurrence) pour des demandes de réunions ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="4d6f4-538">La valeur `null` est renvoyée pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="4d6f4-539">La valeur `undefined` est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="4d6f4-540">Remarque : les demandes de réunion ont une valeur `itemClass` d’IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="4d6f4-541">Remarque : si l’objet de périodicité est `null`, cela indique que l’objet est un rendez-vous unique ou une demande de réunion de rendez-vous unique, et NON une partie d’une série.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-542">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-542">Read mode</span></span>

<span data-ttu-id="4d6f4-543">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook_1_7/office.recurrence) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="4d6f4-544">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-544">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-545">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-545">Compose mode</span></span>

<span data-ttu-id="4d6f4-546">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook_1_7/office.recurrence) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="4d6f4-547">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4d6f4-548">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-548">Type</span></span>

* [<span data-ttu-id="4d6f4-549">Recurrence</span><span class="sxs-lookup"><span data-stu-id="4d6f4-549">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="4d6f4-550">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-550">Requirement</span></span>|<span data-ttu-id="4d6f4-551">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-552">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-553">1.7</span><span class="sxs-lookup"><span data-stu-id="4d6f4-553">1.7</span></span>|
|[<span data-ttu-id="4d6f4-554">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-554">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-555">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-556">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-556">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-557">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-557">Compose or Read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="4d6f4-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="4d6f4-559">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4d6f4-560">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-561">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-561">Read mode</span></span>

<span data-ttu-id="4d6f4-562">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-563">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-563">Compose mode</span></span>

<span data-ttu-id="4d6f4-564">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="4d6f4-565">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-565">Type</span></span>

*   <span data-ttu-id="4d6f4-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-567">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-567">Requirements</span></span>

|<span data-ttu-id="4d6f4-568">Requirement</span><span class="sxs-lookup"><span data-stu-id="4d6f4-568">Requirement</span></span>|<span data-ttu-id="4d6f4-569">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-570">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-571">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-571">1.0</span></span>|
|[<span data-ttu-id="4d6f4-572">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-573">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-574">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-575">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-575">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="4d6f4-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="4d6f4-p128">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4d6f4-p129">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p129">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-581">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-582">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-582">Type</span></span>

*   [<span data-ttu-id="4d6f4-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4d6f4-583">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4d6f4-584">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-584">Requirements</span></span>

|<span data-ttu-id="4d6f4-585">Requirement</span><span class="sxs-lookup"><span data-stu-id="4d6f4-585">Requirement</span></span>|<span data-ttu-id="4d6f4-586">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-587">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-588">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-588">1.0</span></span>|
|[<span data-ttu-id="4d6f4-589">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-590">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-591">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-592">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-593">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-593">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="4d6f4-594">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="4d6f4-595">Permet d’obtenir l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="4d6f4-596">Dans OWA et Outlook, `seriesId` renvoie l’identificateur de services web Exchange (EWS) de l’élément (series) parent auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="4d6f4-597">Dans iOS et Android, `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-598">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4d6f4-599">La propriété `seriesId` n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="4d6f4-600">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4d6f4-601">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="4d6f4-602">La propriété `seriesId` renvoie `null` pour les éléments qui n’ont pas d’élément parent, tels que des rendez-vous uniques, des éléments de séries ou des demandes de réunion, et renvoie `undefined` pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="4d6f4-603">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-603">Type</span></span>

* <span data-ttu-id="4d6f4-604">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4d6f4-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-605">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-605">Requirements</span></span>

|<span data-ttu-id="4d6f4-606">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-606">Requirement</span></span>|<span data-ttu-id="4d6f4-607">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-608">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-609">1.7</span><span class="sxs-lookup"><span data-stu-id="4d6f4-609">1.7</span></span>|
|[<span data-ttu-id="4d6f4-610">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-611">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-612">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-613">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-614">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-614">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="4d6f4-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="4d6f4-616">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4d6f4-p132">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-619">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-619">Read mode</span></span>

<span data-ttu-id="4d6f4-620">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-620">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-621">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-621">Compose mode</span></span>

<span data-ttu-id="4d6f4-622">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4d6f4-623">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-623">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4d6f4-624">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4d6f4-625">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-625">Type</span></span>

*   <span data-ttu-id="4d6f4-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-627">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-627">Requirements</span></span>

|<span data-ttu-id="4d6f4-628">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-628">Requirement</span></span>|<span data-ttu-id="4d6f4-629">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-630">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-631">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-631">1.0</span></span>|
|[<span data-ttu-id="4d6f4-632">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-632">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-633">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-634">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-634">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-635">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-635">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="4d6f4-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="4d6f4-637">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4d6f4-638">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-639">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-639">Read mode</span></span>

<span data-ttu-id="4d6f4-p133">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="4d6f4-642">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-643">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-643">Compose mode</span></span>

<span data-ttu-id="4d6f4-644">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="4d6f4-645">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-645">Type</span></span>

*   <span data-ttu-id="4d6f4-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-647">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-647">Requirements</span></span>

|<span data-ttu-id="4d6f4-648">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-648">Requirement</span></span>|<span data-ttu-id="4d6f4-649">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-650">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-651">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-651">1.0</span></span>|
|[<span data-ttu-id="4d6f4-652">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-653">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-654">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-655">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-655">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="4d6f4-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="4d6f4-657">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4d6f4-658">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4d6f4-659">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-659">Read mode</span></span>

<span data-ttu-id="4d6f4-p135">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="4d6f4-662">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-662">Compose mode</span></span>

<span data-ttu-id="4d6f4-663">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4d6f4-664">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-664">Type</span></span>

*   <span data-ttu-id="4d6f4-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-666">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-666">Requirements</span></span>

|<span data-ttu-id="4d6f4-667">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-667">Requirement</span></span>|<span data-ttu-id="4d6f4-668">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-669">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-670">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-670">1.0</span></span>|
|[<span data-ttu-id="4d6f4-671">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-672">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-673">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-674">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4d6f4-675">Méthodes</span><span class="sxs-lookup"><span data-stu-id="4d6f4-675">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4d6f4-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4d6f4-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4d6f4-677">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4d6f4-678">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4d6f4-679">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-680">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-680">Parameters</span></span>
|<span data-ttu-id="4d6f4-681">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-681">Name</span></span>|<span data-ttu-id="4d6f4-682">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-682">Type</span></span>|<span data-ttu-id="4d6f4-683">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-683">Attributes</span></span>|<span data-ttu-id="4d6f4-684">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="4d6f4-685">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-685">String</span></span>||<span data-ttu-id="4d6f4-p136">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4d6f4-688">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-688">String</span></span>||<span data-ttu-id="4d6f4-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4d6f4-691">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-691">Object</span></span>|<span data-ttu-id="4d6f4-692">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-692">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-693">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4d6f4-694">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-694">Object</span></span>|<span data-ttu-id="4d6f4-695">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-695">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-696">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="4d6f4-697">Boolean</span><span class="sxs-lookup"><span data-stu-id="4d6f4-697">Boolean</span></span>|<span data-ttu-id="4d6f4-698">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-698">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-699">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="4d6f4-700">fonction</span><span class="sxs-lookup"><span data-stu-id="4d6f4-700">function</span></span>|<span data-ttu-id="4d6f4-701">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-701">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-702">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4d6f4-703">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4d6f4-704">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4d6f4-705">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-705">Errors</span></span>

|<span data-ttu-id="4d6f4-706">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-706">Error code</span></span>|<span data-ttu-id="4d6f4-707">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="4d6f4-708">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="4d6f4-709">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4d6f4-710">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-711">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-711">Requirements</span></span>

|<span data-ttu-id="4d6f4-712">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-712">Requirement</span></span>|<span data-ttu-id="4d6f4-713">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-714">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-715">1.1</span><span class="sxs-lookup"><span data-stu-id="4d6f4-715">1.1</span></span>|
|[<span data-ttu-id="4d6f4-716">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-716">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="4d6f4-718">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-718">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-719">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4d6f4-720">Exemples</span><span class="sxs-lookup"><span data-stu-id="4d6f4-720">Examples</span></span>

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

<span data-ttu-id="4d6f4-721">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4d6f4-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4d6f4-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4d6f4-723">Ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4d6f4-724">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-725">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-725">Parameters</span></span>

| <span data-ttu-id="4d6f4-726">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-726">Name</span></span> | <span data-ttu-id="4d6f4-727">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-727">Type</span></span> | <span data-ttu-id="4d6f4-728">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-728">Attributes</span></span> | <span data-ttu-id="4d6f4-729">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4d6f4-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4d6f4-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4d6f4-731">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4d6f4-732">Fonction</span><span class="sxs-lookup"><span data-stu-id="4d6f4-732">Function</span></span> || <span data-ttu-id="4d6f4-p138">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4d6f4-736">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-736">Object</span></span> | <span data-ttu-id="4d6f4-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-737">&lt;optional&gt;</span></span> | <span data-ttu-id="4d6f4-738">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4d6f4-739">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-739">Object</span></span> | <span data-ttu-id="4d6f4-740">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-740">&lt;optional&gt;</span></span> | <span data-ttu-id="4d6f4-741">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4d6f4-742">fonction</span><span class="sxs-lookup"><span data-stu-id="4d6f4-742">function</span></span>| <span data-ttu-id="4d6f4-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-743">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-744">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-745">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-745">Requirements</span></span>

|<span data-ttu-id="4d6f4-746">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-746">Requirement</span></span>| <span data-ttu-id="4d6f4-747">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-748">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d6f4-749">1.7</span><span class="sxs-lookup"><span data-stu-id="4d6f4-749">1.7</span></span> |
|[<span data-ttu-id="4d6f4-750">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d6f4-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-751">ReadItem</span></span> |
|[<span data-ttu-id="4d6f4-752">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4d6f4-753">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="4d6f4-754">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-754">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4d6f4-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4d6f4-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4d6f4-756">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4d6f4-p139">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4d6f4-760">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4d6f4-761">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-761">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-762">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-762">Parameters</span></span>

|<span data-ttu-id="4d6f4-763">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-763">Name</span></span>|<span data-ttu-id="4d6f4-764">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-764">Type</span></span>|<span data-ttu-id="4d6f4-765">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-765">Attributes</span></span>|<span data-ttu-id="4d6f4-766">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="4d6f4-767">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-767">String</span></span>||<span data-ttu-id="4d6f4-p140">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4d6f4-770">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-770">String</span></span>||<span data-ttu-id="4d6f4-771">Objet de l'élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-771">The subject of the item to be attached.</span></span> <span data-ttu-id="4d6f4-772">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4d6f4-773">Object</span><span class="sxs-lookup"><span data-stu-id="4d6f4-773">Object</span></span>|<span data-ttu-id="4d6f4-774">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-774">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-775">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4d6f4-776">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-776">Object</span></span>|<span data-ttu-id="4d6f4-777">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-777">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-778">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4d6f4-779">fonction</span><span class="sxs-lookup"><span data-stu-id="4d6f4-779">function</span></span>|<span data-ttu-id="4d6f4-780">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-780">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-781">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4d6f4-782">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4d6f4-783">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4d6f4-784">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-784">Errors</span></span>

|<span data-ttu-id="4d6f4-785">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-785">Error code</span></span>|<span data-ttu-id="4d6f4-786">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4d6f4-787">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-788">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-788">Requirements</span></span>

|<span data-ttu-id="4d6f4-789">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-789">Requirement</span></span>|<span data-ttu-id="4d6f4-790">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-791">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-792">1.1</span><span class="sxs-lookup"><span data-stu-id="4d6f4-792">1.1</span></span>|
|[<span data-ttu-id="4d6f4-793">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-793">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="4d6f4-795">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-795">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-796">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-797">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-797">Example</span></span>

<span data-ttu-id="4d6f4-798">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="4d6f4-799">close()</span><span class="sxs-lookup"><span data-stu-id="4d6f4-799">close()</span></span>

<span data-ttu-id="4d6f4-800">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4d6f4-p142">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-803">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4d6f4-804">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-805">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-805">Requirements</span></span>

|<span data-ttu-id="4d6f4-806">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-806">Requirement</span></span>|<span data-ttu-id="4d6f4-807">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-808">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-809">1.3</span><span class="sxs-lookup"><span data-stu-id="4d6f4-809">1.3</span></span>|
|[<span data-ttu-id="4d6f4-810">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-811">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4d6f4-811">Restricted</span></span>|
|[<span data-ttu-id="4d6f4-812">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-813">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-813">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="4d6f4-814">displayReplyAllForm (Formulairemode, [callback])</span><span class="sxs-lookup"><span data-stu-id="4d6f4-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="4d6f4-815">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-816">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-816">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4d6f4-817">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-817">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4d6f4-818">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4d6f4-p143">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-822">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-822">Parameters</span></span>

|<span data-ttu-id="4d6f4-823">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-823">Name</span></span>|<span data-ttu-id="4d6f4-824">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-824">Type</span></span>|<span data-ttu-id="4d6f4-825">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-825">Attributes</span></span>|<span data-ttu-id="4d6f4-826">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4d6f4-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4d6f4-827">String &#124; Object</span></span>||<span data-ttu-id="4d6f4-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4d6f4-830">**OU**</span><span class="sxs-lookup"><span data-stu-id="4d6f4-830">**OR**</span></span><br/><span data-ttu-id="4d6f4-p145">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4d6f4-833">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-833">String</span></span>|<span data-ttu-id="4d6f4-834">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-834">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4d6f4-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4d6f4-838">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-838">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-839">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4d6f4-840">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-840">String</span></span>||<span data-ttu-id="4d6f4-p147">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4d6f4-843">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-843">String</span></span>||<span data-ttu-id="4d6f4-844">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4d6f4-845">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4d6f4-845">String</span></span>||<span data-ttu-id="4d6f4-p148">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4d6f4-848">Booléen</span><span class="sxs-lookup"><span data-stu-id="4d6f4-848">Boolean</span></span>||<span data-ttu-id="4d6f4-p149">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4d6f4-851">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-851">String</span></span>||<span data-ttu-id="4d6f4-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4d6f4-855">function</span><span class="sxs-lookup"><span data-stu-id="4d6f4-855">function</span></span>|<span data-ttu-id="4d6f4-856">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-856">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-857">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-858">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-858">Requirements</span></span>

|<span data-ttu-id="4d6f4-859">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-859">Requirement</span></span>|<span data-ttu-id="4d6f4-860">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-861">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-862">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-862">1.0</span></span>|
|[<span data-ttu-id="4d6f4-863">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-864">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-865">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-866">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4d6f4-867">Exemples</span><span class="sxs-lookup"><span data-stu-id="4d6f4-867">Examples</span></span>

<span data-ttu-id="4d6f4-868">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4d6f4-869">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-869">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4d6f4-870">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-870">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4d6f4-871">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4d6f4-872">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4d6f4-873">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="4d6f4-874">displayReplyForm (Formulairemode, [callback])</span><span class="sxs-lookup"><span data-stu-id="4d6f4-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="4d6f4-875">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-876">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-876">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4d6f4-877">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-877">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4d6f4-878">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4d6f4-p151">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-882">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-882">Parameters</span></span>

|<span data-ttu-id="4d6f4-883">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-883">Name</span></span>|<span data-ttu-id="4d6f4-884">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-884">Type</span></span>|<span data-ttu-id="4d6f4-885">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-885">Attributes</span></span>|<span data-ttu-id="4d6f4-886">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4d6f4-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4d6f4-887">String &#124; Object</span></span>||<span data-ttu-id="4d6f4-p152">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4d6f4-890">**OU**</span><span class="sxs-lookup"><span data-stu-id="4d6f4-890">**OR**</span></span><br/><span data-ttu-id="4d6f4-p153">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4d6f4-893">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-893">String</span></span>|<span data-ttu-id="4d6f4-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-894">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4d6f4-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4d6f4-898">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-898">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-899">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4d6f4-900">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-900">String</span></span>||<span data-ttu-id="4d6f4-p155">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4d6f4-903">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-903">String</span></span>||<span data-ttu-id="4d6f4-904">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4d6f4-905">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4d6f4-905">String</span></span>||<span data-ttu-id="4d6f4-p156">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4d6f4-908">Booléen</span><span class="sxs-lookup"><span data-stu-id="4d6f4-908">Boolean</span></span>||<span data-ttu-id="4d6f4-p157">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4d6f4-911">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-911">String</span></span>||<span data-ttu-id="4d6f4-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4d6f4-915">function</span><span class="sxs-lookup"><span data-stu-id="4d6f4-915">function</span></span>|<span data-ttu-id="4d6f4-916">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-916">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-917">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-918">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-918">Requirements</span></span>

|<span data-ttu-id="4d6f4-919">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-919">Requirement</span></span>|<span data-ttu-id="4d6f4-920">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-921">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-922">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-922">1.0</span></span>|
|[<span data-ttu-id="4d6f4-923">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-923">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-924">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-925">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-925">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-926">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4d6f4-927">Exemples</span><span class="sxs-lookup"><span data-stu-id="4d6f4-927">Examples</span></span>

<span data-ttu-id="4d6f4-928">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4d6f4-929">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-929">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4d6f4-930">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-930">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4d6f4-931">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4d6f4-932">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4d6f4-933">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="4d6f4-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4d6f4-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="4d6f4-935">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-936">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-936">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-937">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-937">Requirements</span></span>

|<span data-ttu-id="4d6f4-938">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-938">Requirement</span></span>|<span data-ttu-id="4d6f4-939">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-940">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-941">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-941">1.0</span></span>|
|[<span data-ttu-id="4d6f4-942">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-942">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-943">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-944">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-944">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-945">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d6f4-946">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-946">Returns:</span></span>

<span data-ttu-id="4d6f4-947">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-947">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4d6f4-948">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-948">Example</span></span>

<span data-ttu-id="4d6f4-949">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-949">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="4d6f4-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4d6f4-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4d6f4-951">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-952">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-952">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-953">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-953">Parameters</span></span>

|<span data-ttu-id="4d6f4-954">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-954">Name</span></span>|<span data-ttu-id="4d6f4-955">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-955">Type</span></span>|<span data-ttu-id="4d6f4-956">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="4d6f4-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4d6f4-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="4d6f4-958">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-959">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-959">Requirements</span></span>

|<span data-ttu-id="4d6f4-960">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-960">Requirement</span></span>|<span data-ttu-id="4d6f4-961">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-962">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-963">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-963">1.0</span></span>|
|[<span data-ttu-id="4d6f4-964">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-965">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4d6f4-965">Restricted</span></span>|
|[<span data-ttu-id="4d6f4-966">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-967">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d6f4-968">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-968">Returns:</span></span>

<span data-ttu-id="4d6f4-969">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4d6f4-970">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4d6f4-971">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4d6f4-972">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="4d6f4-973">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="4d6f4-973">Value of `entityType`</span></span>|<span data-ttu-id="4d6f4-974">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="4d6f4-974">Type of objects in returned array</span></span>|<span data-ttu-id="4d6f4-975">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="4d6f4-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="4d6f4-976">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-976">String</span></span>|<span data-ttu-id="4d6f4-977">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4d6f4-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="4d6f4-978">Contact</span><span class="sxs-lookup"><span data-stu-id="4d6f4-978">Contact</span></span>|<span data-ttu-id="4d6f4-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4d6f4-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="4d6f4-980">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-980">String</span></span>|<span data-ttu-id="4d6f4-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4d6f4-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="4d6f4-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4d6f4-982">MeetingSuggestion</span></span>|<span data-ttu-id="4d6f4-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4d6f4-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="4d6f4-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4d6f4-984">PhoneNumber</span></span>|<span data-ttu-id="4d6f4-985">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4d6f4-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="4d6f4-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4d6f4-986">TaskSuggestion</span></span>|<span data-ttu-id="4d6f4-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4d6f4-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="4d6f4-988">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-988">String</span></span>|<span data-ttu-id="4d6f4-989">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4d6f4-989">**Restricted**</span></span>|

<span data-ttu-id="4d6f4-990">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4d6f4-990">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="4d6f4-991">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-991">Example</span></span>

<span data-ttu-id="4d6f4-992">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="4d6f4-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4d6f4-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4d6f4-994">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-995">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-995">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4d6f4-996">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-997">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-997">Parameters</span></span>

|<span data-ttu-id="4d6f4-998">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-998">Name</span></span>|<span data-ttu-id="4d6f4-999">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-999">Type</span></span>|<span data-ttu-id="4d6f4-1000">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4d6f4-1001">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1001">String</span></span>|<span data-ttu-id="4d6f4-1002">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1003">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1003">Requirements</span></span>

|<span data-ttu-id="4d6f4-1004">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1004">Requirement</span></span>|<span data-ttu-id="4d6f4-1005">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1006">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1007">1.0</span></span>|
|[<span data-ttu-id="4d6f4-1008">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1009">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-1010">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1011">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d6f4-1012">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1012">Returns:</span></span>

<span data-ttu-id="4d6f4-p160">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4d6f4-1015">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4d6f4-1015">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="4d6f4-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4d6f4-1017">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-1018">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1018">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4d6f4-p161">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4d6f4-1022">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4d6f4-1023">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4d6f4-p162">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1027">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1027">Requirements</span></span>

|<span data-ttu-id="4d6f4-1028">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1028">Requirement</span></span>|<span data-ttu-id="4d6f4-1029">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1030">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1031">1.0</span></span>|
|[<span data-ttu-id="4d6f4-1032">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1032">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1033">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-1034">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1034">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1035">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d6f4-1036">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1036">Returns:</span></span>

<span data-ttu-id="4d6f4-p163">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="4d6f4-1039">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1039">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4d6f4-1040">Object</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1040">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4d6f4-1041">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1041">Example</span></span>

<span data-ttu-id="4d6f4-1042">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1042">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4d6f4-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4d6f4-1044">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1044">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-1045">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1045">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4d6f4-1046">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1046">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4d6f4-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-1049">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1049">Parameters</span></span>

|<span data-ttu-id="4d6f4-1050">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1050">Name</span></span>|<span data-ttu-id="4d6f4-1051">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1051">Type</span></span>|<span data-ttu-id="4d6f4-1052">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1052">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4d6f4-1053">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1053">String</span></span>|<span data-ttu-id="4d6f4-1054">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1054">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1055">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1055">Requirements</span></span>

|<span data-ttu-id="4d6f4-1056">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1056">Requirement</span></span>|<span data-ttu-id="4d6f4-1057">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1057">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1058">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1058">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1059">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1059">1.0</span></span>|
|[<span data-ttu-id="4d6f4-1060">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1060">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1061">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1061">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-1062">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1062">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1063">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1063">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d6f4-1064">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1064">Returns:</span></span>

<span data-ttu-id="4d6f4-1065">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1065">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="4d6f4-1066">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1066">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4d6f4-1067">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4d6f4-1067">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4d6f4-1068">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1068">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4d6f4-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4d6f4-1070">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1070">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4d6f4-p165">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-1073">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1073">Parameters</span></span>

|<span data-ttu-id="4d6f4-1074">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1074">Name</span></span>|<span data-ttu-id="4d6f4-1075">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1075">Type</span></span>|<span data-ttu-id="4d6f4-1076">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1076">Attributes</span></span>|<span data-ttu-id="4d6f4-1077">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1077">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="4d6f4-1078">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1078">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4d6f4-p166">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="4d6f4-1082">Object</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1082">Object</span></span>|<span data-ttu-id="4d6f4-1083">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1084">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4d6f4-1085">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1085">Object</span></span>|<span data-ttu-id="4d6f4-1086">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1087">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4d6f4-1088">fonction</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1088">function</span></span>||<span data-ttu-id="4d6f4-1089">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4d6f4-1090">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1090">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4d6f4-1091">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1091">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1092">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1092">Requirements</span></span>

|<span data-ttu-id="4d6f4-1093">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1093">Requirement</span></span>|<span data-ttu-id="4d6f4-1094">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1094">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1095">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1095">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1096">1.2</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1096">1.2</span></span>|
|[<span data-ttu-id="4d6f4-1097">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1097">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1098">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1098">ReadWriteItem</span></span>|
|[<span data-ttu-id="4d6f4-1099">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1099">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1100">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1100">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d6f4-1101">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1101">Returns:</span></span>

<span data-ttu-id="4d6f4-1102">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1102">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="4d6f4-1103">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1103">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4d6f4-1104">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1104">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4d6f4-1105">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1105">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="4d6f4-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="4d6f4-p168">Permet d’obtenir les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-1109">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1109">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1110">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1110">Requirements</span></span>

|<span data-ttu-id="4d6f4-1111">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1111">Requirement</span></span>|<span data-ttu-id="4d6f4-1112">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1113">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1114">1.6</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1114">1.6</span></span>|
|[<span data-ttu-id="4d6f4-1115">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1115">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1116">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1116">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-1117">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1117">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1118">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1118">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d6f4-1119">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1119">Returns:</span></span>

<span data-ttu-id="4d6f4-1120">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1120">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4d6f4-1121">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1121">Example</span></span>

<span data-ttu-id="4d6f4-1122">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1122">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="4d6f4-1123">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1123">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="4d6f4-p169">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-1126">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4d6f4-p170">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4d6f4-1130">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1130">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4d6f4-1131">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1131">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4d6f4-p171">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1135">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1135">Requirements</span></span>

|<span data-ttu-id="4d6f4-1136">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1136">Requirement</span></span>|<span data-ttu-id="4d6f4-1137">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1138">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1139">1.6</span></span>|
|[<span data-ttu-id="4d6f4-1140">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1140">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1141">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-1142">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1142">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1143">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d6f4-1144">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1144">Returns:</span></span>

<span data-ttu-id="4d6f4-p172">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="4d6f4-1147">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1147">Example</span></span>

<span data-ttu-id="4d6f4-1148">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1148">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4d6f4-1149">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1149">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4d6f4-1150">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1150">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4d6f4-p173">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-1154">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1154">Parameters</span></span>

|<span data-ttu-id="4d6f4-1155">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1155">Name</span></span>|<span data-ttu-id="4d6f4-1156">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1156">Type</span></span>|<span data-ttu-id="4d6f4-1157">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1157">Attributes</span></span>|<span data-ttu-id="4d6f4-1158">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1158">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="4d6f4-1159">function</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1159">function</span></span>||<span data-ttu-id="4d6f4-1160">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1160">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4d6f4-1161">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1161">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4d6f4-1162">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1162">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="4d6f4-1163">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1163">Object</span></span>|<span data-ttu-id="4d6f4-1164">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1164">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1165">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1165">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4d6f4-1166">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1166">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1167">Requirements</span></span>

|<span data-ttu-id="4d6f4-1168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1168">Requirement</span></span>|<span data-ttu-id="4d6f4-1169">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1171">1.0</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1171">1.0</span></span>|
|[<span data-ttu-id="4d6f4-1172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1173">ReadItem</span></span>|
|[<span data-ttu-id="4d6f4-1174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1175">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1175">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-1176">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1176">Example</span></span>

<span data-ttu-id="4d6f4-p176">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4d6f4-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4d6f4-1181">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1181">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4d6f4-p177">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-1186">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1186">Parameters</span></span>

|<span data-ttu-id="4d6f4-1187">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1187">Name</span></span>|<span data-ttu-id="4d6f4-1188">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1188">Type</span></span>|<span data-ttu-id="4d6f4-1189">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1189">Attributes</span></span>|<span data-ttu-id="4d6f4-1190">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1190">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="4d6f4-1191">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1191">String</span></span>||<span data-ttu-id="4d6f4-1192">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1192">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="4d6f4-1193">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1193">Object</span></span>|<span data-ttu-id="4d6f4-1194">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1195">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1195">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4d6f4-1196">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1196">Object</span></span>|<span data-ttu-id="4d6f4-1197">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1198">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1198">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4d6f4-1199">fonction</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1199">function</span></span>|<span data-ttu-id="4d6f4-1200">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1201">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1201">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4d6f4-1202">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1202">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4d6f4-1203">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1203">Errors</span></span>

|<span data-ttu-id="4d6f4-1204">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1204">Error code</span></span>|<span data-ttu-id="4d6f4-1205">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1205">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="4d6f4-1206">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1206">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1207">Requirements</span></span>

|<span data-ttu-id="4d6f4-1208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1208">Requirement</span></span>|<span data-ttu-id="4d6f4-1209">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1209">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1211">1.1</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1211">1.1</span></span>|
|[<span data-ttu-id="4d6f4-1212">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1213">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1213">ReadWriteItem</span></span>|
|[<span data-ttu-id="4d6f4-1214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1215">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1215">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-1216">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1216">Example</span></span>

<span data-ttu-id="4d6f4-1217">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1217">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4d6f4-1218">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1218">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4d6f4-1219">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1219">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4d6f4-1220">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1220">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-1221">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1221">Parameters</span></span>

| <span data-ttu-id="4d6f4-1222">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1222">Name</span></span> | <span data-ttu-id="4d6f4-1223">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1223">Type</span></span> | <span data-ttu-id="4d6f4-1224">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1224">Attributes</span></span> | <span data-ttu-id="4d6f4-1225">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1225">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4d6f4-1226">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1226">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4d6f4-1227">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1227">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="4d6f4-1228">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1228">Object</span></span> | <span data-ttu-id="4d6f4-1229">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="4d6f4-1230">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4d6f4-1231">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1231">Object</span></span> | <span data-ttu-id="4d6f4-1232">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="4d6f4-1233">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4d6f4-1234">fonction</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1234">function</span></span>| <span data-ttu-id="4d6f4-1235">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1236">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1237">Requirements</span></span>

|<span data-ttu-id="4d6f4-1238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1238">Requirement</span></span>| <span data-ttu-id="4d6f4-1239">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d6f4-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1241">1.7</span></span> |
|[<span data-ttu-id="4d6f4-1242">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d6f4-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1243">ReadItem</span></span> |
|[<span data-ttu-id="4d6f4-1244">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4d6f4-1245">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1245">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="4d6f4-1246">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1246">Example</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="4d6f4-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="4d6f4-1248">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="4d6f4-p178">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-1252">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4d6f4-1253">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4d6f4-p180">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6f4-1257">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4d6f4-1258">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="4d6f4-1259">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="4d6f4-1260">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-1261">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1261">Parameters</span></span>

|<span data-ttu-id="4d6f4-1262">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1262">Name</span></span>|<span data-ttu-id="4d6f4-1263">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1263">Type</span></span>|<span data-ttu-id="4d6f4-1264">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1264">Attributes</span></span>|<span data-ttu-id="4d6f4-1265">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4d6f4-1266">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1266">Object</span></span>|<span data-ttu-id="4d6f4-1267">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1268">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4d6f4-1269">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1269">Object</span></span>|<span data-ttu-id="4d6f4-1270">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1271">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4d6f4-1272">fonction</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1272">function</span></span>||<span data-ttu-id="4d6f4-1273">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4d6f4-1274">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1275">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1275">Requirements</span></span>

|<span data-ttu-id="4d6f4-1276">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1276">Requirement</span></span>|<span data-ttu-id="4d6f4-1277">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1278">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1279">1.3</span></span>|
|[<span data-ttu-id="4d6f4-1280">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="4d6f4-1282">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1283">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4d6f4-1284">範例</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1284">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="4d6f4-p182">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4d6f4-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4d6f4-1288">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4d6f4-p183">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d6f4-1292">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1292">Parameters</span></span>

|<span data-ttu-id="4d6f4-1293">Nom</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1293">Name</span></span>|<span data-ttu-id="4d6f4-1294">Type</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1294">Type</span></span>|<span data-ttu-id="4d6f4-1295">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1295">Attributes</span></span>|<span data-ttu-id="4d6f4-1296">Description</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="4d6f4-1297">String</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1297">String</span></span>||<span data-ttu-id="4d6f4-p184">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="4d6f4-1301">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1301">Object</span></span>|<span data-ttu-id="4d6f4-1302">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1303">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4d6f4-1304">Objet</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1304">Object</span></span>|<span data-ttu-id="4d6f4-1305">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-1306">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="4d6f4-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="4d6f4-1308">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="4d6f4-p185">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4d6f4-p186">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4d6f4-1313">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="4d6f4-1314">fonction</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1314">function</span></span>||<span data-ttu-id="4d6f4-1315">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d6f4-1316">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1316">Requirements</span></span>

|<span data-ttu-id="4d6f4-1317">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1317">Requirement</span></span>|<span data-ttu-id="4d6f4-1318">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d6f4-1319">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4d6f4-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1320">1.2</span></span>|
|[<span data-ttu-id="4d6f4-1321">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4d6f4-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="4d6f4-1323">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4d6f4-1324">Composition</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4d6f4-1325">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d6f4-1325">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
