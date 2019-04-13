---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,6
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: cc7897f791c5a07ed5c17a686b6601a1a7633f00
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838542"
---
# <a name="item"></a><span data-ttu-id="bae4c-102">élément</span><span class="sxs-lookup"><span data-stu-id="bae4c-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="bae4c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="bae4c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="bae4c-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="bae4c-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-106">Requirements</span></span>

|<span data-ttu-id="bae4c-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-107">Requirement</span></span>| <span data-ttu-id="bae4c-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-110">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-110">1.0</span></span>|
|[<span data-ttu-id="bae4c-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="bae4c-112">Restricted</span></span>|
|[<span data-ttu-id="bae4c-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bae4c-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="bae4c-115">Members and methods</span></span>

| <span data-ttu-id="bae4c-116">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-116">Member</span></span> | <span data-ttu-id="bae4c-117">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bae4c-118">attachments</span><span class="sxs-lookup"><span data-stu-id="bae4c-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="bae4c-119">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-119">Member</span></span> |
| [<span data-ttu-id="bae4c-120">bcc</span><span class="sxs-lookup"><span data-stu-id="bae4c-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="bae4c-121">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-121">Member</span></span> |
| [<span data-ttu-id="bae4c-122">body</span><span class="sxs-lookup"><span data-stu-id="bae4c-122">body</span></span>](#body-body) | <span data-ttu-id="bae4c-123">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-123">Member</span></span> |
| [<span data-ttu-id="bae4c-124">cc</span><span class="sxs-lookup"><span data-stu-id="bae4c-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bae4c-125">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-125">Member</span></span> |
| [<span data-ttu-id="bae4c-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="bae4c-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="bae4c-127">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-127">Member</span></span> |
| [<span data-ttu-id="bae4c-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="bae4c-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="bae4c-129">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-129">Member</span></span> |
| [<span data-ttu-id="bae4c-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="bae4c-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="bae4c-131">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-131">Member</span></span> |
| [<span data-ttu-id="bae4c-132">end</span><span class="sxs-lookup"><span data-stu-id="bae4c-132">end</span></span>](#end-datetime) | <span data-ttu-id="bae4c-133">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-133">Member</span></span> |
| [<span data-ttu-id="bae4c-134">from</span><span class="sxs-lookup"><span data-stu-id="bae4c-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="bae4c-135">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-135">Member</span></span> |
| [<span data-ttu-id="bae4c-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="bae4c-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="bae4c-137">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-137">Member</span></span> |
| [<span data-ttu-id="bae4c-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="bae4c-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="bae4c-139">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-139">Member</span></span> |
| [<span data-ttu-id="bae4c-140">itemId</span><span class="sxs-lookup"><span data-stu-id="bae4c-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="bae4c-141">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-141">Member</span></span> |
| [<span data-ttu-id="bae4c-142">itemType</span><span class="sxs-lookup"><span data-stu-id="bae4c-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="bae4c-143">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-143">Member</span></span> |
| [<span data-ttu-id="bae4c-144">location</span><span class="sxs-lookup"><span data-stu-id="bae4c-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="bae4c-145">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-145">Member</span></span> |
| [<span data-ttu-id="bae4c-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="bae4c-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="bae4c-147">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-147">Member</span></span> |
| [<span data-ttu-id="bae4c-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="bae4c-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="bae4c-149">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-149">Member</span></span> |
| [<span data-ttu-id="bae4c-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="bae4c-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bae4c-151">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-151">Member</span></span> |
| [<span data-ttu-id="bae4c-152">organizer</span><span class="sxs-lookup"><span data-stu-id="bae4c-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="bae4c-153">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-153">Member</span></span> |
| [<span data-ttu-id="bae4c-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="bae4c-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bae4c-155">Member</span><span class="sxs-lookup"><span data-stu-id="bae4c-155">Member</span></span> |
| [<span data-ttu-id="bae4c-156">sender</span><span class="sxs-lookup"><span data-stu-id="bae4c-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="bae4c-157">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-157">Member</span></span> |
| [<span data-ttu-id="bae4c-158">start</span><span class="sxs-lookup"><span data-stu-id="bae4c-158">start</span></span>](#start-datetime) | <span data-ttu-id="bae4c-159">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-159">Member</span></span> |
| [<span data-ttu-id="bae4c-160">subject</span><span class="sxs-lookup"><span data-stu-id="bae4c-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="bae4c-161">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-161">Member</span></span> |
| [<span data-ttu-id="bae4c-162">to</span><span class="sxs-lookup"><span data-stu-id="bae4c-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bae4c-163">Membre</span><span class="sxs-lookup"><span data-stu-id="bae4c-163">Member</span></span> |
| [<span data-ttu-id="bae4c-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bae4c-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="bae4c-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-165">Method</span></span> |
| [<span data-ttu-id="bae4c-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bae4c-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="bae4c-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-167">Method</span></span> |
| [<span data-ttu-id="bae4c-168">close</span><span class="sxs-lookup"><span data-stu-id="bae4c-168">close</span></span>](#close) | <span data-ttu-id="bae4c-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-169">Method</span></span> |
| [<span data-ttu-id="bae4c-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="bae4c-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="bae4c-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-171">Method</span></span> |
| [<span data-ttu-id="bae4c-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="bae4c-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="bae4c-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-173">Method</span></span> |
| [<span data-ttu-id="bae4c-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="bae4c-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="bae4c-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-175">Method</span></span> |
| [<span data-ttu-id="bae4c-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="bae4c-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="bae4c-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-177">Method</span></span> |
| [<span data-ttu-id="bae4c-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="bae4c-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="bae4c-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-179">Method</span></span> |
| [<span data-ttu-id="bae4c-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="bae4c-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="bae4c-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-181">Method</span></span> |
| [<span data-ttu-id="bae4c-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="bae4c-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="bae4c-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-183">Method</span></span> |
| [<span data-ttu-id="bae4c-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="bae4c-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="bae4c-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-185">Method</span></span> |
| [<span data-ttu-id="bae4c-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="bae4c-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="bae4c-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-187">Method</span></span> |
| [<span data-ttu-id="bae4c-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="bae4c-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="bae4c-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-189">Method</span></span> |
| [<span data-ttu-id="bae4c-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="bae4c-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="bae4c-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-191">Method</span></span> |
| [<span data-ttu-id="bae4c-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bae4c-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="bae4c-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-193">Method</span></span> |
| [<span data-ttu-id="bae4c-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="bae4c-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="bae4c-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-195">Method</span></span> |
| [<span data-ttu-id="bae4c-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="bae4c-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="bae4c-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="bae4c-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="bae4c-198">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-198">Example</span></span>

<span data-ttu-id="bae4c-199">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="bae4c-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="bae4c-200">Membres</span><span class="sxs-lookup"><span data-stu-id="bae4c-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="bae4c-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="bae4c-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="bae4c-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-204">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="bae4c-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="bae4c-205">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="bae4c-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-206">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-206">Type</span></span>

*   <span data-ttu-id="bae4c-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="bae4c-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-208">Requirements</span></span>

|<span data-ttu-id="bae4c-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-209">Requirement</span></span>| <span data-ttu-id="bae4c-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-212">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-212">1.0</span></span>|
|[<span data-ttu-id="bae4c-213">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-214">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-215">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-216">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-217">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-217">Example</span></span>

<span data-ttu-id="bae4c-218">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="bae4c-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bae4c-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bae4c-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bae4c-220">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="bae4c-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="bae4c-221">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-222">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-222">Type</span></span>

*   [<span data-ttu-id="bae4c-223">Destinataires</span><span class="sxs-lookup"><span data-stu-id="bae4c-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="bae4c-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-224">Requirements</span></span>

|<span data-ttu-id="bae4c-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-225">Requirement</span></span>| <span data-ttu-id="bae4c-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-228">1.1</span><span class="sxs-lookup"><span data-stu-id="bae4c-228">1.1</span></span>|
|[<span data-ttu-id="bae4c-229">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-230">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-231">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-232">Composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-233">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="bae4c-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="bae4c-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="bae4c-235">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-236">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-236">Type</span></span>

*   [<span data-ttu-id="bae4c-237">Body</span><span class="sxs-lookup"><span data-stu-id="bae4c-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="bae4c-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-238">Requirements</span></span>

|<span data-ttu-id="bae4c-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-239">Requirement</span></span>| <span data-ttu-id="bae4c-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-242">1.1</span><span class="sxs-lookup"><span data-stu-id="bae4c-242">1.1</span></span>|
|[<span data-ttu-id="bae4c-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-244">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-247">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-247">Example</span></span>

<span data-ttu-id="bae4c-248">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="bae4c-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="bae4c-249">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bae4c-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bae4c-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bae4c-251">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="bae4c-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="bae4c-252">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="bae4c-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bae4c-253">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-253">Read mode</span></span>

<span data-ttu-id="bae4c-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="bae4c-256">Mode composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-256">Compose mode</span></span>

<span data-ttu-id="bae4c-257">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="bae4c-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bae4c-258">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-258">Type</span></span>

*   <span data-ttu-id="bae4c-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bae4c-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-260">Requirements</span></span>

|<span data-ttu-id="bae4c-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-261">Requirement</span></span>| <span data-ttu-id="bae4c-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-264">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-264">1.0</span></span>|
|[<span data-ttu-id="bae4c-265">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-266">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-267">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-268">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-268">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="bae4c-269">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="bae4c-269">(nullable) conversationId :String</span></span>

<span data-ttu-id="bae4c-270">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="bae4c-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="bae4c-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="bae4c-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-275">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-275">Type</span></span>

*   <span data-ttu-id="bae4c-276">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-277">Requirements</span></span>

|<span data-ttu-id="bae4c-278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-278">Requirement</span></span>| <span data-ttu-id="bae4c-279">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-281">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-281">1.0</span></span>|
|[<span data-ttu-id="bae4c-282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-283">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-285">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-286">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="bae4c-287">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="bae4c-287">dateTimeCreated :Date</span></span>

<span data-ttu-id="bae4c-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-290">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-290">Type</span></span>

*   <span data-ttu-id="bae4c-291">Date</span><span class="sxs-lookup"><span data-stu-id="bae4c-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-292">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-292">Requirements</span></span>

|<span data-ttu-id="bae4c-293">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-293">Requirement</span></span>| <span data-ttu-id="bae4c-294">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-295">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-296">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-296">1.0</span></span>|
|[<span data-ttu-id="bae4c-297">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-298">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-299">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-300">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-301">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="bae4c-302">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="bae4c-302">dateTimeModified :Date</span></span>

<span data-ttu-id="bae4c-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-305">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-305">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-306">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-306">Type</span></span>

*   <span data-ttu-id="bae4c-307">Date</span><span class="sxs-lookup"><span data-stu-id="bae4c-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-308">Requirements</span></span>

|<span data-ttu-id="bae4c-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-309">Requirement</span></span>| <span data-ttu-id="bae4c-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-312">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-312">1.0</span></span>|
|[<span data-ttu-id="bae4c-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-314">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-316">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-317">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="bae4c-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="bae4c-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="bae4c-319">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bae4c-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="bae4c-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bae4c-322">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-322">Read mode</span></span>

<span data-ttu-id="bae4c-323">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="bae4c-324">Mode composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-324">Compose mode</span></span>

<span data-ttu-id="bae4c-325">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="bae4c-326">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="bae4c-326">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="bae4c-327">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="bae4c-328">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-328">Type</span></span>

*   <span data-ttu-id="bae4c-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="bae4c-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-330">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-330">Requirements</span></span>

|<span data-ttu-id="bae4c-331">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-331">Requirement</span></span>| <span data-ttu-id="bae4c-332">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-333">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-334">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-334">1.0</span></span>|
|[<span data-ttu-id="bae4c-335">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-336">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-337">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-338">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="bae4c-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="bae4c-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="bae4c-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="bae4c-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-344">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-345">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-345">Type</span></span>

*   [<span data-ttu-id="bae4c-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bae4c-346">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="example"></a><span data-ttu-id="bae4c-347">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="bae4c-348">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-348">Requirements</span></span>

|<span data-ttu-id="bae4c-349">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-349">Requirement</span></span>| <span data-ttu-id="bae4c-350">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-351">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-352">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-352">1.0</span></span>|
|[<span data-ttu-id="bae4c-353">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-354">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-355">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-356">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="bae4c-357">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="bae4c-357">internetMessageId :String</span></span>

<span data-ttu-id="bae4c-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-360">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-360">Type</span></span>

*   <span data-ttu-id="bae4c-361">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-362">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-362">Requirements</span></span>

|<span data-ttu-id="bae4c-363">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-363">Requirement</span></span>| <span data-ttu-id="bae4c-364">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-365">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-366">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-366">1.0</span></span>|
|[<span data-ttu-id="bae4c-367">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-368">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-369">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-370">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-371">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="bae4c-372">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="bae4c-372">itemClass :String</span></span>

<span data-ttu-id="bae4c-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="bae4c-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="bae4c-377">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-377">Type</span></span> | <span data-ttu-id="bae4c-378">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-378">Description</span></span> | <span data-ttu-id="bae4c-379">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="bae4c-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="bae4c-380">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="bae4c-380">Appointment items</span></span> | <span data-ttu-id="bae4c-381">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="bae4c-382">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="bae4c-382">Message items</span></span> | <span data-ttu-id="bae4c-383">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="bae4c-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="bae4c-384">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-385">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-385">Type</span></span>

*   <span data-ttu-id="bae4c-386">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-387">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-387">Requirements</span></span>

|<span data-ttu-id="bae4c-388">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-388">Requirement</span></span>| <span data-ttu-id="bae4c-389">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-390">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-391">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-391">1.0</span></span>|
|[<span data-ttu-id="bae4c-392">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-393">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-394">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-395">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-396">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="bae4c-397">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="bae4c-397">(nullable) itemId :String</span></span>

<span data-ttu-id="bae4c-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-400">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="bae4c-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="bae4c-401">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="bae4c-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="bae4c-402">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="bae4c-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="bae4c-403">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="bae4c-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="bae4c-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-406">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-406">Type</span></span>

*   <span data-ttu-id="bae4c-407">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-408">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-408">Requirements</span></span>

|<span data-ttu-id="bae4c-409">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-409">Requirement</span></span>| <span data-ttu-id="bae4c-410">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-411">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-412">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-412">1.0</span></span>|
|[<span data-ttu-id="bae4c-413">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-414">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-415">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-416">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-417">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-417">Example</span></span>

<span data-ttu-id="bae4c-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="bae4c-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="bae4c-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="bae4c-421">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="bae4c-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="bae4c-422">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bae4c-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-423">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-423">Type</span></span>

*   [<span data-ttu-id="bae4c-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="bae4c-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="bae4c-425">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-425">Requirements</span></span>

|<span data-ttu-id="bae4c-426">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-426">Requirement</span></span>| <span data-ttu-id="bae4c-427">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-428">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-429">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-429">1.0</span></span>|
|[<span data-ttu-id="bae4c-430">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-431">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-432">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-433">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-434">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="bae4c-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="bae4c-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="bae4c-436">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bae4c-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bae4c-437">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-437">Read mode</span></span>

<span data-ttu-id="bae4c-438">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bae4c-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="bae4c-439">Mode composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-439">Compose mode</span></span>

<span data-ttu-id="bae4c-440">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bae4c-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bae4c-441">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-441">Type</span></span>

*   <span data-ttu-id="bae4c-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="bae4c-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-443">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-443">Requirements</span></span>

|<span data-ttu-id="bae4c-444">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-444">Requirement</span></span>| <span data-ttu-id="bae4c-445">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-446">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-447">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-447">1.0</span></span>|
|[<span data-ttu-id="bae4c-448">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-449">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-450">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-451">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="bae4c-452">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="bae4c-452">normalizedSubject :String</span></span>

<span data-ttu-id="bae4c-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="bae4c-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="bae4c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-457">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-457">Type</span></span>

*   <span data-ttu-id="bae4c-458">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-459">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-459">Requirements</span></span>

|<span data-ttu-id="bae4c-460">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-460">Requirement</span></span>| <span data-ttu-id="bae4c-461">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-462">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-463">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-463">1.0</span></span>|
|[<span data-ttu-id="bae4c-464">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-465">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-466">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-467">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-468">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="bae4c-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="bae4c-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="bae4c-470">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-471">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-471">Type</span></span>

*   [<span data-ttu-id="bae4c-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="bae4c-472">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="bae4c-473">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-473">Requirements</span></span>

|<span data-ttu-id="bae4c-474">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-474">Requirement</span></span>| <span data-ttu-id="bae4c-475">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-476">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-477">1.3</span><span class="sxs-lookup"><span data-stu-id="bae4c-477">1.3</span></span>|
|[<span data-ttu-id="bae4c-478">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-479">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-480">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-481">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-482">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bae4c-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bae4c-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bae4c-484">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="bae4c-485">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="bae4c-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bae4c-486">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-486">Read mode</span></span>

<span data-ttu-id="bae4c-487">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="bae4c-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="bae4c-488">Mode composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-488">Compose mode</span></span>

<span data-ttu-id="bae4c-489">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="bae4c-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bae4c-490">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-490">Type</span></span>

*   <span data-ttu-id="bae4c-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bae4c-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-492">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-492">Requirements</span></span>

|<span data-ttu-id="bae4c-493">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-493">Requirement</span></span>| <span data-ttu-id="bae4c-494">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-495">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-496">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-496">1.0</span></span>|
|[<span data-ttu-id="bae4c-497">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-498">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-499">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-500">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="bae4c-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="bae4c-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="bae4c-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-504">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-504">Type</span></span>

*   [<span data-ttu-id="bae4c-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bae4c-505">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="bae4c-506">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-506">Requirements</span></span>

|<span data-ttu-id="bae4c-507">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-507">Requirement</span></span>| <span data-ttu-id="bae4c-508">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-509">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-510">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-510">1.0</span></span>|
|[<span data-ttu-id="bae4c-511">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-512">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-513">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-514">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-515">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bae4c-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bae4c-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bae4c-517">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="bae4c-518">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="bae4c-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bae4c-519">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-519">Read mode</span></span>

<span data-ttu-id="bae4c-520">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="bae4c-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="bae4c-521">Mode composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-521">Compose mode</span></span>

<span data-ttu-id="bae4c-522">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="bae4c-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="bae4c-523">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-523">Type</span></span>

*   <span data-ttu-id="bae4c-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bae4c-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-525">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-525">Requirements</span></span>

|<span data-ttu-id="bae4c-526">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-526">Requirement</span></span>| <span data-ttu-id="bae4c-527">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-528">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-529">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-529">1.0</span></span>|
|[<span data-ttu-id="bae4c-530">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-531">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-532">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-533">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="bae4c-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="bae4c-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="bae4c-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="bae4c-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-539">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="bae4c-540">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-540">Type</span></span>

*   [<span data-ttu-id="bae4c-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bae4c-541">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="bae4c-542">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-542">Requirements</span></span>

|<span data-ttu-id="bae4c-543">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-543">Requirement</span></span>| <span data-ttu-id="bae4c-544">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-545">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-546">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-546">1.0</span></span>|
|[<span data-ttu-id="bae4c-547">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-548">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-549">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-550">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-551">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="bae4c-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="bae4c-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="bae4c-553">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bae4c-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="bae4c-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bae4c-556">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-556">Read mode</span></span>

<span data-ttu-id="bae4c-557">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="bae4c-558">Mode composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-558">Compose mode</span></span>

<span data-ttu-id="bae4c-559">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="bae4c-560">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="bae4c-560">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="bae4c-561">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="bae4c-562">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-562">Type</span></span>

*   <span data-ttu-id="bae4c-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="bae4c-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-564">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-564">Requirements</span></span>

|<span data-ttu-id="bae4c-565">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-565">Requirement</span></span>| <span data-ttu-id="bae4c-566">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-567">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-568">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-568">1.0</span></span>|
|[<span data-ttu-id="bae4c-569">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-570">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-571">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-572">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-572">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="bae4c-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="bae4c-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="bae4c-574">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="bae4c-575">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="bae4c-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bae4c-576">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-576">Read mode</span></span>

<span data-ttu-id="bae4c-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="bae4c-579">Mode composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-579">Compose mode</span></span>

<span data-ttu-id="bae4c-580">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="bae4c-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="bae4c-581">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-581">Type</span></span>

*   <span data-ttu-id="bae4c-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="bae4c-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-583">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-583">Requirements</span></span>

|<span data-ttu-id="bae4c-584">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-584">Requirement</span></span>| <span data-ttu-id="bae4c-585">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-586">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-587">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-587">1.0</span></span>|
|[<span data-ttu-id="bae4c-588">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-589">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-590">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-591">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-591">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bae4c-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bae4c-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bae4c-593">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="bae4c-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="bae4c-594">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="bae4c-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bae4c-595">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-595">Read mode</span></span>

<span data-ttu-id="bae4c-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="bae4c-598">Mode composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-598">Compose mode</span></span>

<span data-ttu-id="bae4c-599">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="bae4c-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bae4c-600">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-600">Type</span></span>

*   <span data-ttu-id="bae4c-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bae4c-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-602">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-602">Requirements</span></span>

|<span data-ttu-id="bae4c-603">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-603">Requirement</span></span>| <span data-ttu-id="bae4c-604">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-605">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-606">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-606">1.0</span></span>|
|[<span data-ttu-id="bae4c-607">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-608">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-609">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-610">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="bae4c-611">Méthodes</span><span class="sxs-lookup"><span data-stu-id="bae4c-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="bae4c-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bae4c-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bae4c-613">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="bae4c-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="bae4c-614">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="bae4c-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="bae4c-615">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="bae4c-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-616">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-616">Parameters</span></span>

|<span data-ttu-id="bae4c-617">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-617">Name</span></span>| <span data-ttu-id="bae4c-618">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-618">Type</span></span>| <span data-ttu-id="bae4c-619">Attributs</span><span class="sxs-lookup"><span data-stu-id="bae4c-619">Attributes</span></span>| <span data-ttu-id="bae4c-620">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="bae4c-621">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-621">String</span></span>||<span data-ttu-id="bae4c-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="bae4c-624">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-624">String</span></span>||<span data-ttu-id="bae4c-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="bae4c-627">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-627">Object</span></span>| <span data-ttu-id="bae4c-628">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-628">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-629">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="bae4c-630">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-630">Object</span></span> | <span data-ttu-id="bae4c-631">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-631">&lt;optional&gt;</span></span> | <span data-ttu-id="bae4c-632">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="bae4c-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="bae4c-633">Boolean</span></span> | <span data-ttu-id="bae4c-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-634">&lt;optional&gt;</span></span> | <span data-ttu-id="bae4c-635">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="bae4c-636">fonction</span><span class="sxs-lookup"><span data-stu-id="bae4c-636">function</span></span>| <span data-ttu-id="bae4c-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-637">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-638">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bae4c-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bae4c-639">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bae4c-640">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="bae4c-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bae4c-641">Erreurs</span><span class="sxs-lookup"><span data-stu-id="bae4c-641">Errors</span></span>

| <span data-ttu-id="bae4c-642">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="bae4c-642">Error code</span></span> | <span data-ttu-id="bae4c-643">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="bae4c-644">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="bae4c-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="bae4c-645">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="bae4c-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="bae4c-646">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bae4c-647">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-647">Requirements</span></span>

|<span data-ttu-id="bae4c-648">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-648">Requirement</span></span>| <span data-ttu-id="bae4c-649">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-650">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-651">1.1</span><span class="sxs-lookup"><span data-stu-id="bae4c-651">1.1</span></span>|
|[<span data-ttu-id="bae4c-652">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="bae4c-654">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-655">Composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="bae4c-656">Exemples</span><span class="sxs-lookup"><span data-stu-id="bae4c-656">Examples</span></span>

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

<span data-ttu-id="bae4c-657">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="bae4c-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="bae4c-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bae4c-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bae4c-659">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bae4c-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="bae4c-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="bae4c-663">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="bae4c-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="bae4c-664">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="bae4c-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-665">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-665">Parameters</span></span>

|<span data-ttu-id="bae4c-666">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-666">Name</span></span>| <span data-ttu-id="bae4c-667">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-667">Type</span></span>| <span data-ttu-id="bae4c-668">Attributs</span><span class="sxs-lookup"><span data-stu-id="bae4c-668">Attributes</span></span>| <span data-ttu-id="bae4c-669">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="bae4c-670">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-670">String</span></span>||<span data-ttu-id="bae4c-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="bae4c-673">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-673">String</span></span>||<span data-ttu-id="bae4c-674">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="bae4c-674">The subject of the item to be attached.</span></span> <span data-ttu-id="bae4c-675">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="bae4c-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="bae4c-676">Object</span><span class="sxs-lookup"><span data-stu-id="bae4c-676">Object</span></span>| <span data-ttu-id="bae4c-677">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-677">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-678">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bae4c-679">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-679">Object</span></span>| <span data-ttu-id="bae4c-680">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-680">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-681">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bae4c-682">fonction</span><span class="sxs-lookup"><span data-stu-id="bae4c-682">function</span></span>| <span data-ttu-id="bae4c-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-683">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-684">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bae4c-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bae4c-685">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bae4c-686">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="bae4c-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bae4c-687">Erreurs</span><span class="sxs-lookup"><span data-stu-id="bae4c-687">Errors</span></span>

| <span data-ttu-id="bae4c-688">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="bae4c-688">Error code</span></span> | <span data-ttu-id="bae4c-689">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="bae4c-690">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bae4c-691">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-691">Requirements</span></span>

|<span data-ttu-id="bae4c-692">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-692">Requirement</span></span>| <span data-ttu-id="bae4c-693">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-694">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-695">1.1</span><span class="sxs-lookup"><span data-stu-id="bae4c-695">1.1</span></span>|
|[<span data-ttu-id="bae4c-696">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="bae4c-698">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-699">Composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-700">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-700">Example</span></span>

<span data-ttu-id="bae4c-701">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="bae4c-702">close()</span><span class="sxs-lookup"><span data-stu-id="bae4c-702">close()</span></span>

<span data-ttu-id="bae4c-703">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="bae4c-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="bae4c-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-706">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="bae4c-707">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="bae4c-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-708">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-708">Requirements</span></span>

|<span data-ttu-id="bae4c-709">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-709">Requirement</span></span>| <span data-ttu-id="bae4c-710">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-711">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-712">1.3</span><span class="sxs-lookup"><span data-stu-id="bae4c-712">1.3</span></span>|
|[<span data-ttu-id="bae4c-713">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-714">Restreinte</span><span class="sxs-lookup"><span data-stu-id="bae4c-714">Restricted</span></span>|
|[<span data-ttu-id="bae4c-715">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-716">Composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="bae4c-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="bae4c-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="bae4c-718">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="bae4c-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-719">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bae4c-720">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bae4c-721">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="bae4c-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="bae4c-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-725">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-725">Parameters</span></span>

| <span data-ttu-id="bae4c-726">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-726">Name</span></span> | <span data-ttu-id="bae4c-727">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-727">Type</span></span> | <span data-ttu-id="bae4c-728">Attributs</span><span class="sxs-lookup"><span data-stu-id="bae4c-728">Attributes</span></span> | <span data-ttu-id="bae4c-729">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="bae4c-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="bae4c-730">String &#124; Object</span></span>| |<span data-ttu-id="bae4c-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bae4c-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="bae4c-733">**OR**</span></span><br/><span data-ttu-id="bae4c-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="bae4c-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="bae4c-736">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-736">String</span></span> | <span data-ttu-id="bae4c-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-737">&lt;optional&gt;</span></span> | <span data-ttu-id="bae4c-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="bae4c-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="bae4c-741">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-741">&lt;optional&gt;</span></span> | <span data-ttu-id="bae4c-742">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="bae4c-743">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bae4c-743">String</span></span> | | <span data-ttu-id="bae4c-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="bae4c-746">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-746">String</span></span> | | <span data-ttu-id="bae4c-747">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="bae4c-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="bae4c-748">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bae4c-748">String</span></span> | | <span data-ttu-id="bae4c-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="bae4c-751">Booléen</span><span class="sxs-lookup"><span data-stu-id="bae4c-751">Boolean</span></span> | | <span data-ttu-id="bae4c-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="bae4c-754">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-754">String</span></span> | | <span data-ttu-id="bae4c-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="bae4c-758">function</span><span class="sxs-lookup"><span data-stu-id="bae4c-758">function</span></span> | <span data-ttu-id="bae4c-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-759">&lt;optional&gt;</span></span> | <span data-ttu-id="bae4c-760">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bae4c-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bae4c-761">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-761">Requirements</span></span>

|<span data-ttu-id="bae4c-762">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-762">Requirement</span></span>| <span data-ttu-id="bae4c-763">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-764">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-765">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-765">1.0</span></span>|
|[<span data-ttu-id="bae4c-766">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-767">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-768">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-769">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bae4c-770">Exemples</span><span class="sxs-lookup"><span data-stu-id="bae4c-770">Examples</span></span>

<span data-ttu-id="bae4c-771">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="bae4c-772">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="bae4c-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="bae4c-773">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="bae4c-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bae4c-774">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="bae4c-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="bae4c-775">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="bae4c-776">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="bae4c-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="bae4c-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="bae4c-778">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="bae4c-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-779">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bae4c-780">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bae4c-781">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="bae4c-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="bae4c-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-785">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-785">Parameters</span></span>

| <span data-ttu-id="bae4c-786">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-786">Name</span></span> | <span data-ttu-id="bae4c-787">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-787">Type</span></span> | <span data-ttu-id="bae4c-788">Attributs</span><span class="sxs-lookup"><span data-stu-id="bae4c-788">Attributes</span></span> | <span data-ttu-id="bae4c-789">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="bae4c-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="bae4c-790">String &#124; Object</span></span>| | <span data-ttu-id="bae4c-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bae4c-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="bae4c-793">**OR**</span></span><br/><span data-ttu-id="bae4c-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="bae4c-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="bae4c-796">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-796">String</span></span> | <span data-ttu-id="bae4c-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-797">&lt;optional&gt;</span></span> | <span data-ttu-id="bae4c-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="bae4c-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="bae4c-801">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-801">&lt;optional&gt;</span></span> | <span data-ttu-id="bae4c-802">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="bae4c-803">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bae4c-803">String</span></span> | | <span data-ttu-id="bae4c-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="bae4c-806">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-806">String</span></span> | | <span data-ttu-id="bae4c-807">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="bae4c-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="bae4c-808">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bae4c-808">String</span></span> | | <span data-ttu-id="bae4c-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="bae4c-811">Booléen</span><span class="sxs-lookup"><span data-stu-id="bae4c-811">Boolean</span></span> | | <span data-ttu-id="bae4c-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="bae4c-814">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-814">String</span></span> | | <span data-ttu-id="bae4c-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="bae4c-818">function</span><span class="sxs-lookup"><span data-stu-id="bae4c-818">function</span></span> | <span data-ttu-id="bae4c-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-819">&lt;optional&gt;</span></span> | <span data-ttu-id="bae4c-820">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bae4c-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bae4c-821">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-821">Requirements</span></span>

|<span data-ttu-id="bae4c-822">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-822">Requirement</span></span>| <span data-ttu-id="bae4c-823">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-824">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-825">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-825">1.0</span></span>|
|[<span data-ttu-id="bae4c-826">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-827">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-828">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-829">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bae4c-830">Exemples</span><span class="sxs-lookup"><span data-stu-id="bae4c-830">Examples</span></span>

<span data-ttu-id="bae4c-831">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="bae4c-832">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="bae4c-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="bae4c-833">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="bae4c-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bae4c-834">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="bae4c-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="bae4c-835">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="bae4c-836">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="bae4c-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="bae4c-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="bae4c-838">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="bae4c-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-839">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-840">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-840">Requirements</span></span>

|<span data-ttu-id="bae4c-841">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-841">Requirement</span></span>| <span data-ttu-id="bae4c-842">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-843">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-844">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-844">1.0</span></span>|
|[<span data-ttu-id="bae4c-845">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-846">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-847">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-848">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bae4c-849">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="bae4c-849">Returns:</span></span>

<span data-ttu-id="bae4c-850">Type : [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="bae4c-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="bae4c-851">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-851">Example</span></span>

<span data-ttu-id="bae4c-852">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="bae4c-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="bae4c-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="bae4c-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="bae4c-854">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="bae4c-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-855">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-856">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-856">Parameters</span></span>

|<span data-ttu-id="bae4c-857">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-857">Name</span></span>| <span data-ttu-id="bae4c-858">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-858">Type</span></span>| <span data-ttu-id="bae4c-859">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="bae4c-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="bae4c-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="bae4c-861">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="bae4c-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bae4c-862">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-862">Requirements</span></span>

|<span data-ttu-id="bae4c-863">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-863">Requirement</span></span>| <span data-ttu-id="bae4c-864">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-865">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-866">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-866">1.0</span></span>|
|[<span data-ttu-id="bae4c-867">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-868">Restreinte</span><span class="sxs-lookup"><span data-stu-id="bae4c-868">Restricted</span></span>|
|[<span data-ttu-id="bae4c-869">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-870">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bae4c-871">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="bae4c-871">Returns:</span></span>

<span data-ttu-id="bae4c-872">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="bae4c-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="bae4c-873">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="bae4c-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="bae4c-874">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="bae4c-875">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="bae4c-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="bae4c-876">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="bae4c-876">Value of `entityType`</span></span> | <span data-ttu-id="bae4c-877">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="bae4c-877">Type of objects in returned array</span></span> | <span data-ttu-id="bae4c-878">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="bae4c-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="bae4c-879">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-879">String</span></span> | <span data-ttu-id="bae4c-880">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="bae4c-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="bae4c-881">Contact</span><span class="sxs-lookup"><span data-stu-id="bae4c-881">Contact</span></span> | <span data-ttu-id="bae4c-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bae4c-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="bae4c-883">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-883">String</span></span> | <span data-ttu-id="bae4c-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bae4c-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="bae4c-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="bae4c-885">MeetingSuggestion</span></span> | <span data-ttu-id="bae4c-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bae4c-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="bae4c-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="bae4c-887">PhoneNumber</span></span> | <span data-ttu-id="bae4c-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="bae4c-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="bae4c-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="bae4c-889">TaskSuggestion</span></span> | <span data-ttu-id="bae4c-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bae4c-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="bae4c-891">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-891">String</span></span> | <span data-ttu-id="bae4c-892">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="bae4c-892">**Restricted**</span></span> |

<span data-ttu-id="bae4c-893">Type : Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="bae4c-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="bae4c-894">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-894">Example</span></span>

<span data-ttu-id="bae4c-895">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="bae4c-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="bae4c-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="bae4c-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="bae4c-897">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="bae4c-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-898">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bae4c-899">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="bae4c-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-900">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-900">Parameters</span></span>

|<span data-ttu-id="bae4c-901">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-901">Name</span></span>| <span data-ttu-id="bae4c-902">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-902">Type</span></span>| <span data-ttu-id="bae4c-903">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="bae4c-904">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-904">String</span></span>|<span data-ttu-id="bae4c-905">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="bae4c-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bae4c-906">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-906">Requirements</span></span>

|<span data-ttu-id="bae4c-907">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-907">Requirement</span></span>| <span data-ttu-id="bae4c-908">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-909">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-910">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-910">1.0</span></span>|
|[<span data-ttu-id="bae4c-911">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-912">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-913">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-914">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bae4c-915">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="bae4c-915">Returns:</span></span>

<span data-ttu-id="bae4c-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="bae4c-918">Type : Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="bae4c-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="bae4c-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="bae4c-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="bae4c-920">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="bae4c-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-921">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bae4c-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="bae4c-925">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="bae4c-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="bae4c-926">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="bae4c-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-930">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-930">Requirements</span></span>

|<span data-ttu-id="bae4c-931">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-931">Requirement</span></span>| <span data-ttu-id="bae4c-932">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-933">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-934">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-934">1.0</span></span>|
|[<span data-ttu-id="bae4c-935">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-936">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-937">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-938">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bae4c-939">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="bae4c-939">Returns:</span></span>

<span data-ttu-id="bae4c-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="bae4c-942">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="bae4c-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bae4c-943">Object</span><span class="sxs-lookup"><span data-stu-id="bae4c-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="bae4c-944">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-944">Example</span></span>

<span data-ttu-id="bae4c-945">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="bae4c-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="bae4c-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="bae4c-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="bae4c-947">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="bae4c-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-948">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bae4c-949">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="bae4c-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="bae4c-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-952">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-952">Parameters</span></span>

|<span data-ttu-id="bae4c-953">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-953">Name</span></span>| <span data-ttu-id="bae4c-954">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-954">Type</span></span>| <span data-ttu-id="bae4c-955">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="bae4c-956">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-956">String</span></span>|<span data-ttu-id="bae4c-957">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="bae4c-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bae4c-958">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-958">Requirements</span></span>

|<span data-ttu-id="bae4c-959">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-959">Requirement</span></span>| <span data-ttu-id="bae4c-960">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-961">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-962">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-962">1.0</span></span>|
|[<span data-ttu-id="bae4c-963">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-964">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-965">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-966">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bae4c-967">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="bae4c-967">Returns:</span></span>

<span data-ttu-id="bae4c-968">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="bae4c-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="bae4c-969">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="bae4c-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bae4c-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="bae4c-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="bae4c-971">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="bae4c-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="bae4c-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="bae4c-973">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="bae4c-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="bae4c-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-976">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-976">Parameters</span></span>

|<span data-ttu-id="bae4c-977">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-977">Name</span></span>| <span data-ttu-id="bae4c-978">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-978">Type</span></span>| <span data-ttu-id="bae4c-979">Attributs</span><span class="sxs-lookup"><span data-stu-id="bae4c-979">Attributes</span></span>| <span data-ttu-id="bae4c-980">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="bae4c-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="bae4c-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="bae4c-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="bae4c-985">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-985">Object</span></span>| <span data-ttu-id="bae4c-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-986">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-987">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bae4c-988">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-988">Object</span></span>| <span data-ttu-id="bae4c-989">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-989">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-990">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bae4c-991">fonction</span><span class="sxs-lookup"><span data-stu-id="bae4c-991">function</span></span>||<span data-ttu-id="bae4c-992">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bae4c-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bae4c-993">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="bae4c-994">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bae4c-995">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-995">Requirements</span></span>

|<span data-ttu-id="bae4c-996">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-996">Requirement</span></span>| <span data-ttu-id="bae4c-997">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-998">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-999">1.2</span><span class="sxs-lookup"><span data-stu-id="bae4c-999">1.2</span></span>|
|[<span data-ttu-id="bae4c-1000">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="bae4c-1002">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-1003">Composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="bae4c-1004">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="bae4c-1004">Returns:</span></span>

<span data-ttu-id="bae4c-1005">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="bae4c-1006">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="bae4c-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bae4c-1007">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="bae4c-1008">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="bae4c-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="bae4c-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="bae4c-1010">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="bae4c-1011">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="bae4c-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-1012">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-1013">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-1013">Requirements</span></span>

|<span data-ttu-id="bae4c-1014">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-1014">Requirement</span></span>| <span data-ttu-id="bae4c-1015">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-1016">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="bae4c-1017">1.6</span></span> |
|[<span data-ttu-id="bae4c-1018">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-1019">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-1020">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-1021">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bae4c-1022">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="bae4c-1022">Returns:</span></span>

<span data-ttu-id="bae4c-1023">Type : [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="bae4c-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="bae4c-1024">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-1024">Example</span></span>

<span data-ttu-id="bae4c-1025">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="bae4c-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="bae4c-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="bae4c-p164">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="bae4c-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-1029">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bae4c-p165">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="bae4c-1033">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="bae4c-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="bae4c-1034">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="bae4c-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bae4c-1038">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-1038">Requirements</span></span>

|<span data-ttu-id="bae4c-1039">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-1039">Requirement</span></span>| <span data-ttu-id="bae4c-1040">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-1041">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="bae4c-1042">1.6</span></span> |
|[<span data-ttu-id="bae4c-1043">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-1044">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-1045">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-1046">Lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bae4c-1047">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="bae4c-1047">Returns:</span></span>

<span data-ttu-id="bae4c-p167">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="bae4c-1050">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-1050">Example</span></span>

<span data-ttu-id="bae4c-1051">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="bae4c-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bae4c-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="bae4c-1053">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="bae4c-p168">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-1057">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1057">Parameters</span></span>

|<span data-ttu-id="bae4c-1058">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-1058">Name</span></span>| <span data-ttu-id="bae4c-1059">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-1059">Type</span></span>| <span data-ttu-id="bae4c-1060">Attributs</span><span class="sxs-lookup"><span data-stu-id="bae4c-1060">Attributes</span></span>| <span data-ttu-id="bae4c-1061">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="bae4c-1062">function</span><span class="sxs-lookup"><span data-stu-id="bae4c-1062">function</span></span>||<span data-ttu-id="bae4c-1063">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bae4c-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bae4c-1064">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="bae4c-1065">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="bae4c-1066">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-1066">Object</span></span>| <span data-ttu-id="bae4c-1067">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-1068">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="bae4c-1069">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bae4c-1070">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-1070">Requirements</span></span>

|<span data-ttu-id="bae4c-1071">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-1071">Requirement</span></span>| <span data-ttu-id="bae4c-1072">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-1073">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="bae4c-1074">1.0</span></span>|
|[<span data-ttu-id="bae4c-1075">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-1076">ReadItem</span></span>|
|[<span data-ttu-id="bae4c-1077">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-1078">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bae4c-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-1079">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-1079">Example</span></span>

<span data-ttu-id="bae4c-p171">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="bae4c-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bae4c-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="bae4c-1084">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="bae4c-p172">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-1089">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1089">Parameters</span></span>

|<span data-ttu-id="bae4c-1090">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-1090">Name</span></span>| <span data-ttu-id="bae4c-1091">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-1091">Type</span></span>| <span data-ttu-id="bae4c-1092">Attributs</span><span class="sxs-lookup"><span data-stu-id="bae4c-1092">Attributes</span></span>| <span data-ttu-id="bae4c-1093">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="bae4c-1094">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-1094">String</span></span>||<span data-ttu-id="bae4c-1095">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="bae4c-1096">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-1096">Object</span></span>| <span data-ttu-id="bae4c-1097">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-1098">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bae4c-1099">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-1099">Object</span></span>| <span data-ttu-id="bae4c-1100">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-1101">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bae4c-1102">fonction</span><span class="sxs-lookup"><span data-stu-id="bae4c-1102">function</span></span>| <span data-ttu-id="bae4c-1103">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-1104">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bae4c-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bae4c-1105">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bae4c-1106">Erreurs</span><span class="sxs-lookup"><span data-stu-id="bae4c-1106">Errors</span></span>

| <span data-ttu-id="bae4c-1107">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="bae4c-1107">Error code</span></span> | <span data-ttu-id="bae4c-1108">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="bae4c-1109">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bae4c-1110">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-1110">Requirements</span></span>

|<span data-ttu-id="bae4c-1111">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-1111">Requirement</span></span>| <span data-ttu-id="bae4c-1112">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-1113">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="bae4c-1114">1.1</span></span>|
|[<span data-ttu-id="bae4c-1115">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="bae4c-1117">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-1118">Composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-1119">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-1119">Example</span></span>

<span data-ttu-id="bae4c-1120">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="bae4c-1120">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="bae4c-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="bae4c-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="bae4c-1122">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="bae4c-p173">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-1126">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="bae4c-1127">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="bae4c-p175">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="bae4c-1131">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="bae4c-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="bae4c-1132">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1132">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="bae4c-1133">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1133">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="bae4c-1134">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1134">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-1135">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1135">Parameters</span></span>

|<span data-ttu-id="bae4c-1136">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-1136">Name</span></span>| <span data-ttu-id="bae4c-1137">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-1137">Type</span></span>| <span data-ttu-id="bae4c-1138">Attributs</span><span class="sxs-lookup"><span data-stu-id="bae4c-1138">Attributes</span></span>| <span data-ttu-id="bae4c-1139">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-1139">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="bae4c-1140">Object</span><span class="sxs-lookup"><span data-stu-id="bae4c-1140">Object</span></span>| <span data-ttu-id="bae4c-1141">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-1141">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-1142">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1142">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bae4c-1143">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-1143">Object</span></span>| <span data-ttu-id="bae4c-1144">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-1144">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-1145">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1145">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bae4c-1146">fonction</span><span class="sxs-lookup"><span data-stu-id="bae4c-1146">function</span></span>||<span data-ttu-id="bae4c-1147">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bae4c-1147">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bae4c-1148">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1148">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bae4c-1149">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-1149">Requirements</span></span>

|<span data-ttu-id="bae4c-1150">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-1150">Requirement</span></span>| <span data-ttu-id="bae4c-1151">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-1151">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-1152">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1152">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-1153">1.3</span><span class="sxs-lookup"><span data-stu-id="bae4c-1153">1.3</span></span>|
|[<span data-ttu-id="bae4c-1154">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-1154">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-1155">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-1155">ReadWriteItem</span></span>|
|[<span data-ttu-id="bae4c-1156">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-1156">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-1157">Composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-1157">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="bae4c-1158">範例</span><span class="sxs-lookup"><span data-stu-id="bae4c-1158">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="bae4c-p177">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="bae4c-1161">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="bae4c-1161">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="bae4c-1162">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1162">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="bae4c-p178">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bae4c-1166">Paramètres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1166">Parameters</span></span>

|<span data-ttu-id="bae4c-1167">Nom</span><span class="sxs-lookup"><span data-stu-id="bae4c-1167">Name</span></span>| <span data-ttu-id="bae4c-1168">Type</span><span class="sxs-lookup"><span data-stu-id="bae4c-1168">Type</span></span>| <span data-ttu-id="bae4c-1169">Attributs</span><span class="sxs-lookup"><span data-stu-id="bae4c-1169">Attributes</span></span>| <span data-ttu-id="bae4c-1170">Description</span><span class="sxs-lookup"><span data-stu-id="bae4c-1170">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="bae4c-1171">String</span><span class="sxs-lookup"><span data-stu-id="bae4c-1171">String</span></span>||<span data-ttu-id="bae4c-p179">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="bae4c-1175">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-1175">Object</span></span>| <span data-ttu-id="bae4c-1176">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-1176">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-1177">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1177">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bae4c-1178">Objet</span><span class="sxs-lookup"><span data-stu-id="bae4c-1178">Object</span></span>| <span data-ttu-id="bae4c-1179">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-1179">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-1180">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1180">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="bae4c-1181">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="bae4c-1181">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="bae4c-1182">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bae4c-1182">&lt;optional&gt;</span></span>|<span data-ttu-id="bae4c-p180">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="bae4c-p181">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="bae4c-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="bae4c-1187">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="bae4c-1187">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="bae4c-1188">fonction</span><span class="sxs-lookup"><span data-stu-id="bae4c-1188">function</span></span>||<span data-ttu-id="bae4c-1189">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bae4c-1189">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bae4c-1190">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bae4c-1190">Requirements</span></span>

|<span data-ttu-id="bae4c-1191">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bae4c-1191">Requirement</span></span>| <span data-ttu-id="bae4c-1192">Valeur</span><span class="sxs-lookup"><span data-stu-id="bae4c-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="bae4c-1193">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bae4c-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bae4c-1194">1.2</span><span class="sxs-lookup"><span data-stu-id="bae4c-1194">1.2</span></span>|
|[<span data-ttu-id="bae4c-1195">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bae4c-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bae4c-1196">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bae4c-1196">ReadWriteItem</span></span>|
|[<span data-ttu-id="bae4c-1197">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bae4c-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bae4c-1198">Composition</span><span class="sxs-lookup"><span data-stu-id="bae4c-1198">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bae4c-1199">Exemple</span><span class="sxs-lookup"><span data-stu-id="bae4c-1199">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
