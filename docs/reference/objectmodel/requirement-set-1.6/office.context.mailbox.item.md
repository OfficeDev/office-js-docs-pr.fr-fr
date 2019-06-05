---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,6
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: 578e25b4fd7caf08087f24febdfd5b1877ed57bf
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706329"
---
# <a name="item"></a><span data-ttu-id="34b76-102">élément</span><span class="sxs-lookup"><span data-stu-id="34b76-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="34b76-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="34b76-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="34b76-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="34b76-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-106">Requirements</span></span>

|<span data-ttu-id="34b76-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-107">Requirement</span></span>| <span data-ttu-id="34b76-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-110">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-110">1.0</span></span>|
|[<span data-ttu-id="34b76-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="34b76-112">Restricted</span></span>|
|[<span data-ttu-id="34b76-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="34b76-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="34b76-115">Members and methods</span></span>

| <span data-ttu-id="34b76-116">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-116">Member</span></span> | <span data-ttu-id="34b76-117">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="34b76-118">attachments</span><span class="sxs-lookup"><span data-stu-id="34b76-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="34b76-119">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-119">Member</span></span> |
| [<span data-ttu-id="34b76-120">bcc</span><span class="sxs-lookup"><span data-stu-id="34b76-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="34b76-121">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-121">Member</span></span> |
| [<span data-ttu-id="34b76-122">body</span><span class="sxs-lookup"><span data-stu-id="34b76-122">body</span></span>](#body-body) | <span data-ttu-id="34b76-123">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-123">Member</span></span> |
| [<span data-ttu-id="34b76-124">cc</span><span class="sxs-lookup"><span data-stu-id="34b76-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="34b76-125">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-125">Member</span></span> |
| [<span data-ttu-id="34b76-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="34b76-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="34b76-127">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-127">Member</span></span> |
| [<span data-ttu-id="34b76-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="34b76-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="34b76-129">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-129">Member</span></span> |
| [<span data-ttu-id="34b76-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="34b76-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="34b76-131">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-131">Member</span></span> |
| [<span data-ttu-id="34b76-132">end</span><span class="sxs-lookup"><span data-stu-id="34b76-132">end</span></span>](#end-datetime) | <span data-ttu-id="34b76-133">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-133">Member</span></span> |
| [<span data-ttu-id="34b76-134">from</span><span class="sxs-lookup"><span data-stu-id="34b76-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="34b76-135">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-135">Member</span></span> |
| [<span data-ttu-id="34b76-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="34b76-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="34b76-137">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-137">Member</span></span> |
| [<span data-ttu-id="34b76-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="34b76-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="34b76-139">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-139">Member</span></span> |
| [<span data-ttu-id="34b76-140">itemId</span><span class="sxs-lookup"><span data-stu-id="34b76-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="34b76-141">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-141">Member</span></span> |
| [<span data-ttu-id="34b76-142">itemType</span><span class="sxs-lookup"><span data-stu-id="34b76-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="34b76-143">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-143">Member</span></span> |
| [<span data-ttu-id="34b76-144">location</span><span class="sxs-lookup"><span data-stu-id="34b76-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="34b76-145">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-145">Member</span></span> |
| [<span data-ttu-id="34b76-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="34b76-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="34b76-147">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-147">Member</span></span> |
| [<span data-ttu-id="34b76-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="34b76-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="34b76-149">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-149">Member</span></span> |
| [<span data-ttu-id="34b76-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="34b76-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="34b76-151">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-151">Member</span></span> |
| [<span data-ttu-id="34b76-152">organizer</span><span class="sxs-lookup"><span data-stu-id="34b76-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="34b76-153">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-153">Member</span></span> |
| [<span data-ttu-id="34b76-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="34b76-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="34b76-155">Member</span><span class="sxs-lookup"><span data-stu-id="34b76-155">Member</span></span> |
| [<span data-ttu-id="34b76-156">sender</span><span class="sxs-lookup"><span data-stu-id="34b76-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="34b76-157">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-157">Member</span></span> |
| [<span data-ttu-id="34b76-158">start</span><span class="sxs-lookup"><span data-stu-id="34b76-158">start</span></span>](#start-datetime) | <span data-ttu-id="34b76-159">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-159">Member</span></span> |
| [<span data-ttu-id="34b76-160">subject</span><span class="sxs-lookup"><span data-stu-id="34b76-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="34b76-161">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-161">Member</span></span> |
| [<span data-ttu-id="34b76-162">to</span><span class="sxs-lookup"><span data-stu-id="34b76-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="34b76-163">Membre</span><span class="sxs-lookup"><span data-stu-id="34b76-163">Member</span></span> |
| [<span data-ttu-id="34b76-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="34b76-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="34b76-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-165">Method</span></span> |
| [<span data-ttu-id="34b76-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="34b76-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="34b76-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-167">Method</span></span> |
| [<span data-ttu-id="34b76-168">close</span><span class="sxs-lookup"><span data-stu-id="34b76-168">close</span></span>](#close) | <span data-ttu-id="34b76-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-169">Method</span></span> |
| [<span data-ttu-id="34b76-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="34b76-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="34b76-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-171">Method</span></span> |
| [<span data-ttu-id="34b76-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="34b76-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="34b76-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-173">Method</span></span> |
| [<span data-ttu-id="34b76-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="34b76-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="34b76-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-175">Method</span></span> |
| [<span data-ttu-id="34b76-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="34b76-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="34b76-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-177">Method</span></span> |
| [<span data-ttu-id="34b76-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="34b76-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="34b76-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-179">Method</span></span> |
| [<span data-ttu-id="34b76-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="34b76-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="34b76-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-181">Method</span></span> |
| [<span data-ttu-id="34b76-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="34b76-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="34b76-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-183">Method</span></span> |
| [<span data-ttu-id="34b76-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="34b76-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="34b76-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-185">Method</span></span> |
| [<span data-ttu-id="34b76-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="34b76-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="34b76-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-187">Method</span></span> |
| [<span data-ttu-id="34b76-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="34b76-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="34b76-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-189">Method</span></span> |
| [<span data-ttu-id="34b76-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="34b76-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="34b76-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-191">Method</span></span> |
| [<span data-ttu-id="34b76-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="34b76-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="34b76-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-193">Method</span></span> |
| [<span data-ttu-id="34b76-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="34b76-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="34b76-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-195">Method</span></span> |
| [<span data-ttu-id="34b76-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="34b76-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="34b76-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="34b76-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="34b76-198">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-198">Example</span></span>

<span data-ttu-id="34b76-199">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="34b76-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="34b76-200">Membres</span><span class="sxs-lookup"><span data-stu-id="34b76-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="34b76-201">pièces jointes: tableau. <[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="34b76-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="34b76-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-204">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="34b76-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="34b76-205">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="34b76-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-206">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-206">Type</span></span>

*   <span data-ttu-id="34b76-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="34b76-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-208">Requirements</span></span>

|<span data-ttu-id="34b76-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-209">Requirement</span></span>| <span data-ttu-id="34b76-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-212">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-212">1.0</span></span>|
|[<span data-ttu-id="34b76-213">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-214">ReadItem</span></span>|
|[<span data-ttu-id="34b76-215">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-216">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-217">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-217">Example</span></span>

<span data-ttu-id="34b76-218">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="34b76-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="34b76-219">CCI: [destinataires](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34b76-219">bcc: [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="34b76-220">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="34b76-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="34b76-221">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-222">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-222">Type</span></span>

*   [<span data-ttu-id="34b76-223">Destinataires</span><span class="sxs-lookup"><span data-stu-id="34b76-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="34b76-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-224">Requirements</span></span>

|<span data-ttu-id="34b76-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-225">Requirement</span></span>| <span data-ttu-id="34b76-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-228">1.1</span><span class="sxs-lookup"><span data-stu-id="34b76-228">1.1</span></span>|
|[<span data-ttu-id="34b76-229">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-230">ReadItem</span></span>|
|[<span data-ttu-id="34b76-231">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-232">Composition</span><span class="sxs-lookup"><span data-stu-id="34b76-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-233">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="34b76-234">Body: [Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="34b76-234">body: [Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="34b76-235">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-236">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-236">Type</span></span>

*   [<span data-ttu-id="34b76-237">Body</span><span class="sxs-lookup"><span data-stu-id="34b76-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="34b76-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-238">Requirements</span></span>

|<span data-ttu-id="34b76-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-239">Requirement</span></span>| <span data-ttu-id="34b76-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-242">1.1</span><span class="sxs-lookup"><span data-stu-id="34b76-242">1.1</span></span>|
|[<span data-ttu-id="34b76-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-244">ReadItem</span></span>|
|[<span data-ttu-id="34b76-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-247">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-247">Example</span></span>

<span data-ttu-id="34b76-248">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="34b76-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="34b76-249">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="34b76-250">CC: Array. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34b76-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="34b76-251">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="34b76-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="34b76-252">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="34b76-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34b76-253">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-253">Read mode</span></span>

<span data-ttu-id="34b76-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="34b76-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="34b76-256">Mode composition</span><span class="sxs-lookup"><span data-stu-id="34b76-256">Compose mode</span></span>

<span data-ttu-id="34b76-257">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="34b76-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="34b76-258">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-258">Type</span></span>

*   <span data-ttu-id="34b76-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34b76-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-260">Requirements</span></span>

|<span data-ttu-id="34b76-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-261">Requirement</span></span>| <span data-ttu-id="34b76-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-264">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-264">1.0</span></span>|
|[<span data-ttu-id="34b76-265">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-266">ReadItem</span></span>|
|[<span data-ttu-id="34b76-267">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-268">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-268">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="34b76-269">(Nullable) conversationId: chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="34b76-270">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="34b76-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="34b76-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="34b76-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="34b76-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="34b76-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-275">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-275">Type</span></span>

*   <span data-ttu-id="34b76-276">String</span><span class="sxs-lookup"><span data-stu-id="34b76-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-277">Requirements</span></span>

|<span data-ttu-id="34b76-278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-278">Requirement</span></span>| <span data-ttu-id="34b76-279">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-281">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-281">1.0</span></span>|
|[<span data-ttu-id="34b76-282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-283">ReadItem</span></span>|
|[<span data-ttu-id="34b76-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-285">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-286">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="34b76-287">dateTimeCreated: date</span><span class="sxs-lookup"><span data-stu-id="34b76-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="34b76-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-290">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-290">Type</span></span>

*   <span data-ttu-id="34b76-291">Date</span><span class="sxs-lookup"><span data-stu-id="34b76-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-292">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-292">Requirements</span></span>

|<span data-ttu-id="34b76-293">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-293">Requirement</span></span>| <span data-ttu-id="34b76-294">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-295">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-296">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-296">1.0</span></span>|
|[<span data-ttu-id="34b76-297">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-298">ReadItem</span></span>|
|[<span data-ttu-id="34b76-299">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-300">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-301">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="34b76-302">dateTimeModified: date</span><span class="sxs-lookup"><span data-stu-id="34b76-302">dateTimeModified: Date</span></span>

<span data-ttu-id="34b76-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-305">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-305">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-306">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-306">Type</span></span>

*   <span data-ttu-id="34b76-307">Date</span><span class="sxs-lookup"><span data-stu-id="34b76-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-308">Requirements</span></span>

|<span data-ttu-id="34b76-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-309">Requirement</span></span>| <span data-ttu-id="34b76-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-312">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-312">1.0</span></span>|
|[<span data-ttu-id="34b76-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-314">ReadItem</span></span>|
|[<span data-ttu-id="34b76-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-316">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-317">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="34b76-318">fin: date | [Fois](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="34b76-318">end: Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="34b76-319">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="34b76-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="34b76-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="34b76-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34b76-322">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-322">Read mode</span></span>

<span data-ttu-id="34b76-323">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="34b76-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="34b76-324">Mode composition</span><span class="sxs-lookup"><span data-stu-id="34b76-324">Compose mode</span></span>

<span data-ttu-id="34b76-325">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="34b76-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="34b76-326">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="34b76-326">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="34b76-327">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="34b76-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="34b76-328">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-328">Type</span></span>

*   <span data-ttu-id="34b76-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="34b76-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-330">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-330">Requirements</span></span>

|<span data-ttu-id="34b76-331">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-331">Requirement</span></span>| <span data-ttu-id="34b76-332">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-333">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-334">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-334">1.0</span></span>|
|[<span data-ttu-id="34b76-335">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-336">ReadItem</span></span>|
|[<span data-ttu-id="34b76-337">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-338">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="34b76-339">de: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="34b76-339">from: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="34b76-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="34b76-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="34b76-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-344">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="34b76-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-345">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-345">Type</span></span>

*   [<span data-ttu-id="34b76-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="34b76-346">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="example"></a><span data-ttu-id="34b76-347">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="34b76-348">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-348">Requirements</span></span>

|<span data-ttu-id="34b76-349">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-349">Requirement</span></span>| <span data-ttu-id="34b76-350">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-351">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-352">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-352">1.0</span></span>|
|[<span data-ttu-id="34b76-353">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-354">ReadItem</span></span>|
|[<span data-ttu-id="34b76-355">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-356">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="34b76-357">internetMessageId: chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-357">internetMessageId: String</span></span>

<span data-ttu-id="34b76-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-360">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-360">Type</span></span>

*   <span data-ttu-id="34b76-361">String</span><span class="sxs-lookup"><span data-stu-id="34b76-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-362">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-362">Requirements</span></span>

|<span data-ttu-id="34b76-363">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-363">Requirement</span></span>| <span data-ttu-id="34b76-364">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-365">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-366">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-366">1.0</span></span>|
|[<span data-ttu-id="34b76-367">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-368">ReadItem</span></span>|
|[<span data-ttu-id="34b76-369">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-370">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-371">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="34b76-372">itemClass: chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-372">itemClass: String</span></span>

<span data-ttu-id="34b76-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="34b76-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="34b76-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="34b76-377">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-377">Type</span></span> | <span data-ttu-id="34b76-378">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-378">Description</span></span> | <span data-ttu-id="34b76-379">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="34b76-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="34b76-380">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="34b76-380">Appointment items</span></span> | <span data-ttu-id="34b76-381">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="34b76-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="34b76-382">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="34b76-382">Message items</span></span> | <span data-ttu-id="34b76-383">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="34b76-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="34b76-384">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="34b76-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-385">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-385">Type</span></span>

*   <span data-ttu-id="34b76-386">String</span><span class="sxs-lookup"><span data-stu-id="34b76-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-387">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-387">Requirements</span></span>

|<span data-ttu-id="34b76-388">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-388">Requirement</span></span>| <span data-ttu-id="34b76-389">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-390">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-391">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-391">1.0</span></span>|
|[<span data-ttu-id="34b76-392">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-393">ReadItem</span></span>|
|[<span data-ttu-id="34b76-394">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-395">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-396">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="34b76-397">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="34b76-397">(nullable) itemId: String</span></span>

<span data-ttu-id="34b76-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-400">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="34b76-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="34b76-401">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="34b76-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="34b76-402">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="34b76-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="34b76-403">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="34b76-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="34b76-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-406">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-406">Type</span></span>

*   <span data-ttu-id="34b76-407">String</span><span class="sxs-lookup"><span data-stu-id="34b76-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-408">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-408">Requirements</span></span>

|<span data-ttu-id="34b76-409">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-409">Requirement</span></span>| <span data-ttu-id="34b76-410">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-411">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-412">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-412">1.0</span></span>|
|[<span data-ttu-id="34b76-413">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-414">ReadItem</span></span>|
|[<span data-ttu-id="34b76-415">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-416">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-417">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-417">Example</span></span>

<span data-ttu-id="34b76-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="34b76-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="34b76-420">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="34b76-420">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="34b76-421">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="34b76-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="34b76-422">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="34b76-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-423">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-423">Type</span></span>

*   [<span data-ttu-id="34b76-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="34b76-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="34b76-425">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-425">Requirements</span></span>

|<span data-ttu-id="34b76-426">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-426">Requirement</span></span>| <span data-ttu-id="34b76-427">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-428">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-429">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-429">1.0</span></span>|
|[<span data-ttu-id="34b76-430">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-431">ReadItem</span></span>|
|[<span data-ttu-id="34b76-432">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-433">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-434">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="34b76-435">Location: String | [Emplacement](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="34b76-435">location: String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="34b76-436">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="34b76-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34b76-437">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-437">Read mode</span></span>

<span data-ttu-id="34b76-438">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="34b76-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="34b76-439">Mode composition</span><span class="sxs-lookup"><span data-stu-id="34b76-439">Compose mode</span></span>

<span data-ttu-id="34b76-440">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="34b76-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="34b76-441">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-441">Type</span></span>

*   <span data-ttu-id="34b76-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="34b76-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-443">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-443">Requirements</span></span>

|<span data-ttu-id="34b76-444">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-444">Requirement</span></span>| <span data-ttu-id="34b76-445">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-446">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-447">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-447">1.0</span></span>|
|[<span data-ttu-id="34b76-448">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-449">ReadItem</span></span>|
|[<span data-ttu-id="34b76-450">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-451">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="34b76-452">normalizedSubject: chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-452">normalizedSubject: String</span></span>

<span data-ttu-id="34b76-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="34b76-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="34b76-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-457">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-457">Type</span></span>

*   <span data-ttu-id="34b76-458">String</span><span class="sxs-lookup"><span data-stu-id="34b76-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-459">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-459">Requirements</span></span>

|<span data-ttu-id="34b76-460">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-460">Requirement</span></span>| <span data-ttu-id="34b76-461">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-462">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-463">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-463">1.0</span></span>|
|[<span data-ttu-id="34b76-464">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-465">ReadItem</span></span>|
|[<span data-ttu-id="34b76-466">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-467">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-468">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="34b76-469">notificationMessages: [notificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="34b76-469">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="34b76-470">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-471">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-471">Type</span></span>

*   [<span data-ttu-id="34b76-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="34b76-472">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="34b76-473">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-473">Requirements</span></span>

|<span data-ttu-id="34b76-474">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-474">Requirement</span></span>| <span data-ttu-id="34b76-475">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-476">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-477">1.3</span><span class="sxs-lookup"><span data-stu-id="34b76-477">1.3</span></span>|
|[<span data-ttu-id="34b76-478">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-479">ReadItem</span></span>|
|[<span data-ttu-id="34b76-480">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-481">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-482">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="34b76-483">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[](/javascript/api/outlook_1_6/office.recipients) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="34b76-483">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="34b76-484">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="34b76-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="34b76-485">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="34b76-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34b76-486">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-486">Read mode</span></span>

<span data-ttu-id="34b76-487">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="34b76-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="34b76-488">Mode composition</span><span class="sxs-lookup"><span data-stu-id="34b76-488">Compose mode</span></span>

<span data-ttu-id="34b76-489">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="34b76-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="34b76-490">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-490">Type</span></span>

*   <span data-ttu-id="34b76-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34b76-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-492">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-492">Requirements</span></span>

|<span data-ttu-id="34b76-493">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-493">Requirement</span></span>| <span data-ttu-id="34b76-494">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-495">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-496">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-496">1.0</span></span>|
|[<span data-ttu-id="34b76-497">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-498">ReadItem</span></span>|
|[<span data-ttu-id="34b76-499">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-500">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="34b76-501">Organisateur: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="34b76-501">organizer: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="34b76-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-504">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-504">Type</span></span>

*   [<span data-ttu-id="34b76-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="34b76-505">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="34b76-506">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-506">Requirements</span></span>

|<span data-ttu-id="34b76-507">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-507">Requirement</span></span>| <span data-ttu-id="34b76-508">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-509">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-510">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-510">1.0</span></span>|
|[<span data-ttu-id="34b76-511">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-512">ReadItem</span></span>|
|[<span data-ttu-id="34b76-513">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-514">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-515">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="34b76-516">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[](/javascript/api/outlook_1_6/office.recipients) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="34b76-516">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="34b76-517">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="34b76-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="34b76-518">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="34b76-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34b76-519">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-519">Read mode</span></span>

<span data-ttu-id="34b76-520">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="34b76-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="34b76-521">Mode composition</span><span class="sxs-lookup"><span data-stu-id="34b76-521">Compose mode</span></span>

<span data-ttu-id="34b76-522">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="34b76-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="34b76-523">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-523">Type</span></span>

*   <span data-ttu-id="34b76-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34b76-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-525">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-525">Requirements</span></span>

|<span data-ttu-id="34b76-526">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-526">Requirement</span></span>| <span data-ttu-id="34b76-527">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-528">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-529">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-529">1.0</span></span>|
|[<span data-ttu-id="34b76-530">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-531">ReadItem</span></span>|
|[<span data-ttu-id="34b76-532">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-533">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="34b76-534">expéditeur: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="34b76-534">sender: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="34b76-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="34b76-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="34b76-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="34b76-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-539">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="34b76-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="34b76-540">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-540">Type</span></span>

*   [<span data-ttu-id="34b76-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="34b76-541">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="34b76-542">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-542">Requirements</span></span>

|<span data-ttu-id="34b76-543">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-543">Requirement</span></span>| <span data-ttu-id="34b76-544">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-545">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-546">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-546">1.0</span></span>|
|[<span data-ttu-id="34b76-547">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-548">ReadItem</span></span>|
|[<span data-ttu-id="34b76-549">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-550">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-551">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="34b76-552">début: date | [Fois](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="34b76-552">start: Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="34b76-553">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="34b76-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="34b76-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="34b76-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34b76-556">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-556">Read mode</span></span>

<span data-ttu-id="34b76-557">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="34b76-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="34b76-558">Mode composition</span><span class="sxs-lookup"><span data-stu-id="34b76-558">Compose mode</span></span>

<span data-ttu-id="34b76-559">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="34b76-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="34b76-560">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="34b76-560">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="34b76-561">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="34b76-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="34b76-562">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-562">Type</span></span>

*   <span data-ttu-id="34b76-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="34b76-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-564">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-564">Requirements</span></span>

|<span data-ttu-id="34b76-565">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-565">Requirement</span></span>| <span data-ttu-id="34b76-566">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-567">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-568">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-568">1.0</span></span>|
|[<span data-ttu-id="34b76-569">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-570">ReadItem</span></span>|
|[<span data-ttu-id="34b76-571">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-572">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-572">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="34b76-573">Subject: String | [Objet](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="34b76-573">subject: String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="34b76-574">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="34b76-575">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="34b76-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34b76-576">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-576">Read mode</span></span>

<span data-ttu-id="34b76-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="34b76-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="34b76-579">Mode composition</span><span class="sxs-lookup"><span data-stu-id="34b76-579">Compose mode</span></span>

<span data-ttu-id="34b76-580">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="34b76-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="34b76-581">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-581">Type</span></span>

*   <span data-ttu-id="34b76-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="34b76-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-583">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-583">Requirements</span></span>

|<span data-ttu-id="34b76-584">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-584">Requirement</span></span>| <span data-ttu-id="34b76-585">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-586">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-587">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-587">1.0</span></span>|
|[<span data-ttu-id="34b76-588">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-589">ReadItem</span></span>|
|[<span data-ttu-id="34b76-590">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-591">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-591">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="34b76-592">to: Array. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34b76-592">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="34b76-593">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="34b76-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="34b76-594">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="34b76-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34b76-595">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-595">Read mode</span></span>

<span data-ttu-id="34b76-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="34b76-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="34b76-598">Mode composition</span><span class="sxs-lookup"><span data-stu-id="34b76-598">Compose mode</span></span>

<span data-ttu-id="34b76-599">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="34b76-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="34b76-600">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-600">Type</span></span>

*   <span data-ttu-id="34b76-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34b76-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-602">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-602">Requirements</span></span>

|<span data-ttu-id="34b76-603">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-603">Requirement</span></span>| <span data-ttu-id="34b76-604">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-605">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-606">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-606">1.0</span></span>|
|[<span data-ttu-id="34b76-607">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-608">ReadItem</span></span>|
|[<span data-ttu-id="34b76-609">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-610">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="34b76-611">Méthodes</span><span class="sxs-lookup"><span data-stu-id="34b76-611">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="34b76-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="34b76-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="34b76-613">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="34b76-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="34b76-614">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="34b76-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="34b76-615">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="34b76-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-616">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-616">Parameters</span></span>

|<span data-ttu-id="34b76-617">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-617">Name</span></span>| <span data-ttu-id="34b76-618">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-618">Type</span></span>| <span data-ttu-id="34b76-619">Attributs</span><span class="sxs-lookup"><span data-stu-id="34b76-619">Attributes</span></span>| <span data-ttu-id="34b76-620">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="34b76-621">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-621">String</span></span>||<span data-ttu-id="34b76-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="34b76-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="34b76-624">String</span><span class="sxs-lookup"><span data-stu-id="34b76-624">String</span></span>||<span data-ttu-id="34b76-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="34b76-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="34b76-627">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-627">Object</span></span>| <span data-ttu-id="34b76-628">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-628">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-629">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="34b76-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="34b76-630">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-630">Object</span></span> | <span data-ttu-id="34b76-631">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-631">&lt;optional&gt;</span></span> | <span data-ttu-id="34b76-632">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="34b76-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="34b76-633">Boolean</span></span> | <span data-ttu-id="34b76-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-634">&lt;optional&gt;</span></span> | <span data-ttu-id="34b76-635">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="34b76-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="34b76-636">fonction</span><span class="sxs-lookup"><span data-stu-id="34b76-636">function</span></span>| <span data-ttu-id="34b76-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-637">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-638">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="34b76-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="34b76-639">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="34b76-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="34b76-640">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="34b76-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="34b76-641">Erreurs</span><span class="sxs-lookup"><span data-stu-id="34b76-641">Errors</span></span>

| <span data-ttu-id="34b76-642">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="34b76-642">Error code</span></span> | <span data-ttu-id="34b76-643">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="34b76-644">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="34b76-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="34b76-645">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="34b76-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="34b76-646">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="34b76-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34b76-647">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-647">Requirements</span></span>

|<span data-ttu-id="34b76-648">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-648">Requirement</span></span>| <span data-ttu-id="34b76-649">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-650">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-651">1.1</span><span class="sxs-lookup"><span data-stu-id="34b76-651">1.1</span></span>|
|[<span data-ttu-id="34b76-652">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34b76-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="34b76-654">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-655">Composition</span><span class="sxs-lookup"><span data-stu-id="34b76-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="34b76-656">Exemples</span><span class="sxs-lookup"><span data-stu-id="34b76-656">Examples</span></span>

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

<span data-ttu-id="34b76-657">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="34b76-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="34b76-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="34b76-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="34b76-659">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="34b76-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="34b76-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="34b76-663">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="34b76-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="34b76-664">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="34b76-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-665">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-665">Parameters</span></span>

|<span data-ttu-id="34b76-666">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-666">Name</span></span>| <span data-ttu-id="34b76-667">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-667">Type</span></span>| <span data-ttu-id="34b76-668">Attributs</span><span class="sxs-lookup"><span data-stu-id="34b76-668">Attributes</span></span>| <span data-ttu-id="34b76-669">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="34b76-670">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-670">String</span></span>||<span data-ttu-id="34b76-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="34b76-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="34b76-673">String</span><span class="sxs-lookup"><span data-stu-id="34b76-673">String</span></span>||<span data-ttu-id="34b76-674">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="34b76-674">The subject of the item to be attached.</span></span> <span data-ttu-id="34b76-675">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="34b76-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="34b76-676">Object</span><span class="sxs-lookup"><span data-stu-id="34b76-676">Object</span></span>| <span data-ttu-id="34b76-677">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-677">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-678">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="34b76-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34b76-679">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-679">Object</span></span>| <span data-ttu-id="34b76-680">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-680">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-681">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="34b76-682">fonction</span><span class="sxs-lookup"><span data-stu-id="34b76-682">function</span></span>| <span data-ttu-id="34b76-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-683">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-684">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="34b76-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="34b76-685">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="34b76-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="34b76-686">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="34b76-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="34b76-687">Erreurs</span><span class="sxs-lookup"><span data-stu-id="34b76-687">Errors</span></span>

| <span data-ttu-id="34b76-688">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="34b76-688">Error code</span></span> | <span data-ttu-id="34b76-689">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="34b76-690">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="34b76-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34b76-691">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-691">Requirements</span></span>

|<span data-ttu-id="34b76-692">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-692">Requirement</span></span>| <span data-ttu-id="34b76-693">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-694">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-695">1.1</span><span class="sxs-lookup"><span data-stu-id="34b76-695">1.1</span></span>|
|[<span data-ttu-id="34b76-696">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34b76-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="34b76-698">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-699">Composition</span><span class="sxs-lookup"><span data-stu-id="34b76-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-700">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-700">Example</span></span>

<span data-ttu-id="34b76-701">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="34b76-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="34b76-702">close()</span><span class="sxs-lookup"><span data-stu-id="34b76-702">close()</span></span>

<span data-ttu-id="34b76-703">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="34b76-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="34b76-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="34b76-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-706">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="34b76-707">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="34b76-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-708">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-708">Requirements</span></span>

|<span data-ttu-id="34b76-709">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-709">Requirement</span></span>| <span data-ttu-id="34b76-710">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-711">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-712">1.3</span><span class="sxs-lookup"><span data-stu-id="34b76-712">1.3</span></span>|
|[<span data-ttu-id="34b76-713">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-714">Restreinte</span><span class="sxs-lookup"><span data-stu-id="34b76-714">Restricted</span></span>|
|[<span data-ttu-id="34b76-715">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-716">Composition</span><span class="sxs-lookup"><span data-stu-id="34b76-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="34b76-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="34b76-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="34b76-718">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="34b76-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-719">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34b76-720">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="34b76-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="34b76-721">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="34b76-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="34b76-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="34b76-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-725">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-725">Parameters</span></span>

| <span data-ttu-id="34b76-726">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-726">Name</span></span> | <span data-ttu-id="34b76-727">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-727">Type</span></span> | <span data-ttu-id="34b76-728">Attributs</span><span class="sxs-lookup"><span data-stu-id="34b76-728">Attributes</span></span> | <span data-ttu-id="34b76-729">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="34b76-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="34b76-730">String &#124; Object</span></span>| |<span data-ttu-id="34b76-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="34b76-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="34b76-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="34b76-733">**OR**</span></span><br/><span data-ttu-id="34b76-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="34b76-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="34b76-736">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-736">String</span></span> | <span data-ttu-id="34b76-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-737">&lt;optional&gt;</span></span> | <span data-ttu-id="34b76-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="34b76-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="34b76-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="34b76-741">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-741">&lt;optional&gt;</span></span> | <span data-ttu-id="34b76-742">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="34b76-743">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-743">String</span></span> | | <span data-ttu-id="34b76-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="34b76-746">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-746">String</span></span> | | <span data-ttu-id="34b76-747">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="34b76-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="34b76-748">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-748">String</span></span> | | <span data-ttu-id="34b76-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="34b76-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="34b76-751">Booléen</span><span class="sxs-lookup"><span data-stu-id="34b76-751">Boolean</span></span> | | <span data-ttu-id="34b76-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="34b76-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="34b76-754">String</span><span class="sxs-lookup"><span data-stu-id="34b76-754">String</span></span> | | <span data-ttu-id="34b76-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="34b76-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="34b76-758">function</span><span class="sxs-lookup"><span data-stu-id="34b76-758">function</span></span> | <span data-ttu-id="34b76-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-759">&lt;optional&gt;</span></span> | <span data-ttu-id="34b76-760">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="34b76-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34b76-761">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-761">Requirements</span></span>

|<span data-ttu-id="34b76-762">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-762">Requirement</span></span>| <span data-ttu-id="34b76-763">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-764">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-765">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-765">1.0</span></span>|
|[<span data-ttu-id="34b76-766">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-767">ReadItem</span></span>|
|[<span data-ttu-id="34b76-768">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-769">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="34b76-770">Exemples</span><span class="sxs-lookup"><span data-stu-id="34b76-770">Examples</span></span>

<span data-ttu-id="34b76-771">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="34b76-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="34b76-772">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="34b76-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="34b76-773">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="34b76-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="34b76-774">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="34b76-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="34b76-775">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="34b76-776">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="34b76-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="34b76-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="34b76-778">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="34b76-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-779">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34b76-780">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="34b76-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="34b76-781">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="34b76-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="34b76-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="34b76-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-785">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-785">Parameters</span></span>

| <span data-ttu-id="34b76-786">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-786">Name</span></span> | <span data-ttu-id="34b76-787">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-787">Type</span></span> | <span data-ttu-id="34b76-788">Attributs</span><span class="sxs-lookup"><span data-stu-id="34b76-788">Attributes</span></span> | <span data-ttu-id="34b76-789">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="34b76-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="34b76-790">String &#124; Object</span></span>| | <span data-ttu-id="34b76-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="34b76-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="34b76-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="34b76-793">**OR**</span></span><br/><span data-ttu-id="34b76-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="34b76-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="34b76-796">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-796">String</span></span> | <span data-ttu-id="34b76-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-797">&lt;optional&gt;</span></span> | <span data-ttu-id="34b76-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="34b76-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="34b76-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="34b76-801">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-801">&lt;optional&gt;</span></span> | <span data-ttu-id="34b76-802">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="34b76-803">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-803">String</span></span> | | <span data-ttu-id="34b76-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="34b76-806">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-806">String</span></span> | | <span data-ttu-id="34b76-807">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="34b76-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="34b76-808">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-808">String</span></span> | | <span data-ttu-id="34b76-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="34b76-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="34b76-811">Booléen</span><span class="sxs-lookup"><span data-stu-id="34b76-811">Boolean</span></span> | | <span data-ttu-id="34b76-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="34b76-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="34b76-814">String</span><span class="sxs-lookup"><span data-stu-id="34b76-814">String</span></span> | | <span data-ttu-id="34b76-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="34b76-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="34b76-818">function</span><span class="sxs-lookup"><span data-stu-id="34b76-818">function</span></span> | <span data-ttu-id="34b76-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-819">&lt;optional&gt;</span></span> | <span data-ttu-id="34b76-820">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="34b76-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34b76-821">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-821">Requirements</span></span>

|<span data-ttu-id="34b76-822">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-822">Requirement</span></span>| <span data-ttu-id="34b76-823">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-824">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-825">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-825">1.0</span></span>|
|[<span data-ttu-id="34b76-826">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-827">ReadItem</span></span>|
|[<span data-ttu-id="34b76-828">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-829">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="34b76-830">Exemples</span><span class="sxs-lookup"><span data-stu-id="34b76-830">Examples</span></span>

<span data-ttu-id="34b76-831">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="34b76-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="34b76-832">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="34b76-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="34b76-833">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="34b76-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="34b76-834">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="34b76-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="34b76-835">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="34b76-836">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="34b76-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="34b76-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="34b76-838">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="34b76-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-839">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-840">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-840">Requirements</span></span>

|<span data-ttu-id="34b76-841">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-841">Requirement</span></span>| <span data-ttu-id="34b76-842">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-843">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-844">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-844">1.0</span></span>|
|[<span data-ttu-id="34b76-845">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-846">ReadItem</span></span>|
|[<span data-ttu-id="34b76-847">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-848">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34b76-849">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="34b76-849">Returns:</span></span>

<span data-ttu-id="34b76-850">Type : [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="34b76-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="34b76-851">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-851">Example</span></span>

<span data-ttu-id="34b76-852">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="34b76-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="34b76-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="34b76-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="34b76-854">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="34b76-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-855">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-856">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-856">Parameters</span></span>

|<span data-ttu-id="34b76-857">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-857">Name</span></span>| <span data-ttu-id="34b76-858">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-858">Type</span></span>| <span data-ttu-id="34b76-859">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="34b76-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="34b76-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="34b76-861">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="34b76-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34b76-862">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-862">Requirements</span></span>

|<span data-ttu-id="34b76-863">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-863">Requirement</span></span>| <span data-ttu-id="34b76-864">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-865">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-866">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-866">1.0</span></span>|
|[<span data-ttu-id="34b76-867">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-868">Restreinte</span><span class="sxs-lookup"><span data-stu-id="34b76-868">Restricted</span></span>|
|[<span data-ttu-id="34b76-869">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-870">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34b76-871">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="34b76-871">Returns:</span></span>

<span data-ttu-id="34b76-872">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="34b76-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="34b76-873">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="34b76-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="34b76-874">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="34b76-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="34b76-875">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="34b76-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="34b76-876">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="34b76-876">Value of `entityType`</span></span> | <span data-ttu-id="34b76-877">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="34b76-877">Type of objects in returned array</span></span> | <span data-ttu-id="34b76-878">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="34b76-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="34b76-879">String</span><span class="sxs-lookup"><span data-stu-id="34b76-879">String</span></span> | <span data-ttu-id="34b76-880">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="34b76-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="34b76-881">Contact</span><span class="sxs-lookup"><span data-stu-id="34b76-881">Contact</span></span> | <span data-ttu-id="34b76-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="34b76-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="34b76-883">String</span><span class="sxs-lookup"><span data-stu-id="34b76-883">String</span></span> | <span data-ttu-id="34b76-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="34b76-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="34b76-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="34b76-885">MeetingSuggestion</span></span> | <span data-ttu-id="34b76-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="34b76-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="34b76-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="34b76-887">PhoneNumber</span></span> | <span data-ttu-id="34b76-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="34b76-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="34b76-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="34b76-889">TaskSuggestion</span></span> | <span data-ttu-id="34b76-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="34b76-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="34b76-891">String</span><span class="sxs-lookup"><span data-stu-id="34b76-891">String</span></span> | <span data-ttu-id="34b76-892">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="34b76-892">**Restricted**</span></span> |

<span data-ttu-id="34b76-893">Type : Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="34b76-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="34b76-894">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-894">Example</span></span>

<span data-ttu-id="34b76-895">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="34b76-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="34b76-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="34b76-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="34b76-897">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="34b76-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-898">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34b76-899">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="34b76-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-900">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-900">Parameters</span></span>

|<span data-ttu-id="34b76-901">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-901">Name</span></span>| <span data-ttu-id="34b76-902">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-902">Type</span></span>| <span data-ttu-id="34b76-903">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="34b76-904">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-904">String</span></span>|<span data-ttu-id="34b76-905">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="34b76-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34b76-906">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-906">Requirements</span></span>

|<span data-ttu-id="34b76-907">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-907">Requirement</span></span>| <span data-ttu-id="34b76-908">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-909">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-910">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-910">1.0</span></span>|
|[<span data-ttu-id="34b76-911">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-912">ReadItem</span></span>|
|[<span data-ttu-id="34b76-913">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-914">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34b76-915">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="34b76-915">Returns:</span></span>

<span data-ttu-id="34b76-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="34b76-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="34b76-918">Type : Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="34b76-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="34b76-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="34b76-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="34b76-920">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="34b76-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-921">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34b76-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="34b76-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="34b76-925">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="34b76-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="34b76-926">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="34b76-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="34b76-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-930">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-930">Requirements</span></span>

|<span data-ttu-id="34b76-931">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-931">Requirement</span></span>| <span data-ttu-id="34b76-932">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-933">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-934">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-934">1.0</span></span>|
|[<span data-ttu-id="34b76-935">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-936">ReadItem</span></span>|
|[<span data-ttu-id="34b76-937">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-938">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34b76-939">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="34b76-939">Returns:</span></span>

<span data-ttu-id="34b76-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="34b76-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="34b76-942">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="34b76-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="34b76-943">Object</span><span class="sxs-lookup"><span data-stu-id="34b76-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="34b76-944">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-944">Example</span></span>

<span data-ttu-id="34b76-945">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="34b76-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="34b76-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="34b76-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="34b76-947">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="34b76-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-948">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34b76-949">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="34b76-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="34b76-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="34b76-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-952">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-952">Parameters</span></span>

|<span data-ttu-id="34b76-953">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-953">Name</span></span>| <span data-ttu-id="34b76-954">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-954">Type</span></span>| <span data-ttu-id="34b76-955">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="34b76-956">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-956">String</span></span>|<span data-ttu-id="34b76-957">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="34b76-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34b76-958">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-958">Requirements</span></span>

|<span data-ttu-id="34b76-959">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-959">Requirement</span></span>| <span data-ttu-id="34b76-960">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-961">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-962">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-962">1.0</span></span>|
|[<span data-ttu-id="34b76-963">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-964">ReadItem</span></span>|
|[<span data-ttu-id="34b76-965">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-966">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34b76-967">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="34b76-967">Returns:</span></span>

<span data-ttu-id="34b76-968">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="34b76-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="34b76-969">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="34b76-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="34b76-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="34b76-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="34b76-971">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="34b76-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="34b76-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="34b76-973">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="34b76-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="34b76-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="34b76-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-976">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-976">Parameters</span></span>

|<span data-ttu-id="34b76-977">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-977">Name</span></span>| <span data-ttu-id="34b76-978">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-978">Type</span></span>| <span data-ttu-id="34b76-979">Attributs</span><span class="sxs-lookup"><span data-stu-id="34b76-979">Attributes</span></span>| <span data-ttu-id="34b76-980">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="34b76-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="34b76-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="34b76-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="34b76-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="34b76-985">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-985">Object</span></span>| <span data-ttu-id="34b76-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-986">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-987">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="34b76-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34b76-988">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-988">Object</span></span>| <span data-ttu-id="34b76-989">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-989">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-990">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="34b76-991">fonction</span><span class="sxs-lookup"><span data-stu-id="34b76-991">function</span></span>||<span data-ttu-id="34b76-992">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="34b76-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="34b76-993">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="34b76-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="34b76-994">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="34b76-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34b76-995">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-995">Requirements</span></span>

|<span data-ttu-id="34b76-996">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-996">Requirement</span></span>| <span data-ttu-id="34b76-997">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-998">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-999">1.2</span><span class="sxs-lookup"><span data-stu-id="34b76-999">1.2</span></span>|
|[<span data-ttu-id="34b76-1000">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34b76-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="34b76-1002">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-1003">Composition</span><span class="sxs-lookup"><span data-stu-id="34b76-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="34b76-1004">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="34b76-1004">Returns:</span></span>

<span data-ttu-id="34b76-1005">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="34b76-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="34b76-1006">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="34b76-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="34b76-1007">String</span><span class="sxs-lookup"><span data-stu-id="34b76-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="34b76-1008">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="34b76-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="34b76-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="34b76-1010">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="34b76-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="34b76-1011">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="34b76-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-1012">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-1013">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-1013">Requirements</span></span>

|<span data-ttu-id="34b76-1014">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-1014">Requirement</span></span>| <span data-ttu-id="34b76-1015">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-1016">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="34b76-1017">1.6</span></span> |
|[<span data-ttu-id="34b76-1018">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-1019">ReadItem</span></span>|
|[<span data-ttu-id="34b76-1020">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-1021">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34b76-1022">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="34b76-1022">Returns:</span></span>

<span data-ttu-id="34b76-1023">Type : [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="34b76-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="34b76-1024">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-1024">Example</span></span>

<span data-ttu-id="34b76-1025">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="34b76-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="34b76-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="34b76-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="34b76-p164">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="34b76-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-1029">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="34b76-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34b76-p165">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="34b76-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="34b76-1033">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="34b76-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="34b76-1034">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="34b76-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="34b76-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34b76-1038">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-1038">Requirements</span></span>

|<span data-ttu-id="34b76-1039">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-1039">Requirement</span></span>| <span data-ttu-id="34b76-1040">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-1041">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="34b76-1042">1.6</span></span> |
|[<span data-ttu-id="34b76-1043">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-1044">ReadItem</span></span>|
|[<span data-ttu-id="34b76-1045">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-1046">Lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34b76-1047">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="34b76-1047">Returns:</span></span>

<span data-ttu-id="34b76-p167">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="34b76-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="34b76-1050">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-1050">Example</span></span>

<span data-ttu-id="34b76-1051">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="34b76-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="34b76-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="34b76-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="34b76-1053">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="34b76-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="34b76-p168">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="34b76-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-1057">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-1057">Parameters</span></span>

|<span data-ttu-id="34b76-1058">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-1058">Name</span></span>| <span data-ttu-id="34b76-1059">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-1059">Type</span></span>| <span data-ttu-id="34b76-1060">Attributs</span><span class="sxs-lookup"><span data-stu-id="34b76-1060">Attributes</span></span>| <span data-ttu-id="34b76-1061">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="34b76-1062">function</span><span class="sxs-lookup"><span data-stu-id="34b76-1062">function</span></span>||<span data-ttu-id="34b76-1063">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="34b76-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="34b76-1064">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="34b76-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="34b76-1065">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="34b76-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="34b76-1066">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-1066">Object</span></span>| <span data-ttu-id="34b76-1067">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-1068">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="34b76-1069">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34b76-1070">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-1070">Requirements</span></span>

|<span data-ttu-id="34b76-1071">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-1071">Requirement</span></span>| <span data-ttu-id="34b76-1072">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-1073">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="34b76-1074">1.0</span></span>|
|[<span data-ttu-id="34b76-1075">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34b76-1076">ReadItem</span></span>|
|[<span data-ttu-id="34b76-1077">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-1078">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="34b76-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-1079">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-1079">Example</span></span>

<span data-ttu-id="34b76-p171">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="34b76-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="34b76-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="34b76-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="34b76-1084">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="34b76-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="34b76-p172">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="34b76-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-1089">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-1089">Parameters</span></span>

|<span data-ttu-id="34b76-1090">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-1090">Name</span></span>| <span data-ttu-id="34b76-1091">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-1091">Type</span></span>| <span data-ttu-id="34b76-1092">Attributs</span><span class="sxs-lookup"><span data-stu-id="34b76-1092">Attributes</span></span>| <span data-ttu-id="34b76-1093">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="34b76-1094">Chaîne</span><span class="sxs-lookup"><span data-stu-id="34b76-1094">String</span></span>||<span data-ttu-id="34b76-1095">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="34b76-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="34b76-1096">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-1096">Object</span></span>| <span data-ttu-id="34b76-1097">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-1098">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="34b76-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34b76-1099">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-1099">Object</span></span>| <span data-ttu-id="34b76-1100">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-1101">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="34b76-1102">fonction</span><span class="sxs-lookup"><span data-stu-id="34b76-1102">function</span></span>| <span data-ttu-id="34b76-1103">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-1104">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="34b76-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="34b76-1105">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="34b76-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="34b76-1106">Erreurs</span><span class="sxs-lookup"><span data-stu-id="34b76-1106">Errors</span></span>

| <span data-ttu-id="34b76-1107">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="34b76-1107">Error code</span></span> | <span data-ttu-id="34b76-1108">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="34b76-1109">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="34b76-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34b76-1110">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-1110">Requirements</span></span>

|<span data-ttu-id="34b76-1111">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-1111">Requirement</span></span>| <span data-ttu-id="34b76-1112">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-1113">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="34b76-1114">1.1</span></span>|
|[<span data-ttu-id="34b76-1115">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34b76-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="34b76-1117">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-1118">Composition</span><span class="sxs-lookup"><span data-stu-id="34b76-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-1119">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-1119">Example</span></span>

<span data-ttu-id="34b76-1120">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="34b76-1120">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="34b76-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="34b76-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="34b76-1122">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="34b76-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="34b76-p173">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="34b76-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-1126">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="34b76-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="34b76-1127">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="34b76-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="34b76-p175">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="34b76-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="34b76-1131">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="34b76-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="34b76-1132">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="34b76-1132">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="34b76-1133">La `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="34b76-1133">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="34b76-1134">Consultez la rubrique [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide de l’API Office js](https://support.microsoft.com/help/4505745) pour obtenir une solution de contournement.</span><span class="sxs-lookup"><span data-stu-id="34b76-1134">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="34b76-1135">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="34b76-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-1136">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-1136">Parameters</span></span>

|<span data-ttu-id="34b76-1137">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-1137">Name</span></span>| <span data-ttu-id="34b76-1138">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-1138">Type</span></span>| <span data-ttu-id="34b76-1139">Attributs</span><span class="sxs-lookup"><span data-stu-id="34b76-1139">Attributes</span></span>| <span data-ttu-id="34b76-1140">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="34b76-1141">Object</span><span class="sxs-lookup"><span data-stu-id="34b76-1141">Object</span></span>| <span data-ttu-id="34b76-1142">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-1143">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="34b76-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34b76-1144">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-1144">Object</span></span>| <span data-ttu-id="34b76-1145">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-1146">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="34b76-1147">fonction</span><span class="sxs-lookup"><span data-stu-id="34b76-1147">function</span></span>||<span data-ttu-id="34b76-1148">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="34b76-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="34b76-1149">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="34b76-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34b76-1150">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-1150">Requirements</span></span>

|<span data-ttu-id="34b76-1151">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-1151">Requirement</span></span>| <span data-ttu-id="34b76-1152">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-1153">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="34b76-1154">1.3</span></span>|
|[<span data-ttu-id="34b76-1155">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-1155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34b76-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="34b76-1157">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-1157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-1158">Composition</span><span class="sxs-lookup"><span data-stu-id="34b76-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="34b76-1159">範例</span><span class="sxs-lookup"><span data-stu-id="34b76-1159">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="34b76-p177">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="34b76-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="34b76-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="34b76-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="34b76-1163">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="34b76-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="34b76-p178">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="34b76-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34b76-1167">Paramètres</span><span class="sxs-lookup"><span data-stu-id="34b76-1167">Parameters</span></span>

|<span data-ttu-id="34b76-1168">Nom</span><span class="sxs-lookup"><span data-stu-id="34b76-1168">Name</span></span>| <span data-ttu-id="34b76-1169">Type</span><span class="sxs-lookup"><span data-stu-id="34b76-1169">Type</span></span>| <span data-ttu-id="34b76-1170">Attributs</span><span class="sxs-lookup"><span data-stu-id="34b76-1170">Attributes</span></span>| <span data-ttu-id="34b76-1171">Description</span><span class="sxs-lookup"><span data-stu-id="34b76-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="34b76-1172">String</span><span class="sxs-lookup"><span data-stu-id="34b76-1172">String</span></span>||<span data-ttu-id="34b76-p179">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="34b76-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="34b76-1176">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-1176">Object</span></span>| <span data-ttu-id="34b76-1177">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-1178">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="34b76-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34b76-1179">Objet</span><span class="sxs-lookup"><span data-stu-id="34b76-1179">Object</span></span>| <span data-ttu-id="34b76-1180">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-1181">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="34b76-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="34b76-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="34b76-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="34b76-1183">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34b76-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="34b76-p180">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="34b76-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="34b76-p181">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="34b76-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="34b76-1188">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="34b76-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="34b76-1189">fonction</span><span class="sxs-lookup"><span data-stu-id="34b76-1189">function</span></span>||<span data-ttu-id="34b76-1190">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="34b76-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34b76-1191">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="34b76-1191">Requirements</span></span>

|<span data-ttu-id="34b76-1192">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="34b76-1192">Requirement</span></span>| <span data-ttu-id="34b76-1193">Valeur</span><span class="sxs-lookup"><span data-stu-id="34b76-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="34b76-1194">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="34b76-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34b76-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="34b76-1195">1.2</span></span>|
|[<span data-ttu-id="34b76-1196">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="34b76-1196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34b76-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34b76-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="34b76-1198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="34b76-1198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34b76-1199">Composition</span><span class="sxs-lookup"><span data-stu-id="34b76-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="34b76-1200">Exemple</span><span class="sxs-lookup"><span data-stu-id="34b76-1200">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
