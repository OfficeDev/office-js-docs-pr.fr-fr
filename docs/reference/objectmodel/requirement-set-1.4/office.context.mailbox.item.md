---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,4
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: 4a8a97403c43e4af4ee5d840d7a4fd843d5bd398
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167339"
---
# <a name="item"></a><span data-ttu-id="a578b-102">élément</span><span class="sxs-lookup"><span data-stu-id="a578b-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="a578b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="a578b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="a578b-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a578b-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-106">Requirements</span></span>

|<span data-ttu-id="a578b-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-107">Requirement</span></span>| <span data-ttu-id="a578b-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-110">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-110">1.0</span></span>|
|[<span data-ttu-id="a578b-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="a578b-112">Restricted</span></span>|
|[<span data-ttu-id="a578b-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a578b-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="a578b-115">Members and methods</span></span>

| <span data-ttu-id="a578b-116">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-116">Member</span></span> | <span data-ttu-id="a578b-117">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a578b-118">attachments</span><span class="sxs-lookup"><span data-stu-id="a578b-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="a578b-119">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-119">Member</span></span> |
| [<span data-ttu-id="a578b-120">bcc</span><span class="sxs-lookup"><span data-stu-id="a578b-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="a578b-121">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-121">Member</span></span> |
| [<span data-ttu-id="a578b-122">body</span><span class="sxs-lookup"><span data-stu-id="a578b-122">body</span></span>](#body-body) | <span data-ttu-id="a578b-123">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-123">Member</span></span> |
| [<span data-ttu-id="a578b-124">cc</span><span class="sxs-lookup"><span data-stu-id="a578b-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a578b-125">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-125">Member</span></span> |
| [<span data-ttu-id="a578b-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="a578b-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a578b-127">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-127">Member</span></span> |
| [<span data-ttu-id="a578b-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a578b-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a578b-129">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-129">Member</span></span> |
| [<span data-ttu-id="a578b-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a578b-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a578b-131">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-131">Member</span></span> |
| [<span data-ttu-id="a578b-132">end</span><span class="sxs-lookup"><span data-stu-id="a578b-132">end</span></span>](#end-datetime) | <span data-ttu-id="a578b-133">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-133">Member</span></span> |
| [<span data-ttu-id="a578b-134">from</span><span class="sxs-lookup"><span data-stu-id="a578b-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="a578b-135">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-135">Member</span></span> |
| [<span data-ttu-id="a578b-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a578b-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a578b-137">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-137">Member</span></span> |
| [<span data-ttu-id="a578b-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="a578b-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a578b-139">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-139">Member</span></span> |
| [<span data-ttu-id="a578b-140">itemId</span><span class="sxs-lookup"><span data-stu-id="a578b-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a578b-141">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-141">Member</span></span> |
| [<span data-ttu-id="a578b-142">itemType</span><span class="sxs-lookup"><span data-stu-id="a578b-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="a578b-143">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-143">Member</span></span> |
| [<span data-ttu-id="a578b-144">location</span><span class="sxs-lookup"><span data-stu-id="a578b-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="a578b-145">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-145">Member</span></span> |
| [<span data-ttu-id="a578b-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a578b-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a578b-147">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-147">Member</span></span> |
| [<span data-ttu-id="a578b-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="a578b-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="a578b-149">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-149">Member</span></span> |
| [<span data-ttu-id="a578b-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a578b-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a578b-151">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-151">Member</span></span> |
| [<span data-ttu-id="a578b-152">organizer</span><span class="sxs-lookup"><span data-stu-id="a578b-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="a578b-153">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-153">Member</span></span> |
| [<span data-ttu-id="a578b-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a578b-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a578b-155">Member</span><span class="sxs-lookup"><span data-stu-id="a578b-155">Member</span></span> |
| [<span data-ttu-id="a578b-156">sender</span><span class="sxs-lookup"><span data-stu-id="a578b-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="a578b-157">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-157">Member</span></span> |
| [<span data-ttu-id="a578b-158">start</span><span class="sxs-lookup"><span data-stu-id="a578b-158">start</span></span>](#start-datetime) | <span data-ttu-id="a578b-159">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-159">Member</span></span> |
| [<span data-ttu-id="a578b-160">subject</span><span class="sxs-lookup"><span data-stu-id="a578b-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="a578b-161">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-161">Member</span></span> |
| [<span data-ttu-id="a578b-162">to</span><span class="sxs-lookup"><span data-stu-id="a578b-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a578b-163">Membre</span><span class="sxs-lookup"><span data-stu-id="a578b-163">Member</span></span> |
| [<span data-ttu-id="a578b-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a578b-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a578b-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-165">Method</span></span> |
| [<span data-ttu-id="a578b-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a578b-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a578b-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-167">Method</span></span> |
| [<span data-ttu-id="a578b-168">close</span><span class="sxs-lookup"><span data-stu-id="a578b-168">close</span></span>](#close) | <span data-ttu-id="a578b-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-169">Method</span></span> |
| [<span data-ttu-id="a578b-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a578b-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="a578b-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-171">Method</span></span> |
| [<span data-ttu-id="a578b-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a578b-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="a578b-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-173">Method</span></span> |
| [<span data-ttu-id="a578b-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="a578b-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="a578b-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-175">Method</span></span> |
| [<span data-ttu-id="a578b-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a578b-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="a578b-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-177">Method</span></span> |
| [<span data-ttu-id="a578b-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a578b-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="a578b-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-179">Method</span></span> |
| [<span data-ttu-id="a578b-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a578b-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a578b-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-181">Method</span></span> |
| [<span data-ttu-id="a578b-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a578b-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a578b-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-183">Method</span></span> |
| [<span data-ttu-id="a578b-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a578b-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a578b-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-185">Method</span></span> |
| [<span data-ttu-id="a578b-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a578b-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a578b-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-187">Method</span></span> |
| [<span data-ttu-id="a578b-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a578b-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a578b-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-189">Method</span></span> |
| [<span data-ttu-id="a578b-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="a578b-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="a578b-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-191">Method</span></span> |
| [<span data-ttu-id="a578b-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a578b-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a578b-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="a578b-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a578b-194">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-194">Example</span></span>

<span data-ttu-id="a578b-195">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="a578b-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="a578b-196">Membres</span><span class="sxs-lookup"><span data-stu-id="a578b-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-14"></a><span data-ttu-id="a578b-197">pièces jointes : tableau. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="a578b-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

<span data-ttu-id="a578b-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-200">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="a578b-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a578b-201">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a578b-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-202">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-202">Type</span></span>

*   <span data-ttu-id="a578b-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="a578b-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-204">Requirements</span></span>

|<span data-ttu-id="a578b-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-205">Requirement</span></span>| <span data-ttu-id="a578b-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-208">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-208">1.0</span></span>|
|[<span data-ttu-id="a578b-209">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-210">ReadItem</span></span>|
|[<span data-ttu-id="a578b-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-212">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-213">Example</span></span>

<span data-ttu-id="a578b-214">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a578b-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="a578b-215">CCI : [destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-216">Obtient un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour la ligne CCI (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="a578b-216">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a578b-217">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-218">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-218">Type</span></span>

*   [<span data-ttu-id="a578b-219">Destinataires</span><span class="sxs-lookup"><span data-stu-id="a578b-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="a578b-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-220">Requirements</span></span>

|<span data-ttu-id="a578b-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-221">Requirement</span></span>| <span data-ttu-id="a578b-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-224">1.1</span><span class="sxs-lookup"><span data-stu-id="a578b-224">1.1</span></span>|
|[<span data-ttu-id="a578b-225">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-226">ReadItem</span></span>|
|[<span data-ttu-id="a578b-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-228">Composition</span><span class="sxs-lookup"><span data-stu-id="a578b-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-229">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-14"></a><span data-ttu-id="a578b-230">Body : [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-231">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-232">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-232">Type</span></span>

*   [<span data-ttu-id="a578b-233">Body</span><span class="sxs-lookup"><span data-stu-id="a578b-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="a578b-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-234">Requirements</span></span>

|<span data-ttu-id="a578b-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-235">Requirement</span></span>| <span data-ttu-id="a578b-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-238">1.1</span><span class="sxs-lookup"><span data-stu-id="a578b-238">1.1</span></span>|
|[<span data-ttu-id="a578b-239">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-240">ReadItem</span></span>|
|[<span data-ttu-id="a578b-241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-242">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-243">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-243">Example</span></span>

<span data-ttu-id="a578b-244">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="a578b-244">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="a578b-245">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-245">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="a578b-246">CC : Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-247">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="a578b-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a578b-248">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a578b-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a578b-249">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-249">Read mode</span></span>

<span data-ttu-id="a578b-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a578b-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="a578b-252">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a578b-252">Compose mode</span></span>

<span data-ttu-id="a578b-253">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="a578b-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a578b-254">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-254">Type</span></span>

*   <span data-ttu-id="a578b-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-256">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-256">Requirements</span></span>

|<span data-ttu-id="a578b-257">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-257">Requirement</span></span>| <span data-ttu-id="a578b-258">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-259">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-260">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-260">1.0</span></span>|
|[<span data-ttu-id="a578b-261">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-262">ReadItem</span></span>|
|[<span data-ttu-id="a578b-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="a578b-265">(Nullable) conversationId : chaîne</span><span class="sxs-lookup"><span data-stu-id="a578b-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="a578b-266">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="a578b-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a578b-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="a578b-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a578b-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="a578b-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-271">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-271">Type</span></span>

*   <span data-ttu-id="a578b-272">String</span><span class="sxs-lookup"><span data-stu-id="a578b-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-273">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-273">Requirements</span></span>

|<span data-ttu-id="a578b-274">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-274">Requirement</span></span>| <span data-ttu-id="a578b-275">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-276">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-277">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-277">1.0</span></span>|
|[<span data-ttu-id="a578b-278">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-279">ReadItem</span></span>|
|[<span data-ttu-id="a578b-280">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-281">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-282">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="a578b-283">dateTimeCreated : date</span><span class="sxs-lookup"><span data-stu-id="a578b-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="a578b-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-286">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-286">Type</span></span>

*   <span data-ttu-id="a578b-287">Date</span><span class="sxs-lookup"><span data-stu-id="a578b-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-288">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-288">Requirements</span></span>

|<span data-ttu-id="a578b-289">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-289">Requirement</span></span>| <span data-ttu-id="a578b-290">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-291">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-292">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-292">1.0</span></span>|
|[<span data-ttu-id="a578b-293">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-294">ReadItem</span></span>|
|[<span data-ttu-id="a578b-295">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-296">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-297">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="a578b-298">dateTimeModified : date</span><span class="sxs-lookup"><span data-stu-id="a578b-298">dateTimeModified: Date</span></span>

<span data-ttu-id="a578b-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-301">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="a578b-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-302">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-302">Type</span></span>

*   <span data-ttu-id="a578b-303">Date</span><span class="sxs-lookup"><span data-stu-id="a578b-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-304">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-304">Requirements</span></span>

|<span data-ttu-id="a578b-305">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-305">Requirement</span></span>| <span data-ttu-id="a578b-306">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-307">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-308">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-308">1.0</span></span>|
|[<span data-ttu-id="a578b-309">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-310">ReadItem</span></span>|
|[<span data-ttu-id="a578b-311">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-312">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="a578b-314">fin : date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-315">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a578b-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a578b-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="a578b-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a578b-318">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-318">Read mode</span></span>

<span data-ttu-id="a578b-319">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="a578b-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="a578b-320">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a578b-320">Compose mode</span></span>

<span data-ttu-id="a578b-321">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a578b-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a578b-322">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="a578b-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="a578b-323">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a578b-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="a578b-324">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-324">Type</span></span>

*   <span data-ttu-id="a578b-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-326">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-326">Requirements</span></span>

|<span data-ttu-id="a578b-327">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-327">Requirement</span></span>| <span data-ttu-id="a578b-328">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-329">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-330">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-330">1.0</span></span>|
|[<span data-ttu-id="a578b-331">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-332">ReadItem</span></span>|
|[<span data-ttu-id="a578b-333">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-334">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="a578b-335">de : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="a578b-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="a578b-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-340">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a578b-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-341">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-341">Type</span></span>

*   [<span data-ttu-id="a578b-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a578b-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="a578b-343">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-343">Requirements</span></span>

|<span data-ttu-id="a578b-344">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-344">Requirement</span></span>| <span data-ttu-id="a578b-345">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-346">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-347">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-347">1.0</span></span>|
|[<span data-ttu-id="a578b-348">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-349">ReadItem</span></span>|
|[<span data-ttu-id="a578b-350">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-351">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-352">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="a578b-353">internetMessageId : chaîne</span><span class="sxs-lookup"><span data-stu-id="a578b-353">internetMessageId: String</span></span>

<span data-ttu-id="a578b-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-356">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-356">Type</span></span>

*   <span data-ttu-id="a578b-357">String</span><span class="sxs-lookup"><span data-stu-id="a578b-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-358">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-358">Requirements</span></span>

|<span data-ttu-id="a578b-359">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-359">Requirement</span></span>| <span data-ttu-id="a578b-360">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-361">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-362">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-362">1.0</span></span>|
|[<span data-ttu-id="a578b-363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-364">ReadItem</span></span>|
|[<span data-ttu-id="a578b-365">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-366">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-367">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="a578b-368">itemClass : chaîne</span><span class="sxs-lookup"><span data-stu-id="a578b-368">itemClass: String</span></span>

<span data-ttu-id="a578b-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a578b-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a578b-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="a578b-373">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-373">Type</span></span> | <span data-ttu-id="a578b-374">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-374">Description</span></span> | <span data-ttu-id="a578b-375">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="a578b-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="a578b-376">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="a578b-376">Appointment items</span></span> | <span data-ttu-id="a578b-377">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="a578b-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="a578b-378">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="a578b-378">Message items</span></span> | <span data-ttu-id="a578b-379">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="a578b-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="a578b-380">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="a578b-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-381">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-381">Type</span></span>

*   <span data-ttu-id="a578b-382">String</span><span class="sxs-lookup"><span data-stu-id="a578b-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-383">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-383">Requirements</span></span>

|<span data-ttu-id="a578b-384">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-384">Requirement</span></span>| <span data-ttu-id="a578b-385">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-386">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-387">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-387">1.0</span></span>|
|[<span data-ttu-id="a578b-388">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-389">ReadItem</span></span>|
|[<span data-ttu-id="a578b-390">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-391">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-392">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a578b-393">(Nullable) itemId : String</span><span class="sxs-lookup"><span data-stu-id="a578b-393">(nullable) itemId: String</span></span>

<span data-ttu-id="a578b-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-396">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="a578b-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a578b-397">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="a578b-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a578b-398">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a578b-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a578b-399">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a578b-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="a578b-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-402">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-402">Type</span></span>

*   <span data-ttu-id="a578b-403">String</span><span class="sxs-lookup"><span data-stu-id="a578b-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-404">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-404">Requirements</span></span>

|<span data-ttu-id="a578b-405">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-405">Requirement</span></span>| <span data-ttu-id="a578b-406">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-407">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-408">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-408">1.0</span></span>|
|[<span data-ttu-id="a578b-409">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-410">ReadItem</span></span>|
|[<span data-ttu-id="a578b-411">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-412">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-413">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-413">Example</span></span>

<span data-ttu-id="a578b-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a578b-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-14"></a><span data-ttu-id="a578b-416">itemType : [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-417">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="a578b-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a578b-418">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a578b-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-419">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-419">Type</span></span>

*   [<span data-ttu-id="a578b-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a578b-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="a578b-421">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-421">Requirements</span></span>

|<span data-ttu-id="a578b-422">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-422">Requirement</span></span>| <span data-ttu-id="a578b-423">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-425">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-425">1.0</span></span>|
|[<span data-ttu-id="a578b-426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-427">ReadItem</span></span>|
|[<span data-ttu-id="a578b-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-429">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-430">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-430">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-14"></a><span data-ttu-id="a578b-431">Location : String | [Emplacement](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-432">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a578b-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a578b-433">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-433">Read mode</span></span>

<span data-ttu-id="a578b-434">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a578b-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="a578b-435">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a578b-435">Compose mode</span></span>

<span data-ttu-id="a578b-436">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a578b-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a578b-437">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-437">Type</span></span>

*   <span data-ttu-id="a578b-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-439">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-439">Requirements</span></span>

|<span data-ttu-id="a578b-440">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-440">Requirement</span></span>| <span data-ttu-id="a578b-441">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-442">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-443">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-443">1.0</span></span>|
|[<span data-ttu-id="a578b-444">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-445">ReadItem</span></span>|
|[<span data-ttu-id="a578b-446">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-447">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-447">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a578b-448">normalizedSubject : chaîne</span><span class="sxs-lookup"><span data-stu-id="a578b-448">normalizedSubject: String</span></span>

<span data-ttu-id="a578b-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a578b-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="a578b-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-453">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-453">Type</span></span>

*   <span data-ttu-id="a578b-454">String</span><span class="sxs-lookup"><span data-stu-id="a578b-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-455">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-455">Requirements</span></span>

|<span data-ttu-id="a578b-456">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-456">Requirement</span></span>| <span data-ttu-id="a578b-457">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-458">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-459">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-459">1.0</span></span>|
|[<span data-ttu-id="a578b-460">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-461">ReadItem</span></span>|
|[<span data-ttu-id="a578b-462">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-463">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-464">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-464">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-14"></a><span data-ttu-id="a578b-465">notificationMessages : [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-466">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-467">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-467">Type</span></span>

*   [<span data-ttu-id="a578b-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="a578b-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="a578b-469">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-469">Requirements</span></span>

|<span data-ttu-id="a578b-470">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-470">Requirement</span></span>| <span data-ttu-id="a578b-471">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-472">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-473">1.3</span><span class="sxs-lookup"><span data-stu-id="a578b-473">1.3</span></span>|
|[<span data-ttu-id="a578b-474">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-475">ReadItem</span></span>|
|[<span data-ttu-id="a578b-476">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-477">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-478">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-478">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="a578b-479">optionalAttendees : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|des[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.4) de tableau. <</span><span class="sxs-lookup"><span data-stu-id="a578b-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-480">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="a578b-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a578b-481">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a578b-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a578b-482">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-482">Read mode</span></span>

<span data-ttu-id="a578b-483">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="a578b-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a578b-484">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a578b-484">Compose mode</span></span>

<span data-ttu-id="a578b-485">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="a578b-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a578b-486">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-486">Type</span></span>

*   <span data-ttu-id="a578b-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-488">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-488">Requirements</span></span>

|<span data-ttu-id="a578b-489">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-489">Requirement</span></span>| <span data-ttu-id="a578b-490">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-491">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-492">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-492">1.0</span></span>|
|[<span data-ttu-id="a578b-493">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-494">ReadItem</span></span>|
|[<span data-ttu-id="a578b-495">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-496">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-496">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="a578b-497">Organisateur : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-500">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-500">Type</span></span>

*   [<span data-ttu-id="a578b-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a578b-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="a578b-502">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-502">Requirements</span></span>

|<span data-ttu-id="a578b-503">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-503">Requirement</span></span>| <span data-ttu-id="a578b-504">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-505">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-506">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-506">1.0</span></span>|
|[<span data-ttu-id="a578b-507">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-508">ReadItem</span></span>|
|[<span data-ttu-id="a578b-509">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-510">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-511">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-511">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="a578b-512">requiredAttendees : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|des[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.4) de tableau. <</span><span class="sxs-lookup"><span data-stu-id="a578b-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-513">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="a578b-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a578b-514">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a578b-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a578b-515">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-515">Read mode</span></span>

<span data-ttu-id="a578b-516">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="a578b-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a578b-517">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a578b-517">Compose mode</span></span>

<span data-ttu-id="a578b-518">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="a578b-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="a578b-519">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-519">Type</span></span>

*   <span data-ttu-id="a578b-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-521">Requirements</span></span>

|<span data-ttu-id="a578b-522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-522">Requirement</span></span>| <span data-ttu-id="a578b-523">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-524">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-525">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-525">1.0</span></span>|
|[<span data-ttu-id="a578b-526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-527">ReadItem</span></span>|
|[<span data-ttu-id="a578b-528">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-529">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="a578b-530">expéditeur : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a578b-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a578b-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="a578b-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-535">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a578b-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a578b-536">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-536">Type</span></span>

*   [<span data-ttu-id="a578b-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a578b-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="a578b-538">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-538">Requirements</span></span>

|<span data-ttu-id="a578b-539">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-539">Requirement</span></span>| <span data-ttu-id="a578b-540">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-541">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-542">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-542">1.0</span></span>|
|[<span data-ttu-id="a578b-543">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-544">ReadItem</span></span>|
|[<span data-ttu-id="a578b-545">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-546">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-547">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="a578b-548">début : date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-549">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a578b-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a578b-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="a578b-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a578b-552">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-552">Read mode</span></span>

<span data-ttu-id="a578b-553">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="a578b-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="a578b-554">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a578b-554">Compose mode</span></span>

<span data-ttu-id="a578b-555">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a578b-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a578b-556">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="a578b-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="a578b-557">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a578b-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="a578b-558">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-558">Type</span></span>

*   <span data-ttu-id="a578b-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-560">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-560">Requirements</span></span>

|<span data-ttu-id="a578b-561">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-561">Requirement</span></span>| <span data-ttu-id="a578b-562">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-563">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-564">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-564">1.0</span></span>|
|[<span data-ttu-id="a578b-565">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-566">ReadItem</span></span>|
|[<span data-ttu-id="a578b-567">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-568">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-14"></a><span data-ttu-id="a578b-569">Subject : String | [Objet](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-570">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a578b-571">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="a578b-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a578b-572">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-572">Read mode</span></span>

<span data-ttu-id="a578b-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a578b-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="a578b-575">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a578b-575">Compose mode</span></span>

<span data-ttu-id="a578b-576">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="a578b-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="a578b-577">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-577">Type</span></span>

*   <span data-ttu-id="a578b-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-579">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-579">Requirements</span></span>

|<span data-ttu-id="a578b-580">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-580">Requirement</span></span>| <span data-ttu-id="a578b-581">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-582">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-583">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-583">1.0</span></span>|
|[<span data-ttu-id="a578b-584">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-585">ReadItem</span></span>|
|[<span data-ttu-id="a578b-586">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-587">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="a578b-588">to : Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="a578b-589">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="a578b-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a578b-590">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a578b-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a578b-591">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-591">Read mode</span></span>

<span data-ttu-id="a578b-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a578b-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="a578b-594">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a578b-594">Compose mode</span></span>

<span data-ttu-id="a578b-595">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="a578b-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a578b-596">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-596">Type</span></span>

*   <span data-ttu-id="a578b-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-598">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-598">Requirements</span></span>

|<span data-ttu-id="a578b-599">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-599">Requirement</span></span>| <span data-ttu-id="a578b-600">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-601">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-602">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-602">1.0</span></span>|
|[<span data-ttu-id="a578b-603">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-604">ReadItem</span></span>|
|[<span data-ttu-id="a578b-605">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-606">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="a578b-607">Méthodes</span><span class="sxs-lookup"><span data-stu-id="a578b-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a578b-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a578b-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a578b-609">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="a578b-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a578b-610">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="a578b-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a578b-611">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a578b-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-612">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-612">Parameters</span></span>

|<span data-ttu-id="a578b-613">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-613">Name</span></span>| <span data-ttu-id="a578b-614">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-614">Type</span></span>| <span data-ttu-id="a578b-615">Attributs</span><span class="sxs-lookup"><span data-stu-id="a578b-615">Attributes</span></span>| <span data-ttu-id="a578b-616">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="a578b-617">String</span><span class="sxs-lookup"><span data-stu-id="a578b-617">String</span></span>||<span data-ttu-id="a578b-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="a578b-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a578b-620">String</span><span class="sxs-lookup"><span data-stu-id="a578b-620">String</span></span>||<span data-ttu-id="a578b-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a578b-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a578b-623">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-623">Object</span></span>| <span data-ttu-id="a578b-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-624">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-625">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a578b-625">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a578b-626">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-626">Object</span></span>| <span data-ttu-id="a578b-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-627">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-628">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-628">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a578b-629">fonction</span><span class="sxs-lookup"><span data-stu-id="a578b-629">function</span></span>| <span data-ttu-id="a578b-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-630">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-631">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a578b-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a578b-632">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a578b-632">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a578b-633">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="a578b-633">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a578b-634">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a578b-634">Errors</span></span>

| <span data-ttu-id="a578b-635">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a578b-635">Error code</span></span> | <span data-ttu-id="a578b-636">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-636">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="a578b-637">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="a578b-637">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="a578b-638">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="a578b-638">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a578b-639">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a578b-639">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a578b-640">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-640">Requirements</span></span>

|<span data-ttu-id="a578b-641">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-641">Requirement</span></span>| <span data-ttu-id="a578b-642">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-643">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-644">1.1</span><span class="sxs-lookup"><span data-stu-id="a578b-644">1.1</span></span>|
|[<span data-ttu-id="a578b-645">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-646">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a578b-646">ReadWriteItem</span></span>|
|[<span data-ttu-id="a578b-647">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-648">Composition</span><span class="sxs-lookup"><span data-stu-id="a578b-648">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-649">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-649">Example</span></span>

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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a578b-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a578b-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a578b-651">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a578b-651">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a578b-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a578b-655">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a578b-655">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a578b-656">Si votre complément Office est en cours d’exécution dans Outlook sur le Web, `addItemAttachmentAsync` la méthode peut joindre des éléments à des éléments autres que l’élément que vous modifiez ; Toutefois, cette option n’est pas prise en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="a578b-656">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-657">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-657">Parameters</span></span>

|<span data-ttu-id="a578b-658">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-658">Name</span></span>| <span data-ttu-id="a578b-659">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-659">Type</span></span>| <span data-ttu-id="a578b-660">Attributs</span><span class="sxs-lookup"><span data-stu-id="a578b-660">Attributes</span></span>| <span data-ttu-id="a578b-661">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-661">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="a578b-662">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a578b-662">String</span></span>||<span data-ttu-id="a578b-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a578b-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a578b-665">String</span><span class="sxs-lookup"><span data-stu-id="a578b-665">String</span></span>||<span data-ttu-id="a578b-666">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="a578b-666">The subject of the item to be attached.</span></span> <span data-ttu-id="a578b-667">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a578b-667">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a578b-668">Object</span><span class="sxs-lookup"><span data-stu-id="a578b-668">Object</span></span>| <span data-ttu-id="a578b-669">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-669">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-670">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a578b-670">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a578b-671">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-671">Object</span></span>| <span data-ttu-id="a578b-672">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-672">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-673">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-673">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a578b-674">fonction</span><span class="sxs-lookup"><span data-stu-id="a578b-674">function</span></span>| <span data-ttu-id="a578b-675">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-675">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-676">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a578b-676">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a578b-677">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a578b-677">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a578b-678">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="a578b-678">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a578b-679">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a578b-679">Errors</span></span>

| <span data-ttu-id="a578b-680">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a578b-680">Error code</span></span> | <span data-ttu-id="a578b-681">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-681">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a578b-682">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a578b-682">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a578b-683">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-683">Requirements</span></span>

|<span data-ttu-id="a578b-684">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-684">Requirement</span></span>| <span data-ttu-id="a578b-685">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-685">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-686">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-686">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-687">1.1</span><span class="sxs-lookup"><span data-stu-id="a578b-687">1.1</span></span>|
|[<span data-ttu-id="a578b-688">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-688">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-689">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a578b-689">ReadWriteItem</span></span>|
|[<span data-ttu-id="a578b-690">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-690">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-691">Composition</span><span class="sxs-lookup"><span data-stu-id="a578b-691">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-692">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-692">Example</span></span>

<span data-ttu-id="a578b-693">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a578b-693">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="a578b-694">close()</span><span class="sxs-lookup"><span data-stu-id="a578b-694">close()</span></span>

<span data-ttu-id="a578b-695">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="a578b-695">Closes the current item that is being composed.</span></span>

<span data-ttu-id="a578b-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="a578b-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-698">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-698">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="a578b-699">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="a578b-699">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-700">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-700">Requirements</span></span>

|<span data-ttu-id="a578b-701">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-701">Requirement</span></span>| <span data-ttu-id="a578b-702">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-702">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-703">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-703">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-704">1.3</span><span class="sxs-lookup"><span data-stu-id="a578b-704">1.3</span></span>|
|[<span data-ttu-id="a578b-705">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-705">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-706">Restreinte</span><span class="sxs-lookup"><span data-stu-id="a578b-706">Restricted</span></span>|
|[<span data-ttu-id="a578b-707">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-707">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-708">Composition</span><span class="sxs-lookup"><span data-stu-id="a578b-708">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="a578b-709">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a578b-709">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="a578b-710">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a578b-710">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-711">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="a578b-711">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a578b-712">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="a578b-712">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a578b-713">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="a578b-713">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a578b-714">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="a578b-714">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="a578b-715">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="a578b-715">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="a578b-716">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="a578b-716">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-717">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-717">Parameters</span></span>

|<span data-ttu-id="a578b-718">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-718">Name</span></span>| <span data-ttu-id="a578b-719">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-719">Type</span></span>| <span data-ttu-id="a578b-720">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-720">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a578b-721">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a578b-721">String &#124; Object</span></span>| |<span data-ttu-id="a578b-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a578b-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a578b-724">**OU**</span><span class="sxs-lookup"><span data-stu-id="a578b-724">**OR**</span></span><br/><span data-ttu-id="a578b-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="a578b-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a578b-727">String</span><span class="sxs-lookup"><span data-stu-id="a578b-727">String</span></span> | <span data-ttu-id="a578b-728">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-728">&lt;optional&gt;</span></span> | <span data-ttu-id="a578b-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a578b-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="a578b-731">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-731">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="a578b-732">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-732">&lt;optional&gt;</span></span> | <span data-ttu-id="a578b-733">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-733">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="a578b-734">String</span><span class="sxs-lookup"><span data-stu-id="a578b-734">String</span></span> | | <span data-ttu-id="a578b-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="a578b-737">String</span><span class="sxs-lookup"><span data-stu-id="a578b-737">String</span></span> | | <span data-ttu-id="a578b-738">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a578b-738">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="a578b-739">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a578b-739">String</span></span> | | <span data-ttu-id="a578b-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="a578b-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="a578b-742">String</span><span class="sxs-lookup"><span data-stu-id="a578b-742">String</span></span> | | <span data-ttu-id="a578b-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a578b-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="a578b-746">function</span><span class="sxs-lookup"><span data-stu-id="a578b-746">function</span></span> | <span data-ttu-id="a578b-747">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-747">&lt;optional&gt;</span></span> | <span data-ttu-id="a578b-748">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a578b-748">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a578b-749">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-749">Requirements</span></span>

|<span data-ttu-id="a578b-750">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-750">Requirement</span></span>| <span data-ttu-id="a578b-751">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-751">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-752">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-752">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-753">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-753">1.0</span></span>|
|[<span data-ttu-id="a578b-754">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-754">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-755">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-755">ReadItem</span></span>|
|[<span data-ttu-id="a578b-756">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-756">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-757">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-757">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a578b-758">Exemples</span><span class="sxs-lookup"><span data-stu-id="a578b-758">Examples</span></span>

<span data-ttu-id="a578b-759">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a578b-759">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a578b-760">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="a578b-760">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a578b-761">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="a578b-761">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a578b-762">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="a578b-762">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a578b-763">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-763">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a578b-764">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-764">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="a578b-765">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a578b-765">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="a578b-766">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a578b-766">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-767">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="a578b-767">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a578b-768">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="a578b-768">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a578b-769">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="a578b-769">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a578b-770">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="a578b-770">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="a578b-771">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="a578b-771">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="a578b-772">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="a578b-772">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-773">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-773">Parameters</span></span>

|<span data-ttu-id="a578b-774">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-774">Name</span></span>| <span data-ttu-id="a578b-775">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-775">Type</span></span>| <span data-ttu-id="a578b-776">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-776">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a578b-777">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a578b-777">String &#124; Object</span></span>| | <span data-ttu-id="a578b-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a578b-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a578b-780">**OU**</span><span class="sxs-lookup"><span data-stu-id="a578b-780">**OR**</span></span><br/><span data-ttu-id="a578b-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="a578b-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a578b-783">String</span><span class="sxs-lookup"><span data-stu-id="a578b-783">String</span></span> | <span data-ttu-id="a578b-784">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-784">&lt;optional&gt;</span></span> | <span data-ttu-id="a578b-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a578b-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="a578b-787">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-787">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="a578b-788">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-788">&lt;optional&gt;</span></span> | <span data-ttu-id="a578b-789">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-789">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="a578b-790">String</span><span class="sxs-lookup"><span data-stu-id="a578b-790">String</span></span> | | <span data-ttu-id="a578b-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="a578b-793">String</span><span class="sxs-lookup"><span data-stu-id="a578b-793">String</span></span> | | <span data-ttu-id="a578b-794">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a578b-794">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="a578b-795">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a578b-795">String</span></span> | | <span data-ttu-id="a578b-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="a578b-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="a578b-798">String</span><span class="sxs-lookup"><span data-stu-id="a578b-798">String</span></span> | | <span data-ttu-id="a578b-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a578b-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="a578b-802">function</span><span class="sxs-lookup"><span data-stu-id="a578b-802">function</span></span> | <span data-ttu-id="a578b-803">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-803">&lt;optional&gt;</span></span> | <span data-ttu-id="a578b-804">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a578b-804">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a578b-805">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-805">Requirements</span></span>

|<span data-ttu-id="a578b-806">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-806">Requirement</span></span>| <span data-ttu-id="a578b-807">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-808">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-809">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-809">1.0</span></span>|
|[<span data-ttu-id="a578b-810">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-811">ReadItem</span></span>|
|[<span data-ttu-id="a578b-812">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-813">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-813">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a578b-814">Exemples</span><span class="sxs-lookup"><span data-stu-id="a578b-814">Examples</span></span>

<span data-ttu-id="a578b-815">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a578b-815">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a578b-816">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="a578b-816">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a578b-817">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="a578b-817">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a578b-818">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="a578b-818">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a578b-819">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-819">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a578b-820">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-820">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-14"></a><span data-ttu-id="a578b-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="a578b-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="a578b-822">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a578b-822">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-823">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="a578b-823">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-824">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-824">Requirements</span></span>

|<span data-ttu-id="a578b-825">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-825">Requirement</span></span>| <span data-ttu-id="a578b-826">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-827">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-828">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-828">1.0</span></span>|
|[<span data-ttu-id="a578b-829">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-830">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-830">ReadItem</span></span>|
|[<span data-ttu-id="a578b-831">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-832">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-832">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a578b-833">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a578b-833">Returns:</span></span>

<span data-ttu-id="a578b-834">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="a578b-834">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span></span>

##### <a name="example"></a><span data-ttu-id="a578b-835">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-835">Example</span></span>

<span data-ttu-id="a578b-836">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a578b-836">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="a578b-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="a578b-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="a578b-838">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a578b-838">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-839">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="a578b-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-840">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-840">Parameters</span></span>

|<span data-ttu-id="a578b-841">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-841">Name</span></span>| <span data-ttu-id="a578b-842">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-842">Type</span></span>| <span data-ttu-id="a578b-843">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-843">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="a578b-844">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a578b-844">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.4)|<span data-ttu-id="a578b-845">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="a578b-845">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a578b-846">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-846">Requirements</span></span>

|<span data-ttu-id="a578b-847">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-847">Requirement</span></span>| <span data-ttu-id="a578b-848">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-848">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-849">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-849">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-850">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-850">1.0</span></span>|
|[<span data-ttu-id="a578b-851">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-851">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-852">Restreinte</span><span class="sxs-lookup"><span data-stu-id="a578b-852">Restricted</span></span>|
|[<span data-ttu-id="a578b-853">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-853">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-854">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-854">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a578b-855">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a578b-855">Returns:</span></span>

<span data-ttu-id="a578b-856">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="a578b-856">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a578b-857">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="a578b-857">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a578b-858">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a578b-858">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a578b-859">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="a578b-859">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="a578b-860">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="a578b-860">Value of `entityType`</span></span> | <span data-ttu-id="a578b-861">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="a578b-861">Type of objects in returned array</span></span> | <span data-ttu-id="a578b-862">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="a578b-862">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="a578b-863">String</span><span class="sxs-lookup"><span data-stu-id="a578b-863">String</span></span> | <span data-ttu-id="a578b-864">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a578b-864">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="a578b-865">Contact</span><span class="sxs-lookup"><span data-stu-id="a578b-865">Contact</span></span> | <span data-ttu-id="a578b-866">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a578b-866">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="a578b-867">String</span><span class="sxs-lookup"><span data-stu-id="a578b-867">String</span></span> | <span data-ttu-id="a578b-868">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a578b-868">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="a578b-869">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a578b-869">MeetingSuggestion</span></span> | <span data-ttu-id="a578b-870">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a578b-870">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="a578b-871">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a578b-871">PhoneNumber</span></span> | <span data-ttu-id="a578b-872">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a578b-872">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="a578b-873">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a578b-873">TaskSuggestion</span></span> | <span data-ttu-id="a578b-874">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a578b-874">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="a578b-875">String</span><span class="sxs-lookup"><span data-stu-id="a578b-875">String</span></span> | <span data-ttu-id="a578b-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a578b-876">**Restricted**</span></span> |

<span data-ttu-id="a578b-877">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="a578b-877">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

##### <a name="example"></a><span data-ttu-id="a578b-878">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-878">Example</span></span>

<span data-ttu-id="a578b-879">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a578b-879">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="a578b-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="a578b-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="a578b-881">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a578b-881">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-882">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="a578b-882">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a578b-883">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="a578b-883">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-884">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-884">Parameters</span></span>

|<span data-ttu-id="a578b-885">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-885">Name</span></span>| <span data-ttu-id="a578b-886">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-886">Type</span></span>| <span data-ttu-id="a578b-887">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a578b-888">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a578b-888">String</span></span>|<span data-ttu-id="a578b-889">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="a578b-889">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a578b-890">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-890">Requirements</span></span>

|<span data-ttu-id="a578b-891">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-891">Requirement</span></span>| <span data-ttu-id="a578b-892">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-893">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-894">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-894">1.0</span></span>|
|[<span data-ttu-id="a578b-895">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-896">ReadItem</span></span>|
|[<span data-ttu-id="a578b-897">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-898">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a578b-899">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a578b-899">Returns:</span></span>

<span data-ttu-id="a578b-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="a578b-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a578b-902">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="a578b-902">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="a578b-903">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a578b-903">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a578b-904">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a578b-904">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-905">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="a578b-905">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a578b-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="a578b-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a578b-909">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="a578b-909">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a578b-910">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a578b-910">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a578b-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a578b-914">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-914">Requirements</span></span>

|<span data-ttu-id="a578b-915">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-915">Requirement</span></span>| <span data-ttu-id="a578b-916">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-917">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-918">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-918">1.0</span></span>|
|[<span data-ttu-id="a578b-919">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-919">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-920">ReadItem</span></span>|
|[<span data-ttu-id="a578b-921">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-921">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-922">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-922">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a578b-923">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a578b-923">Returns:</span></span>

<span data-ttu-id="a578b-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="a578b-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="a578b-926">Type : objet</span><span class="sxs-lookup"><span data-stu-id="a578b-926">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="a578b-927">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-927">Example</span></span>

<span data-ttu-id="a578b-928">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="a578b-928">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a578b-929">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="a578b-929">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a578b-930">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a578b-930">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-931">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="a578b-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a578b-932">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="a578b-932">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a578b-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="a578b-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-935">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-935">Parameters</span></span>

|<span data-ttu-id="a578b-936">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-936">Name</span></span>| <span data-ttu-id="a578b-937">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-937">Type</span></span>| <span data-ttu-id="a578b-938">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-938">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a578b-939">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a578b-939">String</span></span>|<span data-ttu-id="a578b-940">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="a578b-940">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a578b-941">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-941">Requirements</span></span>

|<span data-ttu-id="a578b-942">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-942">Requirement</span></span>| <span data-ttu-id="a578b-943">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-944">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-945">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-945">1.0</span></span>|
|[<span data-ttu-id="a578b-946">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-947">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-947">ReadItem</span></span>|
|[<span data-ttu-id="a578b-948">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-949">Lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-949">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a578b-950">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a578b-950">Returns:</span></span>

<span data-ttu-id="a578b-951">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a578b-951">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="a578b-952">Type : Array. < String ></span><span class="sxs-lookup"><span data-stu-id="a578b-952">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="a578b-953">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-953">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a578b-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a578b-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a578b-955">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="a578b-955">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a578b-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="a578b-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-958">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-958">Parameters</span></span>

|<span data-ttu-id="a578b-959">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-959">Name</span></span>| <span data-ttu-id="a578b-960">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-960">Type</span></span>| <span data-ttu-id="a578b-961">Attributs</span><span class="sxs-lookup"><span data-stu-id="a578b-961">Attributes</span></span>| <span data-ttu-id="a578b-962">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-962">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="a578b-963">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a578b-963">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a578b-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="a578b-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="a578b-967">Object</span><span class="sxs-lookup"><span data-stu-id="a578b-967">Object</span></span>| <span data-ttu-id="a578b-968">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-968">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-969">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a578b-969">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a578b-970">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-970">Object</span></span>| <span data-ttu-id="a578b-971">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-971">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-972">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-972">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a578b-973">fonction</span><span class="sxs-lookup"><span data-stu-id="a578b-973">function</span></span>||<span data-ttu-id="a578b-974">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a578b-974">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a578b-975">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="a578b-975">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a578b-976">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="a578b-976">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a578b-977">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-977">Requirements</span></span>

|<span data-ttu-id="a578b-978">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-978">Requirement</span></span>| <span data-ttu-id="a578b-979">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-979">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-980">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-980">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-981">1.2</span><span class="sxs-lookup"><span data-stu-id="a578b-981">1.2</span></span>|
|[<span data-ttu-id="a578b-982">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-982">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-983">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-983">ReadItem</span></span>|
|[<span data-ttu-id="a578b-984">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-984">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-985">Composition</span><span class="sxs-lookup"><span data-stu-id="a578b-985">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a578b-986">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a578b-986">Returns:</span></span>

<span data-ttu-id="a578b-987">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="a578b-987">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="a578b-988">Type : String</span><span class="sxs-lookup"><span data-stu-id="a578b-988">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a578b-989">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-989">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a578b-990">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a578b-990">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a578b-991">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a578b-991">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a578b-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="a578b-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-995">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-995">Parameters</span></span>

|<span data-ttu-id="a578b-996">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-996">Name</span></span>| <span data-ttu-id="a578b-997">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-997">Type</span></span>| <span data-ttu-id="a578b-998">Attributs</span><span class="sxs-lookup"><span data-stu-id="a578b-998">Attributes</span></span>| <span data-ttu-id="a578b-999">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-999">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a578b-1000">function</span><span class="sxs-lookup"><span data-stu-id="a578b-1000">function</span></span>||<span data-ttu-id="a578b-1001">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a578b-1001">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a578b-1002">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a578b-1002">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a578b-1003">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="a578b-1003">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="a578b-1004">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-1004">Object</span></span>| <span data-ttu-id="a578b-1005">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-1005">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-1006">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-1006">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a578b-1007">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-1007">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a578b-1008">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-1008">Requirements</span></span>

|<span data-ttu-id="a578b-1009">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-1009">Requirement</span></span>| <span data-ttu-id="a578b-1010">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-1010">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-1011">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-1011">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-1012">1.0</span><span class="sxs-lookup"><span data-stu-id="a578b-1012">1.0</span></span>|
|[<span data-ttu-id="a578b-1013">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-1013">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-1014">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a578b-1014">ReadItem</span></span>|
|[<span data-ttu-id="a578b-1015">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-1015">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-1016">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a578b-1016">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-1017">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-1017">Example</span></span>

<span data-ttu-id="a578b-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="a578b-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a578b-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a578b-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a578b-1022">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a578b-1022">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a578b-1023">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-1023">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="a578b-1024">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="a578b-1024">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="a578b-1025">Dans Outlook sur le Web et les appareils mobiles, l’identificateur de pièce jointe est valide uniquement au sein de la même session.</span><span class="sxs-lookup"><span data-stu-id="a578b-1025">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="a578b-1026">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="a578b-1026">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-1027">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-1027">Parameters</span></span>

|<span data-ttu-id="a578b-1028">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-1028">Name</span></span>| <span data-ttu-id="a578b-1029">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-1029">Type</span></span>| <span data-ttu-id="a578b-1030">Attributs</span><span class="sxs-lookup"><span data-stu-id="a578b-1030">Attributes</span></span>| <span data-ttu-id="a578b-1031">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-1031">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="a578b-1032">String</span><span class="sxs-lookup"><span data-stu-id="a578b-1032">String</span></span>||<span data-ttu-id="a578b-1033">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="a578b-1033">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="a578b-1034">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-1034">Object</span></span>| <span data-ttu-id="a578b-1035">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-1036">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a578b-1036">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a578b-1037">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-1037">Object</span></span>| <span data-ttu-id="a578b-1038">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-1039">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-1039">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a578b-1040">fonction</span><span class="sxs-lookup"><span data-stu-id="a578b-1040">function</span></span>| <span data-ttu-id="a578b-1041">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-1042">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a578b-1042">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a578b-1043">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="a578b-1043">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a578b-1044">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a578b-1044">Errors</span></span>

| <span data-ttu-id="a578b-1045">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a578b-1045">Error code</span></span> | <span data-ttu-id="a578b-1046">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-1046">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="a578b-1047">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="a578b-1047">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a578b-1048">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-1048">Requirements</span></span>

|<span data-ttu-id="a578b-1049">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-1049">Requirement</span></span>| <span data-ttu-id="a578b-1050">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-1051">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-1052">1.1</span><span class="sxs-lookup"><span data-stu-id="a578b-1052">1.1</span></span>|
|[<span data-ttu-id="a578b-1053">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-1053">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-1054">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a578b-1054">ReadWriteItem</span></span>|
|[<span data-ttu-id="a578b-1055">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-1055">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-1056">Composition</span><span class="sxs-lookup"><span data-stu-id="a578b-1056">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-1057">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-1057">Example</span></span>

<span data-ttu-id="a578b-1058">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="a578b-1058">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="a578b-1059">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a578b-1059">saveAsync([options], callback)</span></span>

<span data-ttu-id="a578b-1060">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a578b-1060">Asynchronously saves an item.</span></span>

<span data-ttu-id="a578b-1061">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-1061">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="a578b-1062">Dans Outlook sur le Web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="a578b-1062">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="a578b-1063">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="a578b-1063">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-1064">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="a578b-1064">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="a578b-1065">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="a578b-1065">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="a578b-p168">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="a578b-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="a578b-1069">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="a578b-1069">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="a578b-1070">Outlook sur Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="a578b-1070">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="a578b-1071">La `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="a578b-1071">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="a578b-1072">Consultez la rubrique [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide de l’API Office js](https://support.microsoft.com/help/4505745) pour obtenir une solution de contournement.</span><span class="sxs-lookup"><span data-stu-id="a578b-1072">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="a578b-1073">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="a578b-1073">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-1074">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-1074">Parameters</span></span>

|<span data-ttu-id="a578b-1075">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-1075">Name</span></span>| <span data-ttu-id="a578b-1076">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-1076">Type</span></span>| <span data-ttu-id="a578b-1077">Attributs</span><span class="sxs-lookup"><span data-stu-id="a578b-1077">Attributes</span></span>| <span data-ttu-id="a578b-1078">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-1078">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="a578b-1079">Object</span><span class="sxs-lookup"><span data-stu-id="a578b-1079">Object</span></span>| <span data-ttu-id="a578b-1080">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-1080">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-1081">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a578b-1081">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a578b-1082">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-1082">Object</span></span>| <span data-ttu-id="a578b-1083">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-1084">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-1084">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="a578b-1085">fonction</span><span class="sxs-lookup"><span data-stu-id="a578b-1085">function</span></span>||<span data-ttu-id="a578b-1086">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a578b-1086">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a578b-1087">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a578b-1087">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a578b-1088">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-1088">Requirements</span></span>

|<span data-ttu-id="a578b-1089">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-1089">Requirement</span></span>| <span data-ttu-id="a578b-1090">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-1091">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-1092">1.3</span><span class="sxs-lookup"><span data-stu-id="a578b-1092">1.3</span></span>|
|[<span data-ttu-id="a578b-1093">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-1093">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a578b-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="a578b-1095">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-1095">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-1096">Composition</span><span class="sxs-lookup"><span data-stu-id="a578b-1096">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a578b-1097">範例</span><span class="sxs-lookup"><span data-stu-id="a578b-1097">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="a578b-p170">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a578b-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a578b-1100">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a578b-1100">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a578b-1101">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a578b-1101">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a578b-p171">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="a578b-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a578b-1105">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a578b-1105">Parameters</span></span>

|<span data-ttu-id="a578b-1106">Nom</span><span class="sxs-lookup"><span data-stu-id="a578b-1106">Name</span></span>| <span data-ttu-id="a578b-1107">Type</span><span class="sxs-lookup"><span data-stu-id="a578b-1107">Type</span></span>| <span data-ttu-id="a578b-1108">Attributs</span><span class="sxs-lookup"><span data-stu-id="a578b-1108">Attributes</span></span>| <span data-ttu-id="a578b-1109">Description</span><span class="sxs-lookup"><span data-stu-id="a578b-1109">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="a578b-1110">String</span><span class="sxs-lookup"><span data-stu-id="a578b-1110">String</span></span>||<span data-ttu-id="a578b-p172">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="a578b-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="a578b-1114">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-1114">Object</span></span>| <span data-ttu-id="a578b-1115">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-1115">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-1116">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a578b-1116">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a578b-1117">Objet</span><span class="sxs-lookup"><span data-stu-id="a578b-1117">Object</span></span>| <span data-ttu-id="a578b-1118">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-1118">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-1119">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a578b-1119">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="a578b-1120">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a578b-1120">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="a578b-1121">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a578b-1121">&lt;optional&gt;</span></span>|<span data-ttu-id="a578b-1122">Si `text`, le style actuel est appliqué dans Outlook sur le Web et les clients de bureau.</span><span class="sxs-lookup"><span data-stu-id="a578b-1122">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="a578b-1123">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="a578b-1123">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a578b-1124">Si `html` et que le champ prend en charge le format html (l’objet ne l’est pas), le style actuel est appliqué dans Outlook sur le Web et le style par défaut est appliqué dans les clients de bureau Outlook.</span><span class="sxs-lookup"><span data-stu-id="a578b-1124">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="a578b-1125">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="a578b-1125">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a578b-1126">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="a578b-1126">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="a578b-1127">fonction</span><span class="sxs-lookup"><span data-stu-id="a578b-1127">function</span></span>||<span data-ttu-id="a578b-1128">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a578b-1128">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a578b-1129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a578b-1129">Requirements</span></span>

|<span data-ttu-id="a578b-1130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a578b-1130">Requirement</span></span>| <span data-ttu-id="a578b-1131">Valeur</span><span class="sxs-lookup"><span data-stu-id="a578b-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="a578b-1132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a578b-1132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a578b-1133">1.2</span><span class="sxs-lookup"><span data-stu-id="a578b-1133">1.2</span></span>|
|[<span data-ttu-id="a578b-1134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a578b-1134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a578b-1135">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a578b-1135">ReadWriteItem</span></span>|
|[<span data-ttu-id="a578b-1136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a578b-1136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a578b-1137">Composition</span><span class="sxs-lookup"><span data-stu-id="a578b-1137">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a578b-1138">Exemple</span><span class="sxs-lookup"><span data-stu-id="a578b-1138">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
