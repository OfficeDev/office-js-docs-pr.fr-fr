
# <a name="item"></a><span data-ttu-id="29088-101">élément</span><span class="sxs-lookup"><span data-stu-id="29088-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="29088-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="29088-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="29088-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="29088-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-105">Requirements</span></span>

|<span data-ttu-id="29088-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-106">Requirement</span></span>| <span data-ttu-id="29088-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-109">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-109">1.0</span></span>|
|[<span data-ttu-id="29088-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="29088-111">Restricted</span></span>|
|[<span data-ttu-id="29088-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="29088-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="29088-114">Members and methods</span></span>

| <span data-ttu-id="29088-115">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-115">Member</span></span> | <span data-ttu-id="29088-116">Type</span><span class="sxs-lookup"><span data-stu-id="29088-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="29088-117">attachments</span><span class="sxs-lookup"><span data-stu-id="29088-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="29088-118">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-118">Member</span></span> |
| [<span data-ttu-id="29088-119">bcc</span><span class="sxs-lookup"><span data-stu-id="29088-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="29088-120">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-120">Member</span></span> |
| [<span data-ttu-id="29088-121">body</span><span class="sxs-lookup"><span data-stu-id="29088-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="29088-122">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-122">Member</span></span> |
| [<span data-ttu-id="29088-123">cc</span><span class="sxs-lookup"><span data-stu-id="29088-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="29088-124">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-124">Member</span></span> |
| [<span data-ttu-id="29088-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="29088-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="29088-126">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-126">Member</span></span> |
| [<span data-ttu-id="29088-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="29088-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="29088-128">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-128">Member</span></span> |
| [<span data-ttu-id="29088-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="29088-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="29088-130">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-130">Member</span></span> |
| [<span data-ttu-id="29088-131">end</span><span class="sxs-lookup"><span data-stu-id="29088-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="29088-132">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-132">Member</span></span> |
| [<span data-ttu-id="29088-133">from</span><span class="sxs-lookup"><span data-stu-id="29088-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="29088-134">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-134">Member</span></span> |
| [<span data-ttu-id="29088-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="29088-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="29088-136">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-136">Member</span></span> |
| [<span data-ttu-id="29088-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="29088-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="29088-138">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-138">Member</span></span> |
| [<span data-ttu-id="29088-139">itemId</span><span class="sxs-lookup"><span data-stu-id="29088-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="29088-140">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-140">Member</span></span> |
| [<span data-ttu-id="29088-141">itemType</span><span class="sxs-lookup"><span data-stu-id="29088-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="29088-142">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-142">Member</span></span> |
| [<span data-ttu-id="29088-143">location</span><span class="sxs-lookup"><span data-stu-id="29088-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="29088-144">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-144">Member</span></span> |
| [<span data-ttu-id="29088-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="29088-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="29088-146">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-146">Member</span></span> |
| [<span data-ttu-id="29088-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="29088-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="29088-148">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-148">Member</span></span> |
| [<span data-ttu-id="29088-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="29088-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="29088-150">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-150">Member</span></span> |
| [<span data-ttu-id="29088-151">organizer</span><span class="sxs-lookup"><span data-stu-id="29088-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="29088-152">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-152">Member</span></span> |
| [<span data-ttu-id="29088-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="29088-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="29088-154">Member</span><span class="sxs-lookup"><span data-stu-id="29088-154">Member</span></span> |
| [<span data-ttu-id="29088-155">sender</span><span class="sxs-lookup"><span data-stu-id="29088-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="29088-156">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-156">Member</span></span> |
| [<span data-ttu-id="29088-157">start</span><span class="sxs-lookup"><span data-stu-id="29088-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="29088-158">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-158">Member</span></span> |
| [<span data-ttu-id="29088-159">subject</span><span class="sxs-lookup"><span data-stu-id="29088-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="29088-160">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-160">Member</span></span> |
| [<span data-ttu-id="29088-161">to</span><span class="sxs-lookup"><span data-stu-id="29088-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="29088-162">Membre</span><span class="sxs-lookup"><span data-stu-id="29088-162">Member</span></span> |
| [<span data-ttu-id="29088-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="29088-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="29088-164">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-164">Method</span></span> |
| [<span data-ttu-id="29088-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="29088-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="29088-166">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-166">Method</span></span> |
| [<span data-ttu-id="29088-167">close</span><span class="sxs-lookup"><span data-stu-id="29088-167">close</span></span>](#close) | <span data-ttu-id="29088-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-168">Method</span></span> |
| [<span data-ttu-id="29088-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="29088-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="29088-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-170">Method</span></span> |
| [<span data-ttu-id="29088-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="29088-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="29088-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-172">Method</span></span> |
| [<span data-ttu-id="29088-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="29088-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="29088-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-174">Method</span></span> |
| [<span data-ttu-id="29088-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="29088-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="29088-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-176">Method</span></span> |
| [<span data-ttu-id="29088-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="29088-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="29088-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-178">Method</span></span> |
| [<span data-ttu-id="29088-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="29088-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="29088-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-180">Method</span></span> |
| [<span data-ttu-id="29088-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="29088-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="29088-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-182">Method</span></span> |
| [<span data-ttu-id="29088-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="29088-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="29088-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-184">Method</span></span> |
| [<span data-ttu-id="29088-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="29088-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="29088-186">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-186">Method</span></span> |
| [<span data-ttu-id="29088-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="29088-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="29088-188">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-188">Method</span></span> |
| [<span data-ttu-id="29088-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="29088-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="29088-190">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-190">Method</span></span> |
| [<span data-ttu-id="29088-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="29088-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="29088-192">Méthode</span><span class="sxs-lookup"><span data-stu-id="29088-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="29088-193">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-193">Example</span></span>

<span data-ttu-id="29088-194">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="29088-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="29088-195">Membres</span><span class="sxs-lookup"><span data-stu-id="29088-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="29088-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="29088-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="29088-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-199">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="29088-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="29088-200">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="29088-200">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="29088-201">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-201">Type:</span></span>

*   <span data-ttu-id="29088-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="29088-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-203">Requirements</span></span>

|<span data-ttu-id="29088-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-204">Requirement</span></span>| <span data-ttu-id="29088-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-207">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-207">1.0</span></span>|
|[<span data-ttu-id="29088-208">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-209">ReadItem</span></span>|
|[<span data-ttu-id="29088-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-211">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-212">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-212">Example</span></span>

<span data-ttu-id="29088-213">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="29088-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="29088-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="29088-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="29088-215">Obtient un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="29088-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="29088-216">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-217">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-217">Type:</span></span>

*   [<span data-ttu-id="29088-218">Destinataires</span><span class="sxs-lookup"><span data-stu-id="29088-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="29088-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-219">Requirements</span></span>

|<span data-ttu-id="29088-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-220">Requirement</span></span>| <span data-ttu-id="29088-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-223">1.1</span><span class="sxs-lookup"><span data-stu-id="29088-223">1.1</span></span>|
|[<span data-ttu-id="29088-224">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-225">ReadItem</span></span>|
|[<span data-ttu-id="29088-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-227">Composition</span><span class="sxs-lookup"><span data-stu-id="29088-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-228">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-228">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="29088-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="29088-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="29088-230">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="29088-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-231">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-231">Type:</span></span>

*   [<span data-ttu-id="29088-232">Corps</span><span class="sxs-lookup"><span data-stu-id="29088-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="29088-233">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-233">Requirements</span></span>

|<span data-ttu-id="29088-234">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-234">Requirement</span></span>| <span data-ttu-id="29088-235">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-236">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-237">1.1</span><span class="sxs-lookup"><span data-stu-id="29088-237">1.1</span></span>|
|[<span data-ttu-id="29088-238">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-239">ReadItem</span></span>|
|[<span data-ttu-id="29088-240">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-241">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="29088-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="29088-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="29088-243">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="29088-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="29088-244">Le type d’objet et le niveau d’accès varie selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="29088-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="29088-245">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-245">Read mode</span></span>

<span data-ttu-id="29088-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="29088-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="29088-248">Mode composition</span><span class="sxs-lookup"><span data-stu-id="29088-248">Compose mode</span></span>

<span data-ttu-id="29088-249">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="29088-249">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-250">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-250">Type:</span></span>

*   <span data-ttu-id="29088-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="29088-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-252">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-252">Requirements</span></span>

|<span data-ttu-id="29088-253">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-253">Requirement</span></span>| <span data-ttu-id="29088-254">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-255">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-255">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-256">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-256">1.0</span></span>|
|[<span data-ttu-id="29088-257">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-258">ReadItem</span></span>|
|[<span data-ttu-id="29088-259">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-260">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-261">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-261">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="29088-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="29088-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="29088-263">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="29088-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="29088-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="29088-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="29088-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="29088-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-268">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-268">Type:</span></span>

*   <span data-ttu-id="29088-269">Chaîne</span><span class="sxs-lookup"><span data-stu-id="29088-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-270">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-270">Requirements</span></span>

|<span data-ttu-id="29088-271">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-271">Requirement</span></span>| <span data-ttu-id="29088-272">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-273">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-274">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-274">1.0</span></span>|
|[<span data-ttu-id="29088-275">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-276">ReadItem</span></span>|
|[<span data-ttu-id="29088-277">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-278">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="29088-279">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="29088-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="29088-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-282">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-282">Type:</span></span>

*   <span data-ttu-id="29088-283">Date</span><span class="sxs-lookup"><span data-stu-id="29088-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-284">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-284">Requirements</span></span>

|<span data-ttu-id="29088-285">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-285">Requirement</span></span>| <span data-ttu-id="29088-286">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-287">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-288">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-288">1.0</span></span>|
|[<span data-ttu-id="29088-289">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-290">ReadItem</span></span>|
|[<span data-ttu-id="29088-291">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-292">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-293">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-293">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="29088-294">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="29088-294">dateTimeModified :Date</span></span>

<span data-ttu-id="29088-p110">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-297">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="29088-297">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-298">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-298">Type:</span></span>

*   <span data-ttu-id="29088-299">Date</span><span class="sxs-lookup"><span data-stu-id="29088-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-300">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-300">Requirements</span></span>

|<span data-ttu-id="29088-301">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-301">Requirement</span></span>| <span data-ttu-id="29088-302">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-303">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-304">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-304">1.0</span></span>|
|[<span data-ttu-id="29088-305">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-306">ReadItem</span></span>|
|[<span data-ttu-id="29088-307">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-308">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-309">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-309">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="29088-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="29088-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="29088-311">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="29088-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="29088-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="29088-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="29088-314">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="29088-314">Read mode</span></span>

<span data-ttu-id="29088-315">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="29088-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="29088-316">Mode composition</span><span class="sxs-lookup"><span data-stu-id="29088-316">Compose mode</span></span>

<span data-ttu-id="29088-317">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="29088-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="29088-318">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="29088-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-319">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-319">Type:</span></span>

*   <span data-ttu-id="29088-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="29088-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-321">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-321">Requirements</span></span>

|<span data-ttu-id="29088-322">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-322">Requirement</span></span>| <span data-ttu-id="29088-323">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-324">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-325">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-325">1.0</span></span>|
|[<span data-ttu-id="29088-326">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-327">ReadItem</span></span>|
|[<span data-ttu-id="29088-328">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-329">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-330">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-330">Example</span></span>

<span data-ttu-id="29088-331">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="29088-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="29088-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="29088-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="29088-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="29088-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="29088-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-337">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="29088-337">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-338">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-338">Type:</span></span>

*   [<span data-ttu-id="29088-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="29088-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="29088-340">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-340">Requirements</span></span>

|<span data-ttu-id="29088-341">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-341">Requirement</span></span>| <span data-ttu-id="29088-342">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-343">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-344">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-344">1.0</span></span>|
|[<span data-ttu-id="29088-345">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-346">ReadItem</span></span>|
|[<span data-ttu-id="29088-347">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-348">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="29088-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="29088-349">internetMessageId :String</span></span>

<span data-ttu-id="29088-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-352">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-352">Type:</span></span>

*   <span data-ttu-id="29088-353">Chaîne</span><span class="sxs-lookup"><span data-stu-id="29088-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-354">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-354">Requirements</span></span>

|<span data-ttu-id="29088-355">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-355">Requirement</span></span>| <span data-ttu-id="29088-356">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-357">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-358">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-358">1.0</span></span>|
|[<span data-ttu-id="29088-359">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-360">ReadItem</span></span>|
|[<span data-ttu-id="29088-361">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-362">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-363">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-363">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="29088-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="29088-364">itemClass :String</span></span>

<span data-ttu-id="29088-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="29088-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="29088-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="29088-369">Type</span><span class="sxs-lookup"><span data-stu-id="29088-369">Type</span></span> | <span data-ttu-id="29088-370">Description</span><span class="sxs-lookup"><span data-stu-id="29088-370">Description</span></span> | <span data-ttu-id="29088-371">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="29088-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="29088-372">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="29088-372">Appointment items</span></span> | <span data-ttu-id="29088-373">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="29088-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="29088-374">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="29088-374">Message items</span></span> | <span data-ttu-id="29088-375">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="29088-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="29088-376">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="29088-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-377">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-377">Type:</span></span>

*   <span data-ttu-id="29088-378">Chaîne</span><span class="sxs-lookup"><span data-stu-id="29088-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-379">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-379">Requirements</span></span>

|<span data-ttu-id="29088-380">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-380">Requirement</span></span>| <span data-ttu-id="29088-381">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-382">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-383">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-383">1.0</span></span>|
|[<span data-ttu-id="29088-384">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-385">ReadItem</span></span>|
|[<span data-ttu-id="29088-386">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-387">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-388">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-388">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="29088-389">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="29088-389">(nullable) itemId :String</span></span>

<span data-ttu-id="29088-p117">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-392">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="29088-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="29088-393">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="29088-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="29088-394">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="29088-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="29088-395">Pour plus d’informations, consultez la rubrique [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="29088-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="29088-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-398">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-398">Type:</span></span>

*   <span data-ttu-id="29088-399">Chaîne</span><span class="sxs-lookup"><span data-stu-id="29088-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-400">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-400">Requirements</span></span>

|<span data-ttu-id="29088-401">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-401">Requirement</span></span>| <span data-ttu-id="29088-402">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-403">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-404">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-404">1.0</span></span>|
|[<span data-ttu-id="29088-405">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-406">ReadItem</span></span>|
|[<span data-ttu-id="29088-407">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-408">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-409">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-409">Example</span></span>

<span data-ttu-id="29088-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="29088-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="29088-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="29088-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="29088-413">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="29088-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="29088-414">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="29088-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-415">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-415">Type:</span></span>

*   [<span data-ttu-id="29088-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="29088-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="29088-417">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-417">Requirements</span></span>

|<span data-ttu-id="29088-418">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-418">Requirement</span></span>| <span data-ttu-id="29088-419">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-420">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-421">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-421">1.0</span></span>|
|[<span data-ttu-id="29088-422">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-423">ReadItem</span></span>|
|[<span data-ttu-id="29088-424">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-425">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-426">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-426">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="29088-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="29088-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="29088-428">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="29088-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="29088-429">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="29088-429">Read mode</span></span>

<span data-ttu-id="29088-430">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="29088-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="29088-431">Mode composition</span><span class="sxs-lookup"><span data-stu-id="29088-431">Compose mode</span></span>

<span data-ttu-id="29088-432">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="29088-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-433">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-433">Type:</span></span>

*   <span data-ttu-id="29088-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="29088-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-435">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-435">Requirements</span></span>

|<span data-ttu-id="29088-436">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-436">Requirement</span></span>| <span data-ttu-id="29088-437">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-438">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-439">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-439">1.0</span></span>|
|[<span data-ttu-id="29088-440">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-441">ReadItem</span></span>|
|[<span data-ttu-id="29088-442">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-443">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-444">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-444">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="29088-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="29088-445">normalizedSubject :String</span></span>

<span data-ttu-id="29088-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="29088-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="29088-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-450">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-450">Type:</span></span>

*   <span data-ttu-id="29088-451">Chaîne</span><span class="sxs-lookup"><span data-stu-id="29088-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-452">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-452">Requirements</span></span>

|<span data-ttu-id="29088-453">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-453">Requirement</span></span>| <span data-ttu-id="29088-454">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-455">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-456">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-456">1.0</span></span>|
|[<span data-ttu-id="29088-457">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-458">ReadItem</span></span>|
|[<span data-ttu-id="29088-459">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-460">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-461">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-461">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="29088-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="29088-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="29088-463">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="29088-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-464">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-464">Type:</span></span>

*   [<span data-ttu-id="29088-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="29088-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="29088-466">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-466">Requirements</span></span>

|<span data-ttu-id="29088-467">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-467">Requirement</span></span>| <span data-ttu-id="29088-468">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-469">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-470">1.3</span><span class="sxs-lookup"><span data-stu-id="29088-470">1.3</span></span>|
|[<span data-ttu-id="29088-471">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-472">ReadItem</span></span>|
|[<span data-ttu-id="29088-473">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-474">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="29088-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="29088-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="29088-476">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="29088-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="29088-477">Le type d’objet et le niveau d’accès varie selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="29088-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="29088-478">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-478">Read mode</span></span>

<span data-ttu-id="29088-479">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="29088-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="29088-480">Mode composition</span><span class="sxs-lookup"><span data-stu-id="29088-480">Compose mode</span></span>

<span data-ttu-id="29088-481">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="29088-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-482">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-482">Type:</span></span>

*   <span data-ttu-id="29088-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="29088-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-484">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-484">Requirements</span></span>

|<span data-ttu-id="29088-485">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-485">Requirement</span></span>| <span data-ttu-id="29088-486">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-487">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-488">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-488">1.0</span></span>|
|[<span data-ttu-id="29088-489">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-490">ReadItem</span></span>|
|[<span data-ttu-id="29088-491">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-492">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-493">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-493">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="29088-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="29088-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="29088-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-497">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-497">Type:</span></span>

*   [<span data-ttu-id="29088-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="29088-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="29088-499">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-499">Requirements</span></span>

|<span data-ttu-id="29088-500">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-500">Requirement</span></span>| <span data-ttu-id="29088-501">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-502">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-503">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-503">1.0</span></span>|
|[<span data-ttu-id="29088-504">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-505">ReadItem</span></span>|
|[<span data-ttu-id="29088-506">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-507">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-508">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-508">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="29088-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="29088-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="29088-510">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="29088-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="29088-511">Le type d’objet et le niveau d’accès varie selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="29088-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="29088-512">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-512">Read mode</span></span>

<span data-ttu-id="29088-513">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="29088-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="29088-514">Mode composition</span><span class="sxs-lookup"><span data-stu-id="29088-514">Compose mode</span></span>

<span data-ttu-id="29088-515">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="29088-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-516">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-516">Type:</span></span>

*   <span data-ttu-id="29088-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="29088-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-518">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-518">Requirements</span></span>

|<span data-ttu-id="29088-519">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-519">Requirement</span></span>| <span data-ttu-id="29088-520">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-521">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-522">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-522">1.0</span></span>|
|[<span data-ttu-id="29088-523">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-524">ReadItem</span></span>|
|[<span data-ttu-id="29088-525">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-526">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-527">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-527">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="29088-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="29088-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="29088-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="29088-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="29088-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="29088-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-533">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="29088-533">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-534">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-534">Type:</span></span>

*   [<span data-ttu-id="29088-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="29088-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="29088-536">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-536">Requirements</span></span>

|<span data-ttu-id="29088-537">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-537">Requirement</span></span>| <span data-ttu-id="29088-538">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-539">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-540">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-540">1.0</span></span>|
|[<span data-ttu-id="29088-541">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-542">ReadItem</span></span>|
|[<span data-ttu-id="29088-543">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-544">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-545">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-545">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="29088-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="29088-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="29088-547">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="29088-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="29088-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="29088-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="29088-550">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="29088-550">Read mode</span></span>

<span data-ttu-id="29088-551">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="29088-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="29088-552">Mode composition</span><span class="sxs-lookup"><span data-stu-id="29088-552">Compose mode</span></span>

<span data-ttu-id="29088-553">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="29088-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="29088-554">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="29088-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-555">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-555">Type:</span></span>

*   <span data-ttu-id="29088-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="29088-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-557">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-557">Requirements</span></span>

|<span data-ttu-id="29088-558">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-558">Requirement</span></span>| <span data-ttu-id="29088-559">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-560">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-560">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-561">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-561">1.0</span></span>|
|[<span data-ttu-id="29088-562">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-563">ReadItem</span></span>|
|[<span data-ttu-id="29088-564">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-565">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-566">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-566">Example</span></span>

<span data-ttu-id="29088-567">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="29088-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="29088-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="29088-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="29088-569">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="29088-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="29088-570">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="29088-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="29088-571">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="29088-571">Read mode</span></span>

<span data-ttu-id="29088-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="29088-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="29088-574">Mode composition</span><span class="sxs-lookup"><span data-stu-id="29088-574">Compose mode</span></span>

<span data-ttu-id="29088-575">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="29088-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="29088-576">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-576">Type:</span></span>

*   <span data-ttu-id="29088-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="29088-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-578">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-578">Requirements</span></span>

|<span data-ttu-id="29088-579">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-579">Requirement</span></span>| <span data-ttu-id="29088-580">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-581">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-582">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-582">1.0</span></span>|
|[<span data-ttu-id="29088-583">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-584">ReadItem</span></span>|
|[<span data-ttu-id="29088-585">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-586">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="29088-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="29088-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="29088-588">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="29088-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="29088-589">Le type d’objet et le niveau d’accès varie selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="29088-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="29088-590">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-590">Read mode</span></span>

<span data-ttu-id="29088-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="29088-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="29088-593">Mode composition</span><span class="sxs-lookup"><span data-stu-id="29088-593">Compose mode</span></span>

<span data-ttu-id="29088-594">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="29088-594">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="29088-595">Type :</span><span class="sxs-lookup"><span data-stu-id="29088-595">Type:</span></span>

*   <span data-ttu-id="29088-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="29088-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-597">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-597">Requirements</span></span>

|<span data-ttu-id="29088-598">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-598">Requirement</span></span>| <span data-ttu-id="29088-599">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-600">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-600">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-601">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-601">1.0</span></span>|
|[<span data-ttu-id="29088-602">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-603">ReadItem</span></span>|
|[<span data-ttu-id="29088-604">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-605">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-606">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-606">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="29088-607">Méthodes</span><span class="sxs-lookup"><span data-stu-id="29088-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="29088-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="29088-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="29088-609">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="29088-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="29088-610">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="29088-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="29088-611">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="29088-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-612">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-612">Parameters:</span></span>

|<span data-ttu-id="29088-613">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-613">Name</span></span>| <span data-ttu-id="29088-614">Type</span><span class="sxs-lookup"><span data-stu-id="29088-614">Type</span></span>| <span data-ttu-id="29088-615">Attributs</span><span class="sxs-lookup"><span data-stu-id="29088-615">Attributes</span></span>| <span data-ttu-id="29088-616">Description</span><span class="sxs-lookup"><span data-stu-id="29088-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="29088-617">String</span><span class="sxs-lookup"><span data-stu-id="29088-617">String</span></span>||<span data-ttu-id="29088-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="29088-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="29088-620">String</span><span class="sxs-lookup"><span data-stu-id="29088-620">String</span></span>||<span data-ttu-id="29088-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="29088-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="29088-623">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-623">Object</span></span>| <span data-ttu-id="29088-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-624">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-625">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="29088-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="29088-626">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-626">Object</span></span> | <span data-ttu-id="29088-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-627">&lt;optional&gt;</span></span> | <span data-ttu-id="29088-628">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="29088-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="29088-629">Boolean</span></span> | <span data-ttu-id="29088-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-630">&lt;optional&gt;</span></span> | <span data-ttu-id="29088-631">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="29088-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="29088-632">fonction</span><span class="sxs-lookup"><span data-stu-id="29088-632">function</span></span>| <span data-ttu-id="29088-633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-633">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-634">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29088-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="29088-635">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="29088-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="29088-636">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="29088-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="29088-637">Erreurs</span><span class="sxs-lookup"><span data-stu-id="29088-637">Errors</span></span>

| <span data-ttu-id="29088-638">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="29088-638">Error code</span></span> | <span data-ttu-id="29088-639">Description</span><span class="sxs-lookup"><span data-stu-id="29088-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="29088-640">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="29088-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="29088-641">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="29088-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="29088-642">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="29088-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="29088-643">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-643">Requirements</span></span>

|<span data-ttu-id="29088-644">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-644">Requirement</span></span>| <span data-ttu-id="29088-645">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-646">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-647">1.1</span><span class="sxs-lookup"><span data-stu-id="29088-647">1.1</span></span>|
|[<span data-ttu-id="29088-648">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="29088-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="29088-650">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-651">Composition</span><span class="sxs-lookup"><span data-stu-id="29088-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="29088-652">Exemples</span><span class="sxs-lookup"><span data-stu-id="29088-652">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="29088-653">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="29088-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
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
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="29088-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="29088-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="29088-655">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="29088-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="29088-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="29088-659">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="29088-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="29088-660">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="29088-660">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-661">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-661">Parameters:</span></span>

|<span data-ttu-id="29088-662">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-662">Name</span></span>| <span data-ttu-id="29088-663">Type</span><span class="sxs-lookup"><span data-stu-id="29088-663">Type</span></span>| <span data-ttu-id="29088-664">Attributs</span><span class="sxs-lookup"><span data-stu-id="29088-664">Attributes</span></span>| <span data-ttu-id="29088-665">Description</span><span class="sxs-lookup"><span data-stu-id="29088-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="29088-666">String</span><span class="sxs-lookup"><span data-stu-id="29088-666">String</span></span>||<span data-ttu-id="29088-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="29088-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="29088-669">String</span><span class="sxs-lookup"><span data-stu-id="29088-669">String</span></span>||<span data-ttu-id="29088-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="29088-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="29088-672">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-672">Object</span></span>| <span data-ttu-id="29088-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-673">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-674">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="29088-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="29088-675">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-675">Object</span></span>| <span data-ttu-id="29088-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-676">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-677">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="29088-678">fonction</span><span class="sxs-lookup"><span data-stu-id="29088-678">function</span></span>| <span data-ttu-id="29088-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-679">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-680">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29088-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="29088-681">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="29088-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="29088-682">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="29088-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="29088-683">Erreurs</span><span class="sxs-lookup"><span data-stu-id="29088-683">Errors</span></span>

| <span data-ttu-id="29088-684">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="29088-684">Error code</span></span> | <span data-ttu-id="29088-685">Description</span><span class="sxs-lookup"><span data-stu-id="29088-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="29088-686">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="29088-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="29088-687">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-687">Requirements</span></span>

|<span data-ttu-id="29088-688">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-688">Requirement</span></span>| <span data-ttu-id="29088-689">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-690">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-691">1.1</span><span class="sxs-lookup"><span data-stu-id="29088-691">1.1</span></span>|
|[<span data-ttu-id="29088-692">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="29088-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="29088-694">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-695">Composition</span><span class="sxs-lookup"><span data-stu-id="29088-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-696">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-696">Example</span></span>

<span data-ttu-id="29088-697">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="29088-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="29088-698">close()</span><span class="sxs-lookup"><span data-stu-id="29088-698">close()</span></span>

<span data-ttu-id="29088-699">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="29088-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="29088-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="29088-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-702">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="29088-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="29088-703">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="29088-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-704">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-704">Requirements</span></span>

|<span data-ttu-id="29088-705">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-705">Requirement</span></span>| <span data-ttu-id="29088-706">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-707">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-708">1.3</span><span class="sxs-lookup"><span data-stu-id="29088-708">1.3</span></span>|
|[<span data-ttu-id="29088-709">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-710">Restreinte</span><span class="sxs-lookup"><span data-stu-id="29088-710">Restricted</span></span>|
|[<span data-ttu-id="29088-711">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-712">Composition</span><span class="sxs-lookup"><span data-stu-id="29088-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="29088-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="29088-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="29088-714">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="29088-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-715">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="29088-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="29088-716">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="29088-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="29088-717">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="29088-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="29088-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="29088-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-721">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-721">Parameters:</span></span>

| <span data-ttu-id="29088-722">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-722">Name</span></span> | <span data-ttu-id="29088-723">Type</span><span class="sxs-lookup"><span data-stu-id="29088-723">Type</span></span> | <span data-ttu-id="29088-724">Attributs</span><span class="sxs-lookup"><span data-stu-id="29088-724">Attributes</span></span> | <span data-ttu-id="29088-725">Description</span><span class="sxs-lookup"><span data-stu-id="29088-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="29088-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="29088-726">String &#124; Object</span></span>| |<span data-ttu-id="29088-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="29088-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="29088-729">**OU**</span><span class="sxs-lookup"><span data-stu-id="29088-729">**OR**</span></span><br/><span data-ttu-id="29088-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="29088-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="29088-732">String</span><span class="sxs-lookup"><span data-stu-id="29088-732">String</span></span> | <span data-ttu-id="29088-733">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-733">&lt;optional&gt;</span></span> | <span data-ttu-id="29088-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="29088-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="29088-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="29088-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-737">&lt;optional&gt;</span></span> | <span data-ttu-id="29088-738">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="29088-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="29088-739">Chaîne</span><span class="sxs-lookup"><span data-stu-id="29088-739">String</span></span> | | <span data-ttu-id="29088-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="29088-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="29088-742">String</span><span class="sxs-lookup"><span data-stu-id="29088-742">String</span></span> | | <span data-ttu-id="29088-743">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="29088-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="29088-744">Chaîne</span><span class="sxs-lookup"><span data-stu-id="29088-744">String</span></span> | | <span data-ttu-id="29088-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="29088-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="29088-747">Booléen</span><span class="sxs-lookup"><span data-stu-id="29088-747">Boolean</span></span> | | <span data-ttu-id="29088-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="29088-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="29088-750">String</span><span class="sxs-lookup"><span data-stu-id="29088-750">String</span></span> | | <span data-ttu-id="29088-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="29088-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="29088-754">function</span><span class="sxs-lookup"><span data-stu-id="29088-754">function</span></span> | <span data-ttu-id="29088-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-755">&lt;optional&gt;</span></span> | <span data-ttu-id="29088-756">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29088-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="29088-757">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-757">Requirements</span></span>

|<span data-ttu-id="29088-758">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-758">Requirement</span></span>| <span data-ttu-id="29088-759">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-760">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-761">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-761">1.0</span></span>|
|[<span data-ttu-id="29088-762">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-763">ReadItem</span></span>|
|[<span data-ttu-id="29088-764">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-765">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="29088-766">Exemples</span><span class="sxs-lookup"><span data-stu-id="29088-766">Examples</span></span>

<span data-ttu-id="29088-767">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="29088-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="29088-768">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="29088-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="29088-769">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="29088-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="29088-770">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="29088-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="29088-771">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="29088-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="29088-772">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="29088-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="29088-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="29088-774">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="29088-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-775">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="29088-775">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="29088-776">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="29088-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="29088-777">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="29088-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="29088-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="29088-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-781">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-781">Parameters:</span></span>

| <span data-ttu-id="29088-782">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-782">Name</span></span> | <span data-ttu-id="29088-783">Type</span><span class="sxs-lookup"><span data-stu-id="29088-783">Type</span></span> | <span data-ttu-id="29088-784">Attributs</span><span class="sxs-lookup"><span data-stu-id="29088-784">Attributes</span></span> | <span data-ttu-id="29088-785">Description</span><span class="sxs-lookup"><span data-stu-id="29088-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="29088-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="29088-786">String &#124; Object</span></span>| | <span data-ttu-id="29088-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="29088-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="29088-789">**OU**</span><span class="sxs-lookup"><span data-stu-id="29088-789">**OR**</span></span><br/><span data-ttu-id="29088-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="29088-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="29088-792">String</span><span class="sxs-lookup"><span data-stu-id="29088-792">String</span></span> | <span data-ttu-id="29088-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-793">&lt;optional&gt;</span></span> | <span data-ttu-id="29088-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="29088-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="29088-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="29088-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-797">&lt;optional&gt;</span></span> | <span data-ttu-id="29088-798">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="29088-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="29088-799">Chaîne</span><span class="sxs-lookup"><span data-stu-id="29088-799">String</span></span> | | <span data-ttu-id="29088-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="29088-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="29088-802">String</span><span class="sxs-lookup"><span data-stu-id="29088-802">String</span></span> | | <span data-ttu-id="29088-803">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="29088-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="29088-804">Chaîne</span><span class="sxs-lookup"><span data-stu-id="29088-804">String</span></span> | | <span data-ttu-id="29088-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="29088-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="29088-807">Booléen</span><span class="sxs-lookup"><span data-stu-id="29088-807">Boolean</span></span> | | <span data-ttu-id="29088-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="29088-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="29088-810">String</span><span class="sxs-lookup"><span data-stu-id="29088-810">String</span></span> | | <span data-ttu-id="29088-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="29088-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="29088-814">function</span><span class="sxs-lookup"><span data-stu-id="29088-814">function</span></span> | <span data-ttu-id="29088-815">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-815">&lt;optional&gt;</span></span> | <span data-ttu-id="29088-816">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29088-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="29088-817">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-817">Requirements</span></span>

|<span data-ttu-id="29088-818">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-818">Requirement</span></span>| <span data-ttu-id="29088-819">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-820">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-821">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-821">1.0</span></span>|
|[<span data-ttu-id="29088-822">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-823">ReadItem</span></span>|
|[<span data-ttu-id="29088-824">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-825">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="29088-826">Exemples</span><span class="sxs-lookup"><span data-stu-id="29088-826">Examples</span></span>

<span data-ttu-id="29088-827">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="29088-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="29088-828">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="29088-828">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="29088-829">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="29088-829">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="29088-830">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="29088-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="29088-831">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="29088-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="29088-832">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="29088-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="29088-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="29088-834">Obtient les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="29088-834">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-835">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="29088-835">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-836">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-836">Requirements</span></span>

|<span data-ttu-id="29088-837">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-837">Requirement</span></span>| <span data-ttu-id="29088-838">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-839">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-840">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-840">1.0</span></span>|
|[<span data-ttu-id="29088-841">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-842">ReadItem</span></span>|
|[<span data-ttu-id="29088-843">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-844">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="29088-845">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="29088-845">Returns:</span></span>

<span data-ttu-id="29088-846">Type : [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="29088-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="29088-847">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-847">Example</span></span>

<span data-ttu-id="29088-848">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="29088-848">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="29088-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="29088-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="29088-850">Obtient un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="29088-850">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-851">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="29088-851">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-852">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-852">Parameters:</span></span>

|<span data-ttu-id="29088-853">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-853">Name</span></span>| <span data-ttu-id="29088-854">Type</span><span class="sxs-lookup"><span data-stu-id="29088-854">Type</span></span>| <span data-ttu-id="29088-855">Description</span><span class="sxs-lookup"><span data-stu-id="29088-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="29088-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="29088-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="29088-857">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="29088-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29088-858">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-858">Requirements</span></span>

|<span data-ttu-id="29088-859">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-859">Requirement</span></span>| <span data-ttu-id="29088-860">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-861">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-862">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-862">1.0</span></span>|
|[<span data-ttu-id="29088-863">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-864">Restreinte</span><span class="sxs-lookup"><span data-stu-id="29088-864">Restricted</span></span>|
|[<span data-ttu-id="29088-865">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-866">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="29088-867">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="29088-867">Returns:</span></span>

<span data-ttu-id="29088-868">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="29088-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="29088-869">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="29088-869">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="29088-870">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="29088-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="29088-871">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="29088-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="29088-872">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="29088-872">Value of `entityType`</span></span> | <span data-ttu-id="29088-873">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="29088-873">Type of objects in returned array</span></span> | <span data-ttu-id="29088-874">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="29088-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="29088-875">String</span><span class="sxs-lookup"><span data-stu-id="29088-875">String</span></span> | <span data-ttu-id="29088-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="29088-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="29088-877">Contact</span><span class="sxs-lookup"><span data-stu-id="29088-877">Contact</span></span> | <span data-ttu-id="29088-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="29088-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="29088-879">String</span><span class="sxs-lookup"><span data-stu-id="29088-879">String</span></span> | <span data-ttu-id="29088-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="29088-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="29088-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="29088-881">MeetingSuggestion</span></span> | <span data-ttu-id="29088-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="29088-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="29088-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="29088-883">PhoneNumber</span></span> | <span data-ttu-id="29088-884">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="29088-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="29088-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="29088-885">TaskSuggestion</span></span> | <span data-ttu-id="29088-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="29088-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="29088-887">String</span><span class="sxs-lookup"><span data-stu-id="29088-887">String</span></span> | <span data-ttu-id="29088-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="29088-888">**Restricted**</span></span> |

<span data-ttu-id="29088-889">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="29088-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="29088-890">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-890">Example</span></span>

<span data-ttu-id="29088-891">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="29088-891">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="29088-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="29088-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="29088-893">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="29088-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-894">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="29088-894">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="29088-895">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="29088-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-896">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-896">Parameters:</span></span>

|<span data-ttu-id="29088-897">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-897">Name</span></span>| <span data-ttu-id="29088-898">Type</span><span class="sxs-lookup"><span data-stu-id="29088-898">Type</span></span>| <span data-ttu-id="29088-899">Description</span><span class="sxs-lookup"><span data-stu-id="29088-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="29088-900">String</span><span class="sxs-lookup"><span data-stu-id="29088-900">String</span></span>|<span data-ttu-id="29088-901">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="29088-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29088-902">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-902">Requirements</span></span>

|<span data-ttu-id="29088-903">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-903">Requirement</span></span>| <span data-ttu-id="29088-904">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-905">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-906">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-906">1.0</span></span>|
|[<span data-ttu-id="29088-907">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-908">ReadItem</span></span>|
|[<span data-ttu-id="29088-909">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-910">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="29088-911">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="29088-911">Returns:</span></span>

<span data-ttu-id="29088-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="29088-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="29088-914">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="29088-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="29088-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="29088-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="29088-916">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="29088-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-917">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="29088-917">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="29088-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="29088-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="29088-921">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="29088-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="29088-922">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="29088-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="29088-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="29088-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="29088-926">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-926">Requirements</span></span>

|<span data-ttu-id="29088-927">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-927">Requirement</span></span>| <span data-ttu-id="29088-928">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-929">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-930">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-930">1.0</span></span>|
|[<span data-ttu-id="29088-931">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-932">ReadItem</span></span>|
|[<span data-ttu-id="29088-933">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-934">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="29088-935">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="29088-935">Returns:</span></span>

<span data-ttu-id="29088-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="29088-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="29088-938">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="29088-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="29088-939">Object</span><span class="sxs-lookup"><span data-stu-id="29088-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="29088-940">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-940">Example</span></span>

<span data-ttu-id="29088-941">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="29088-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="29088-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="29088-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="29088-943">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="29088-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-944">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="29088-944">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="29088-945">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="29088-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="29088-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="29088-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-948">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-948">Parameters:</span></span>

|<span data-ttu-id="29088-949">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-949">Name</span></span>| <span data-ttu-id="29088-950">Type</span><span class="sxs-lookup"><span data-stu-id="29088-950">Type</span></span>| <span data-ttu-id="29088-951">Description</span><span class="sxs-lookup"><span data-stu-id="29088-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="29088-952">String</span><span class="sxs-lookup"><span data-stu-id="29088-952">String</span></span>|<span data-ttu-id="29088-953">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="29088-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29088-954">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-954">Requirements</span></span>

|<span data-ttu-id="29088-955">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-955">Requirement</span></span>| <span data-ttu-id="29088-956">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-957">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-958">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-958">1.0</span></span>|
|[<span data-ttu-id="29088-959">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-960">ReadItem</span></span>|
|[<span data-ttu-id="29088-961">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-962">Lecture</span><span class="sxs-lookup"><span data-stu-id="29088-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="29088-963">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="29088-963">Returns:</span></span>

<span data-ttu-id="29088-964">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="29088-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="29088-965">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="29088-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="29088-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="29088-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="29088-967">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-967">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="29088-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="29088-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="29088-969">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="29088-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="29088-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="29088-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-972">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-972">Parameters:</span></span>

|<span data-ttu-id="29088-973">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-973">Name</span></span>| <span data-ttu-id="29088-974">Type</span><span class="sxs-lookup"><span data-stu-id="29088-974">Type</span></span>| <span data-ttu-id="29088-975">Attributs</span><span class="sxs-lookup"><span data-stu-id="29088-975">Attributes</span></span>| <span data-ttu-id="29088-976">Description</span><span class="sxs-lookup"><span data-stu-id="29088-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="29088-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="29088-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="29088-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="29088-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="29088-981">Object</span><span class="sxs-lookup"><span data-stu-id="29088-981">Object</span></span>| <span data-ttu-id="29088-982">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-982">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-983">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="29088-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="29088-984">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-984">Object</span></span>| <span data-ttu-id="29088-985">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-985">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-986">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="29088-987">fonction</span><span class="sxs-lookup"><span data-stu-id="29088-987">function</span></span>||<span data-ttu-id="29088-988">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29088-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="29088-989">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="29088-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="29088-990">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="29088-990">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29088-991">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-991">Requirements</span></span>

|<span data-ttu-id="29088-992">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-992">Requirement</span></span>| <span data-ttu-id="29088-993">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-994">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-995">1.2</span><span class="sxs-lookup"><span data-stu-id="29088-995">1.2</span></span>|
|[<span data-ttu-id="29088-996">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="29088-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="29088-998">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-999">Composition</span><span class="sxs-lookup"><span data-stu-id="29088-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="29088-1000">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="29088-1000">Returns:</span></span>

<span data-ttu-id="29088-1001">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="29088-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="29088-1002">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="29088-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="29088-1003">String</span><span class="sxs-lookup"><span data-stu-id="29088-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="29088-1004">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-1004">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="29088-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="29088-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="29088-1006">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="29088-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="29088-p163">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="29088-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-1010">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-1010">Parameters:</span></span>

|<span data-ttu-id="29088-1011">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-1011">Name</span></span>| <span data-ttu-id="29088-1012">Type</span><span class="sxs-lookup"><span data-stu-id="29088-1012">Type</span></span>| <span data-ttu-id="29088-1013">Attributs</span><span class="sxs-lookup"><span data-stu-id="29088-1013">Attributes</span></span>| <span data-ttu-id="29088-1014">Description</span><span class="sxs-lookup"><span data-stu-id="29088-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="29088-1015">function</span><span class="sxs-lookup"><span data-stu-id="29088-1015">function</span></span>||<span data-ttu-id="29088-1016">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29088-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="29088-1017">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="29088-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="29088-1018">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="29088-1018">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="29088-1019">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-1019">Object</span></span>| <span data-ttu-id="29088-1020">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-1021">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-1021">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="29088-1022">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29088-1023">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-1023">Requirements</span></span>

|<span data-ttu-id="29088-1024">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-1024">Requirement</span></span>| <span data-ttu-id="29088-1025">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-1026">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="29088-1027">1.0</span></span>|
|[<span data-ttu-id="29088-1028">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29088-1029">ReadItem</span></span>|
|[<span data-ttu-id="29088-1030">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-1031">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="29088-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-1032">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-1032">Example</span></span>

<span data-ttu-id="29088-p166">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="29088-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="29088-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="29088-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="29088-1037">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="29088-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="29088-p167">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="29088-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-1042">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-1042">Parameters:</span></span>

|<span data-ttu-id="29088-1043">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-1043">Name</span></span>| <span data-ttu-id="29088-1044">Type</span><span class="sxs-lookup"><span data-stu-id="29088-1044">Type</span></span>| <span data-ttu-id="29088-1045">Attributs</span><span class="sxs-lookup"><span data-stu-id="29088-1045">Attributes</span></span>| <span data-ttu-id="29088-1046">Description</span><span class="sxs-lookup"><span data-stu-id="29088-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="29088-1047">String</span><span class="sxs-lookup"><span data-stu-id="29088-1047">String</span></span>||<span data-ttu-id="29088-p168">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="29088-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="29088-1050">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-1050">Object</span></span>| <span data-ttu-id="29088-1051">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-1052">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="29088-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="29088-1053">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-1053">Object</span></span>| <span data-ttu-id="29088-1054">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-1055">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="29088-1056">fonction</span><span class="sxs-lookup"><span data-stu-id="29088-1056">function</span></span>| <span data-ttu-id="29088-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-1058">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29088-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="29088-1059">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="29088-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="29088-1060">Erreurs</span><span class="sxs-lookup"><span data-stu-id="29088-1060">Errors</span></span>

| <span data-ttu-id="29088-1061">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="29088-1061">Error code</span></span> | <span data-ttu-id="29088-1062">Description</span><span class="sxs-lookup"><span data-stu-id="29088-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="29088-1063">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="29088-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="29088-1064">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-1064">Requirements</span></span>

|<span data-ttu-id="29088-1065">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-1065">Requirement</span></span>| <span data-ttu-id="29088-1066">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-1067">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="29088-1068">1.1</span></span>|
|[<span data-ttu-id="29088-1069">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="29088-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="29088-1071">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-1072">Composition</span><span class="sxs-lookup"><span data-stu-id="29088-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-1073">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-1073">Example</span></span>

<span data-ttu-id="29088-1074">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="29088-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="29088-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="29088-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="29088-1076">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="29088-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="29088-p169">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="29088-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-1080">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="29088-1080">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="29088-1081">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="29088-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="29088-p171">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="29088-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="29088-1085">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="29088-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="29088-1086">Outlook pour Mac ne prend en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="29088-1086">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="29088-1087">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="29088-1087">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="29088-1088">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="29088-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-1089">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-1089">Parameters:</span></span>

|<span data-ttu-id="29088-1090">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-1090">Name</span></span>| <span data-ttu-id="29088-1091">Type</span><span class="sxs-lookup"><span data-stu-id="29088-1091">Type</span></span>| <span data-ttu-id="29088-1092">Attributs</span><span class="sxs-lookup"><span data-stu-id="29088-1092">Attributes</span></span>| <span data-ttu-id="29088-1093">Description</span><span class="sxs-lookup"><span data-stu-id="29088-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="29088-1094">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-1094">Object</span></span>| <span data-ttu-id="29088-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-1096">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="29088-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="29088-1097">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-1097">Object</span></span>| <span data-ttu-id="29088-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-1099">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="29088-1100">fonction</span><span class="sxs-lookup"><span data-stu-id="29088-1100">function</span></span>||<span data-ttu-id="29088-1101">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29088-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="29088-1102">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="29088-1102">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29088-1103">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-1103">Requirements</span></span>

|<span data-ttu-id="29088-1104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-1104">Requirement</span></span>| <span data-ttu-id="29088-1105">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-1106">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="29088-1107">1.3</span></span>|
|[<span data-ttu-id="29088-1108">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="29088-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="29088-1110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-1111">Composition</span><span class="sxs-lookup"><span data-stu-id="29088-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="29088-1112">範例</span><span class="sxs-lookup"><span data-stu-id="29088-1112">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="29088-p173">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="29088-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="29088-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="29088-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="29088-1116">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="29088-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="29088-p174">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="29088-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29088-1120">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="29088-1120">Parameters:</span></span>

|<span data-ttu-id="29088-1121">Nom</span><span class="sxs-lookup"><span data-stu-id="29088-1121">Name</span></span>| <span data-ttu-id="29088-1122">Type</span><span class="sxs-lookup"><span data-stu-id="29088-1122">Type</span></span>| <span data-ttu-id="29088-1123">Attributs</span><span class="sxs-lookup"><span data-stu-id="29088-1123">Attributes</span></span>| <span data-ttu-id="29088-1124">Description</span><span class="sxs-lookup"><span data-stu-id="29088-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="29088-1125">String</span><span class="sxs-lookup"><span data-stu-id="29088-1125">String</span></span>||<span data-ttu-id="29088-p175">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="29088-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="29088-1129">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-1129">Object</span></span>| <span data-ttu-id="29088-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-1131">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="29088-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="29088-1132">Objet</span><span class="sxs-lookup"><span data-stu-id="29088-1132">Object</span></span>| <span data-ttu-id="29088-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-1134">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="29088-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="29088-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="29088-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="29088-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29088-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="29088-p176">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="29088-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="29088-p177">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="29088-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="29088-1141">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="29088-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="29088-1142">fonction</span><span class="sxs-lookup"><span data-stu-id="29088-1142">function</span></span>||<span data-ttu-id="29088-1143">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29088-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="29088-1144">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="29088-1144">Requirements</span></span>

|<span data-ttu-id="29088-1145">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="29088-1145">Requirement</span></span>| <span data-ttu-id="29088-1146">Valeur</span><span class="sxs-lookup"><span data-stu-id="29088-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="29088-1147">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="29088-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29088-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="29088-1148">1.2</span></span>|
|[<span data-ttu-id="29088-1149">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="29088-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29088-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="29088-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="29088-1151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="29088-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29088-1152">Composition</span><span class="sxs-lookup"><span data-stu-id="29088-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="29088-1153">Exemple</span><span class="sxs-lookup"><span data-stu-id="29088-1153">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```