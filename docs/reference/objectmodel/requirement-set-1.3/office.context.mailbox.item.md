
# <a name="item"></a><span data-ttu-id="40310-101">item</span><span class="sxs-lookup"><span data-stu-id="40310-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="40310-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="40310-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="40310-p101">Utiliser l’espace-nom `item` pour accéder a votre message, réunion, demande de réunion ou rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="40310-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-105">Requirements</span></span>

|<span data-ttu-id="40310-106">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-106">Requirement</span></span>| <span data-ttu-id="40310-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-108">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-109">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-109">1.0</span></span>|
|[<span data-ttu-id="40310-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-111">Restreint</span><span class="sxs-lookup"><span data-stu-id="40310-111">Restricted</span></span>|
|[<span data-ttu-id="40310-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="40310-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-114">Example</span></span>

<span data-ttu-id="40310-115">Cet exemple de code JavaScript montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="40310-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
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

### <a name="members"></a><span data-ttu-id="40310-116">Membres</span><span class="sxs-lookup"><span data-stu-id="40310-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="40310-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="40310-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="40310-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-120">Certains types de fichiers sont bloqués par Outlook en raison de problèmes de sécurité potentiels et ne sont donc pas rendus.</span><span class="sxs-lookup"><span data-stu-id="40310-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="40310-121">Pour plus d’information, voir les [pièces jointes bloquées par Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="40310-121">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="40310-122">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-122">Type:</span></span>

*   <span data-ttu-id="40310-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="40310-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-124">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-124">Requirements</span></span>

|<span data-ttu-id="40310-125">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-125">Requirement</span></span>| <span data-ttu-id="40310-126">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-127">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-127">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-128">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-128">1.0</span></span>|
|[<span data-ttu-id="40310-129">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-130">ReadItem</span></span>|
|[<span data-ttu-id="40310-131">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-133">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-133">Example</span></span>

<span data-ttu-id="40310-134">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="40310-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="40310-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="40310-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="40310-136">Obtient un objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires des Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="40310-136">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="40310-137">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-138">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-138">Type:</span></span>

*   [<span data-ttu-id="40310-139">Recipients</span><span class="sxs-lookup"><span data-stu-id="40310-139">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="40310-140">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-140">Requirements</span></span>

|<span data-ttu-id="40310-141">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-141">Requirement</span></span>| <span data-ttu-id="40310-142">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-143">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-143">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-144">1.1</span><span class="sxs-lookup"><span data-stu-id="40310-144">1.1</span></span>|
|[<span data-ttu-id="40310-145">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-146">ReadItem</span></span>|
|[<span data-ttu-id="40310-147">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-148">Composition</span><span class="sxs-lookup"><span data-stu-id="40310-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-149">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="40310-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="40310-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="40310-151">Obtient un objet qui fournit des méthodes permettant de manipuler le texte d’un élément.</span><span class="sxs-lookup"><span data-stu-id="40310-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-152">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-152">Type:</span></span>

*   [<span data-ttu-id="40310-153">Body</span><span class="sxs-lookup"><span data-stu-id="40310-153">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="40310-154">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-154">Requirements</span></span>

|<span data-ttu-id="40310-155">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-155">Requirement</span></span>| <span data-ttu-id="40310-156">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-157">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-157">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-158">1.1</span><span class="sxs-lookup"><span data-stu-id="40310-158">1.1</span></span>|
|[<span data-ttu-id="40310-159">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-160">ReadItem</span></span>|
|[<span data-ttu-id="40310-161">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-162">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="40310-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="40310-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="40310-164">Permet d’accéder aux destinataires Cc (copie carbone) d’un message.</span><span class="sxs-lookup"><span data-stu-id="40310-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="40310-165">Le type d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="40310-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="40310-166">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="40310-166">Read mode</span></span>

<span data-ttu-id="40310-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="40310-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="40310-169">Mode composition</span><span class="sxs-lookup"><span data-stu-id="40310-169">Compose mode</span></span>

<span data-ttu-id="40310-170">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="40310-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-171">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-171">Type:</span></span>

*   <span data-ttu-id="40310-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="40310-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-173">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-173">Requirements</span></span>

|<span data-ttu-id="40310-174">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-174">Requirement</span></span>| <span data-ttu-id="40310-175">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-176">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-176">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-177">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-177">1.0</span></span>|
|[<span data-ttu-id="40310-178">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-179">ReadItem</span></span>|
|[<span data-ttu-id="40310-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-181">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-182">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="40310-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="40310-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="40310-184">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="40310-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="40310-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’identificateur de conversation de ce message changera et la valeur que vous aurez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="40310-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="40310-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renverra une valeur.</span><span class="sxs-lookup"><span data-stu-id="40310-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-189">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-189">Type:</span></span>

*   <span data-ttu-id="40310-190">String</span><span class="sxs-lookup"><span data-stu-id="40310-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-191">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-191">Requirements</span></span>

|<span data-ttu-id="40310-192">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-192">Requirement</span></span>| <span data-ttu-id="40310-193">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-194">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-195">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-195">1.0</span></span>|
|[<span data-ttu-id="40310-196">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-197">ReadItem</span></span>|
|[<span data-ttu-id="40310-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-199">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="40310-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="40310-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="40310-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-203">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-203">Type:</span></span>

*   <span data-ttu-id="40310-204">Date</span><span class="sxs-lookup"><span data-stu-id="40310-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-205">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-205">Requirements</span></span>

|<span data-ttu-id="40310-206">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-206">Requirement</span></span>| <span data-ttu-id="40310-207">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-208">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-208">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-209">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-209">1.0</span></span>|
|[<span data-ttu-id="40310-210">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-211">ReadItem</span></span>|
|[<span data-ttu-id="40310-212">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-213">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-214">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="40310-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="40310-215">dateTimeModified :Date</span></span>

<span data-ttu-id="40310-p110">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-218">Ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="40310-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-219">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-219">Type:</span></span>

*   <span data-ttu-id="40310-220">Date</span><span class="sxs-lookup"><span data-stu-id="40310-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-221">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-221">Requirements</span></span>

|<span data-ttu-id="40310-222">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-222">Requirement</span></span>| <span data-ttu-id="40310-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-224">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-225">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-225">1.0</span></span>|
|[<span data-ttu-id="40310-226">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-227">ReadItem</span></span>|
|[<span data-ttu-id="40310-228">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-229">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-230">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="40310-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="40310-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="40310-232">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="40310-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="40310-p111">La propriété `end` est exprimée en date et heure U.T.C. (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="40310-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="40310-235">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="40310-235">Read mode</span></span>

<span data-ttu-id="40310-236">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="40310-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="40310-237">Mode composition</span><span class="sxs-lookup"><span data-stu-id="40310-237">Compose mode</span></span>

<span data-ttu-id="40310-238">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="40310-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="40310-239">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="40310-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-240">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-240">Type:</span></span>

*   <span data-ttu-id="40310-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="40310-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-242">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-242">Requirements</span></span>

|<span data-ttu-id="40310-243">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-243">Requirement</span></span>| <span data-ttu-id="40310-244">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-245">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-245">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-246">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-246">1.0</span></span>|
|[<span data-ttu-id="40310-247">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-248">ReadItem</span></span>|
|[<span data-ttu-id="40310-249">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-250">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-251">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-251">Example</span></span>

<span data-ttu-id="40310-252">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="40310-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="40310-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="40310-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="40310-p112">Obtient l’adresse e-mail de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="40310-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété expéditeur représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="40310-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-258">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="40310-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-259">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-259">Type:</span></span>

*   [<span data-ttu-id="40310-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="40310-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="40310-261">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-261">Requirements</span></span>

|<span data-ttu-id="40310-262">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-262">Requirement</span></span>| <span data-ttu-id="40310-263">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-264">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-265">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-265">1.0</span></span>|
|[<span data-ttu-id="40310-266">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-267">ReadItem</span></span>|
|[<span data-ttu-id="40310-268">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-269">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="40310-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="40310-270">internetMessageId :String</span></span>

<span data-ttu-id="40310-p114">Obtient l’identificateur de message Internet d’un e-mail. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-273">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-273">Type:</span></span>

*   <span data-ttu-id="40310-274">String</span><span class="sxs-lookup"><span data-stu-id="40310-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-275">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-275">Requirements</span></span>

|<span data-ttu-id="40310-276">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-276">Requirement</span></span>| <span data-ttu-id="40310-277">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-278">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-279">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-279">1.0</span></span>|
|[<span data-ttu-id="40310-280">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-281">ReadItem</span></span>|
|[<span data-ttu-id="40310-282">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-283">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-284">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="40310-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="40310-285">itemClass :String</span></span>

<span data-ttu-id="40310-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="40310-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="40310-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="40310-290">Type</span><span class="sxs-lookup"><span data-stu-id="40310-290">Type</span></span> | <span data-ttu-id="40310-291">Description</span><span class="sxs-lookup"><span data-stu-id="40310-291">Description</span></span> | <span data-ttu-id="40310-292">Classe d’élément</span><span class="sxs-lookup"><span data-stu-id="40310-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="40310-293">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="40310-293">Appointment items</span></span> | <span data-ttu-id="40310-294">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="40310-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="40310-295">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="40310-295">Message items</span></span> | <span data-ttu-id="40310-296">Ces éléments incluent les e-mails dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="40310-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="40310-297">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="40310-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-298">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-298">Type:</span></span>

*   <span data-ttu-id="40310-299">String</span><span class="sxs-lookup"><span data-stu-id="40310-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-300">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-300">Requirements</span></span>

|<span data-ttu-id="40310-301">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-301">Requirement</span></span>| <span data-ttu-id="40310-302">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-303">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-304">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-304">1.0</span></span>|
|[<span data-ttu-id="40310-305">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-306">ReadItem</span></span>|
|[<span data-ttu-id="40310-307">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-308">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-309">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="40310-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="40310-310">(nullable) itemId :String</span></span>

<span data-ttu-id="40310-p117">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-313">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="40310-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="40310-314">La propriété `itemId` n’est pas identique à l’identificateur d’entrée Outlook ou l’ID utilisé par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="40310-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="40310-315">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="40310-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="40310-316">Pour plus d’informations, voir [Utiliser les API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="40310-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="40310-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-319">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-319">Type:</span></span>

*   <span data-ttu-id="40310-320">String</span><span class="sxs-lookup"><span data-stu-id="40310-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-321">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-321">Requirements</span></span>

|<span data-ttu-id="40310-322">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-322">Requirement</span></span>| <span data-ttu-id="40310-323">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-324">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-325">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-325">1.0</span></span>|
|[<span data-ttu-id="40310-326">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-327">ReadItem</span></span>|
|[<span data-ttu-id="40310-328">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-329">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-330">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-330">Example</span></span>

<span data-ttu-id="40310-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="40310-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="40310-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="40310-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="40310-334">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="40310-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="40310-335">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="40310-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-336">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-336">Type:</span></span>

*   [<span data-ttu-id="40310-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="40310-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="40310-338">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-338">Requirements</span></span>

|<span data-ttu-id="40310-339">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-339">Requirement</span></span>| <span data-ttu-id="40310-340">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-341">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-341">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-342">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-342">1.0</span></span>|
|[<span data-ttu-id="40310-343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-344">ReadItem</span></span>|
|[<span data-ttu-id="40310-345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-346">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-347">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="40310-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="40310-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="40310-349">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="40310-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="40310-350">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="40310-350">Read mode</span></span>

<span data-ttu-id="40310-351">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="40310-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="40310-352">Mode composition</span><span class="sxs-lookup"><span data-stu-id="40310-352">Compose mode</span></span>

<span data-ttu-id="40310-353">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="40310-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-354">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-354">Type:</span></span>

*   <span data-ttu-id="40310-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="40310-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-356">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-356">Requirements</span></span>

|<span data-ttu-id="40310-357">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-357">Requirement</span></span>| <span data-ttu-id="40310-358">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-359">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-359">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-360">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-360">1.0</span></span>|
|[<span data-ttu-id="40310-361">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-362">ReadItem</span></span>|
|[<span data-ttu-id="40310-363">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-364">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-365">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="40310-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="40310-366">normalizedSubject :String</span></span>

<span data-ttu-id="40310-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="40310-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject).</span><span class="sxs-lookup"><span data-stu-id="40310-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-371">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-371">Type:</span></span>

*   <span data-ttu-id="40310-372">String</span><span class="sxs-lookup"><span data-stu-id="40310-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-373">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-373">Requirements</span></span>

|<span data-ttu-id="40310-374">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-374">Requirement</span></span>| <span data-ttu-id="40310-375">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-376">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-376">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-377">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-377">1.0</span></span>|
|[<span data-ttu-id="40310-378">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-379">ReadItem</span></span>|
|[<span data-ttu-id="40310-380">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-381">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-382">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="40310-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="40310-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="40310-384">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="40310-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-385">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-385">Type:</span></span>

*   [<span data-ttu-id="40310-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="40310-386">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="40310-387">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-387">Requirements</span></span>

|<span data-ttu-id="40310-388">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-388">Requirement</span></span>| <span data-ttu-id="40310-389">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-390">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-390">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-391">1.3</span><span class="sxs-lookup"><span data-stu-id="40310-391">1.3</span></span>|
|[<span data-ttu-id="40310-392">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-393">ReadItem</span></span>|
|[<span data-ttu-id="40310-394">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-395">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="40310-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="40310-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="40310-397">Fournit l’accès aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="40310-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="40310-398">Le type d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="40310-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="40310-399">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="40310-399">Read mode</span></span>

<span data-ttu-id="40310-400">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="40310-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="40310-401">Mode composition</span><span class="sxs-lookup"><span data-stu-id="40310-401">Compose mode</span></span>

<span data-ttu-id="40310-402">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="40310-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-403">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-403">Type:</span></span>

*   <span data-ttu-id="40310-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="40310-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-405">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-405">Requirements</span></span>

|<span data-ttu-id="40310-406">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-406">Requirement</span></span>| <span data-ttu-id="40310-407">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-408">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-408">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-409">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-409">1.0</span></span>|
|[<span data-ttu-id="40310-410">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-411">ReadItem</span></span>|
|[<span data-ttu-id="40310-412">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-413">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-414">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="40310-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="40310-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="40310-p124">Obtient l’adresse e-mail de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-418">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-418">Type:</span></span>

*   [<span data-ttu-id="40310-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="40310-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="40310-420">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-420">Requirements</span></span>

|<span data-ttu-id="40310-421">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-421">Requirement</span></span>| <span data-ttu-id="40310-422">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-423">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-424">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-424">1.0</span></span>|
|[<span data-ttu-id="40310-425">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-426">ReadItem</span></span>|
|[<span data-ttu-id="40310-427">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-428">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-429">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="40310-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="40310-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="40310-431">Fournit l’accès aux participants obligatoires d’un événement.</span><span class="sxs-lookup"><span data-stu-id="40310-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="40310-432">Le type d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="40310-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="40310-433">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="40310-433">Read mode</span></span>

<span data-ttu-id="40310-434">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant obligatoires de la réunion.</span><span class="sxs-lookup"><span data-stu-id="40310-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="40310-435">Mode composition</span><span class="sxs-lookup"><span data-stu-id="40310-435">Compose mode</span></span>

<span data-ttu-id="40310-436">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="40310-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-437">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-437">Type:</span></span>

*   <span data-ttu-id="40310-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="40310-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-439">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-439">Requirements</span></span>

|<span data-ttu-id="40310-440">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-440">Requirement</span></span>| <span data-ttu-id="40310-441">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-442">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-443">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-443">1.0</span></span>|
|[<span data-ttu-id="40310-444">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-445">ReadItem</span></span>|
|[<span data-ttu-id="40310-446">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-447">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-448">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="40310-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="40310-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="40310-p126">Obtient l’adresse e-mail de l’expéditeur d’un e-mail. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="40310-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="40310-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété expéditeur représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="40310-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-454">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="40310-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-455">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-455">Type:</span></span>

*   [<span data-ttu-id="40310-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="40310-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="40310-457">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-457">Requirements</span></span>

|<span data-ttu-id="40310-458">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-458">Requirement</span></span>| <span data-ttu-id="40310-459">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-460">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-460">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-461">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-461">1.0</span></span>|
|[<span data-ttu-id="40310-462">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-463">ReadItem</span></span>|
|[<span data-ttu-id="40310-464">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-465">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-466">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="40310-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="40310-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="40310-468">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="40310-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="40310-p128">La propriété `start` est exprimée en date et heure U.T.C. (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="40310-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="40310-471">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="40310-471">Read mode</span></span>

<span data-ttu-id="40310-472">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="40310-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="40310-473">Mode composition</span><span class="sxs-lookup"><span data-stu-id="40310-473">Compose mode</span></span>

<span data-ttu-id="40310-474">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="40310-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="40310-475">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format U.T.C. pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="40310-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-476">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-476">Type:</span></span>

*   <span data-ttu-id="40310-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="40310-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-478">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-478">Requirements</span></span>

|<span data-ttu-id="40310-479">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-479">Requirement</span></span>| <span data-ttu-id="40310-480">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-481">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-481">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-482">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-482">1.0</span></span>|
|[<span data-ttu-id="40310-483">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-484">ReadItem</span></span>|
|[<span data-ttu-id="40310-485">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-486">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-487">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-487">Example</span></span>

<span data-ttu-id="40310-488">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="40310-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="40310-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="40310-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="40310-490">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="40310-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="40310-491">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="40310-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="40310-492">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="40310-492">Read mode</span></span>

<span data-ttu-id="40310-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="40310-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="40310-495">Mode composition</span><span class="sxs-lookup"><span data-stu-id="40310-495">Compose mode</span></span>

<span data-ttu-id="40310-496">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="40310-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="40310-497">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-497">Type:</span></span>

*   <span data-ttu-id="40310-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="40310-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-499">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-499">Requirements</span></span>

|<span data-ttu-id="40310-500">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-500">Requirement</span></span>| <span data-ttu-id="40310-501">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-502">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-503">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-503">1.0</span></span>|
|[<span data-ttu-id="40310-504">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-505">ReadItem</span></span>|
|[<span data-ttu-id="40310-506">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-507">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="40310-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="40310-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="40310-509">Permet d’accéder aux destinataires de la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="40310-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="40310-510">Le type d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="40310-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="40310-511">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="40310-511">Read mode</span></span>

<span data-ttu-id="40310-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="40310-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="40310-514">Mode composition</span><span class="sxs-lookup"><span data-stu-id="40310-514">Compose mode</span></span>

<span data-ttu-id="40310-515">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="40310-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="40310-516">Type :</span><span class="sxs-lookup"><span data-stu-id="40310-516">Type:</span></span>

*   <span data-ttu-id="40310-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="40310-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-518">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-518">Requirements</span></span>

|<span data-ttu-id="40310-519">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-519">Requirement</span></span>| <span data-ttu-id="40310-520">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-521">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-522">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-522">1.0</span></span>|
|[<span data-ttu-id="40310-523">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-524">ReadItem</span></span>|
|[<span data-ttu-id="40310-525">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-526">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-527">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="40310-528">Méthodes</span><span class="sxs-lookup"><span data-stu-id="40310-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="40310-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="40310-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="40310-530">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="40310-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="40310-531">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="40310-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="40310-532">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="40310-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-533">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-533">Parameters:</span></span>

|<span data-ttu-id="40310-534">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-534">Name</span></span>| <span data-ttu-id="40310-535">Type</span><span class="sxs-lookup"><span data-stu-id="40310-535">Type</span></span>| <span data-ttu-id="40310-536">Attributs</span><span class="sxs-lookup"><span data-stu-id="40310-536">Attributes</span></span>| <span data-ttu-id="40310-537">Description</span><span class="sxs-lookup"><span data-stu-id="40310-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="40310-538">String</span><span class="sxs-lookup"><span data-stu-id="40310-538">String</span></span>||<span data-ttu-id="40310-p132">L’URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="40310-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="40310-541">String</span><span class="sxs-lookup"><span data-stu-id="40310-541">String</span></span>||<span data-ttu-id="40310-p133">Nom de la pièce jointe affiché lors de son chargement. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="40310-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="40310-544">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-544">Object</span></span>| <span data-ttu-id="40310-545">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-545">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-546">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="40310-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="40310-547">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-547">Object</span></span>| <span data-ttu-id="40310-548">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-548">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-549">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="40310-550">fonction</span><span class="sxs-lookup"><span data-stu-id="40310-550">function</span></span>| <span data-ttu-id="40310-551">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-551">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-552">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="40310-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="40310-553">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="40310-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="40310-554">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="40310-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="40310-555">Erreurs</span><span class="sxs-lookup"><span data-stu-id="40310-555">Errors</span></span>

| <span data-ttu-id="40310-556">Code d’erreur</span><span class="sxs-lookup"><span data-stu-id="40310-556">Error code</span></span> | <span data-ttu-id="40310-557">Description</span><span class="sxs-lookup"><span data-stu-id="40310-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="40310-558">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="40310-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="40310-559">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="40310-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="40310-560">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="40310-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="40310-561">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-561">Requirements</span></span>

|<span data-ttu-id="40310-562">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-562">Requirement</span></span>| <span data-ttu-id="40310-563">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-564">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-565">1.1</span><span class="sxs-lookup"><span data-stu-id="40310-565">1.1</span></span>|
|[<span data-ttu-id="40310-566">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="40310-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="40310-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-569">Composition</span><span class="sxs-lookup"><span data-stu-id="40310-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-570">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="40310-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="40310-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="40310-572">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="40310-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="40310-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="40310-576">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="40310-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="40310-577">Si votre complément Office est exécuté dans la Outlook Web App , la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="40310-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-578">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-578">Parameters:</span></span>

|<span data-ttu-id="40310-579">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-579">Name</span></span>| <span data-ttu-id="40310-580">Type</span><span class="sxs-lookup"><span data-stu-id="40310-580">Type</span></span>| <span data-ttu-id="40310-581">Attributs</span><span class="sxs-lookup"><span data-stu-id="40310-581">Attributes</span></span>| <span data-ttu-id="40310-582">Description</span><span class="sxs-lookup"><span data-stu-id="40310-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="40310-583">String</span><span class="sxs-lookup"><span data-stu-id="40310-583">String</span></span>||<span data-ttu-id="40310-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="40310-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="40310-586">String</span><span class="sxs-lookup"><span data-stu-id="40310-586">String</span></span>||<span data-ttu-id="40310-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="40310-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="40310-589">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-589">Object</span></span>| <span data-ttu-id="40310-590">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-590">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-591">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="40310-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="40310-592">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-592">Object</span></span>| <span data-ttu-id="40310-593">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-593">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-594">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="40310-595">fonction</span><span class="sxs-lookup"><span data-stu-id="40310-595">function</span></span>| <span data-ttu-id="40310-596">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-596">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-597">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="40310-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="40310-598">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="40310-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="40310-599">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="40310-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="40310-600">Erreurs</span><span class="sxs-lookup"><span data-stu-id="40310-600">Errors</span></span>

| <span data-ttu-id="40310-601">Code d’erreur</span><span class="sxs-lookup"><span data-stu-id="40310-601">Error code</span></span> | <span data-ttu-id="40310-602">Description</span><span class="sxs-lookup"><span data-stu-id="40310-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="40310-603">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="40310-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="40310-604">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-604">Requirements</span></span>

|<span data-ttu-id="40310-605">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-605">Requirement</span></span>| <span data-ttu-id="40310-606">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-607">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-607">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-608">1.1</span><span class="sxs-lookup"><span data-stu-id="40310-608">1.1</span></span>|
|[<span data-ttu-id="40310-609">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="40310-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="40310-611">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-612">Composition</span><span class="sxs-lookup"><span data-stu-id="40310-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-613">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-613">Example</span></span>

<span data-ttu-id="40310-614">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="40310-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
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

####  <a name="close"></a><span data-ttu-id="40310-615">close()</span><span class="sxs-lookup"><span data-stu-id="40310-615">close()</span></span>

<span data-ttu-id="40310-616">Ferme l’élément actif qui est en train d’être composé.</span><span class="sxs-lookup"><span data-stu-id="40310-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="40310-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action fermer.</span><span class="sxs-lookup"><span data-stu-id="40310-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-619">Sur Outlook Web Access, si l’élément est un rendez-vous qui a déjà été sauvegardé en utilisant la méthode `saveAsync` , l’utilisateur sera inviter à sauvegarder, abandonner ou annuler même si l’élément n'a subi aucun changement depuis sa dernière sauvegarde.</span><span class="sxs-lookup"><span data-stu-id="40310-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="40310-620">Dans la version logicielle d’Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="40310-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-621">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-621">Requirements</span></span>

|<span data-ttu-id="40310-622">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-622">Requirement</span></span>| <span data-ttu-id="40310-623">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-624">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-624">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-625">1.3</span><span class="sxs-lookup"><span data-stu-id="40310-625">1.3</span></span>|
|[<span data-ttu-id="40310-626">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-627">Restreint</span><span class="sxs-lookup"><span data-stu-id="40310-627">Restricted</span></span>|
|[<span data-ttu-id="40310-628">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-629">Composition</span><span class="sxs-lookup"><span data-stu-id="40310-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="40310-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="40310-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="40310-631">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="40310-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-632">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="40310-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="40310-633">Sur Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="40310-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="40310-634">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="40310-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="40310-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, alors aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="40310-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-638">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-638">Parameters:</span></span>

|<span data-ttu-id="40310-639">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-639">Name</span></span>| <span data-ttu-id="40310-640">Type</span><span class="sxs-lookup"><span data-stu-id="40310-640">Type</span></span>| <span data-ttu-id="40310-641">Description</span><span class="sxs-lookup"><span data-stu-id="40310-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="40310-642">String | Object</span><span class="sxs-lookup"><span data-stu-id="40310-642">String &#124; Object</span></span>| |<span data-ttu-id="40310-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="40310-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="40310-645">**OR**</span><span class="sxs-lookup"><span data-stu-id="40310-645">**OR**</span></span><br/><span data-ttu-id="40310-p140">Objet qui contient les données du texte du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="40310-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="40310-648">String</span><span class="sxs-lookup"><span data-stu-id="40310-648">String</span></span> | <span data-ttu-id="40310-649">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-649">&lt;optional&gt;</span></span> | <span data-ttu-id="40310-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="40310-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="40310-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="40310-653">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-653">&lt;optional&gt;</span></span> | <span data-ttu-id="40310-654">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="40310-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="40310-655">String</span><span class="sxs-lookup"><span data-stu-id="40310-655">String</span></span> | | <span data-ttu-id="40310-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="40310-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="40310-658">String</span><span class="sxs-lookup"><span data-stu-id="40310-658">String</span></span> | | <span data-ttu-id="40310-659">Chaîne qui contient le nom de la pièce jointe, d’une longueur maximale de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="40310-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="40310-660">String</span><span class="sxs-lookup"><span data-stu-id="40310-660">String</span></span> | | <span data-ttu-id="40310-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="40310-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="40310-663">String</span><span class="sxs-lookup"><span data-stu-id="40310-663">String</span></span> | | <span data-ttu-id="40310-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’identificateur de l’élément EWS de la pièce jointe. Cette chaîne doit être d’une longueur maximale de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="40310-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="40310-667">fonction</span><span class="sxs-lookup"><span data-stu-id="40310-667">function</span></span> | <span data-ttu-id="40310-668">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-668">&lt;optional&gt;</span></span> | <span data-ttu-id="40310-669">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="40310-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="40310-670">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-670">Requirements</span></span>

|<span data-ttu-id="40310-671">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-671">Requirement</span></span>| <span data-ttu-id="40310-672">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-673">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-673">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-674">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-674">1.0</span></span>|
|[<span data-ttu-id="40310-675">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-676">ReadItem</span></span>|
|[<span data-ttu-id="40310-677">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-678">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="40310-679">Exemples</span><span class="sxs-lookup"><span data-stu-id="40310-679">Examples</span></span>

<span data-ttu-id="40310-680">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="40310-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="40310-681">Réponse sans texte.</span><span class="sxs-lookup"><span data-stu-id="40310-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="40310-682">Réponse avec un texte.</span><span class="sxs-lookup"><span data-stu-id="40310-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="40310-683">Réponse avec un texte et un fichier comme pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="40310-683">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="40310-684">Réponse avec un texte et un élément comme pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="40310-684">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="40310-685">Réponse avec un texte, un fichier comme pièce jointe, un élément comme pièce jointe et un rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="40310-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="40310-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="40310-687">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="40310-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-688">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="40310-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="40310-689">Sur Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="40310-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="40310-690">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="40310-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="40310-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, alors aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="40310-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-694">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-694">Parameters:</span></span>

|<span data-ttu-id="40310-695">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-695">Name</span></span>| <span data-ttu-id="40310-696">Type</span><span class="sxs-lookup"><span data-stu-id="40310-696">Type</span></span>| <span data-ttu-id="40310-697">Description</span><span class="sxs-lookup"><span data-stu-id="40310-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="40310-698">String | Object</span><span class="sxs-lookup"><span data-stu-id="40310-698">String &#124; Object</span></span>| | <span data-ttu-id="40310-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="40310-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="40310-701">**OR**</span><span class="sxs-lookup"><span data-stu-id="40310-701">**OR**</span></span><br/><span data-ttu-id="40310-p147">Objet qui contient les données du texte du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="40310-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="40310-704">String</span><span class="sxs-lookup"><span data-stu-id="40310-704">String</span></span> | <span data-ttu-id="40310-705">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-705">&lt;optional&gt;</span></span> | <span data-ttu-id="40310-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="40310-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="40310-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="40310-709">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-709">&lt;optional&gt;</span></span> | <span data-ttu-id="40310-710">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="40310-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="40310-711">String</span><span class="sxs-lookup"><span data-stu-id="40310-711">String</span></span> | | <span data-ttu-id="40310-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="40310-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="40310-714">String</span><span class="sxs-lookup"><span data-stu-id="40310-714">String</span></span> | | <span data-ttu-id="40310-715">Chaîne qui contient le nom de la pièce jointe, d’une longueur maximale de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="40310-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="40310-716">String</span><span class="sxs-lookup"><span data-stu-id="40310-716">String</span></span> | | <span data-ttu-id="40310-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="40310-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="40310-719">String</span><span class="sxs-lookup"><span data-stu-id="40310-719">String</span></span> | | <span data-ttu-id="40310-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’identificateur de l’élément EWS de la pièce jointe. Cette chaîne doit être d’une longueur maximale de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="40310-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="40310-723">fonction</span><span class="sxs-lookup"><span data-stu-id="40310-723">function</span></span> | <span data-ttu-id="40310-724">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-724">&lt;optional&gt;</span></span> | <span data-ttu-id="40310-725">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="40310-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="40310-726">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-726">Requirements</span></span>

|<span data-ttu-id="40310-727">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-727">Requirement</span></span>| <span data-ttu-id="40310-728">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-729">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-729">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-730">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-730">1.0</span></span>|
|[<span data-ttu-id="40310-731">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-732">ReadItem</span></span>|
|[<span data-ttu-id="40310-733">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-734">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="40310-735">Exemples</span><span class="sxs-lookup"><span data-stu-id="40310-735">Examples</span></span>

<span data-ttu-id="40310-736">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="40310-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="40310-737">Réponse sans texte.</span><span class="sxs-lookup"><span data-stu-id="40310-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="40310-738">Réponse avec un texte.</span><span class="sxs-lookup"><span data-stu-id="40310-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="40310-739">Réponse avec un texte et un fichier comme pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="40310-739">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="40310-740">Réponse avec un texte et un élément comme pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="40310-740">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="40310-741">Réponse avec un texte, un fichier comme pièce jointe, un élément comme pièce jointe et un rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="40310-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="40310-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="40310-743">Obtient les entités figurant dans le texte de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="40310-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-744">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="40310-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-745">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-745">Requirements</span></span>

|<span data-ttu-id="40310-746">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-746">Requirement</span></span>| <span data-ttu-id="40310-747">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-748">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-748">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-749">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-749">1.0</span></span>|
|[<span data-ttu-id="40310-750">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-751">ReadItem</span></span>|
|[<span data-ttu-id="40310-752">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-753">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="40310-754">Retourne :</span><span class="sxs-lookup"><span data-stu-id="40310-754">Returns:</span></span>

<span data-ttu-id="40310-755">Type : [Entités](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="40310-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="40310-756">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-756">Example</span></span>

<span data-ttu-id="40310-757">L’exemple suivant accède aux entités de contacts dans l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="40310-757">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="40310-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="40310-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="40310-759">Obtient un tableau de toutes les entités du type spécifié trouvées dans le texte de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="40310-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-760">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="40310-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-761">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-761">Parameters:</span></span>

|<span data-ttu-id="40310-762">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-762">Name</span></span>| <span data-ttu-id="40310-763">Type</span><span class="sxs-lookup"><span data-stu-id="40310-763">Type</span></span>| <span data-ttu-id="40310-764">Description</span><span class="sxs-lookup"><span data-stu-id="40310-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="40310-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="40310-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="40310-766">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="40310-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="40310-767">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-767">Requirements</span></span>

|<span data-ttu-id="40310-768">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-768">Requirement</span></span>| <span data-ttu-id="40310-769">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-770">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-770">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-771">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-771">1.0</span></span>|
|[<span data-ttu-id="40310-772">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-773">Restreint</span><span class="sxs-lookup"><span data-stu-id="40310-773">Restricted</span></span>|
|[<span data-ttu-id="40310-774">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-775">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="40310-776">Retourne :</span><span class="sxs-lookup"><span data-stu-id="40310-776">Returns:</span></span>

<span data-ttu-id="40310-777">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="40310-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="40310-778">Si aucune entité du type spécifié n’est présente dans le texte de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="40310-778">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="40310-779">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="40310-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="40310-780">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="40310-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="40310-781">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="40310-781">Value of `entityType`</span></span> | <span data-ttu-id="40310-782">Type des objets dans le tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="40310-782">Type of objects in returned array</span></span> | <span data-ttu-id="40310-783">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="40310-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="40310-784">String</span><span class="sxs-lookup"><span data-stu-id="40310-784">String</span></span> | <span data-ttu-id="40310-785">**Restreint**</span><span class="sxs-lookup"><span data-stu-id="40310-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="40310-786">Contact</span><span class="sxs-lookup"><span data-stu-id="40310-786">Contact</span></span> | <span data-ttu-id="40310-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="40310-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="40310-788">String</span><span class="sxs-lookup"><span data-stu-id="40310-788">String</span></span> | <span data-ttu-id="40310-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="40310-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="40310-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="40310-790">MeetingSuggestion</span></span> | <span data-ttu-id="40310-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="40310-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="40310-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="40310-792">PhoneNumber</span></span> | <span data-ttu-id="40310-793">**Restreint**</span><span class="sxs-lookup"><span data-stu-id="40310-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="40310-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="40310-794">TaskSuggestion</span></span> | <span data-ttu-id="40310-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="40310-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="40310-796">String</span><span class="sxs-lookup"><span data-stu-id="40310-796">String</span></span> | <span data-ttu-id="40310-797">**Restreint**</span><span class="sxs-lookup"><span data-stu-id="40310-797">**Restricted**</span></span> |

<span data-ttu-id="40310-798">Type : Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="40310-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="40310-799">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-799">Example</span></span>

<span data-ttu-id="40310-800">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le texte de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="40310-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="40310-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="40310-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="40310-802">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="40310-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-803">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="40310-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="40310-804">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="40310-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-805">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-805">Parameters:</span></span>

|<span data-ttu-id="40310-806">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-806">Name</span></span>| <span data-ttu-id="40310-807">Type</span><span class="sxs-lookup"><span data-stu-id="40310-807">Type</span></span>| <span data-ttu-id="40310-808">Description</span><span class="sxs-lookup"><span data-stu-id="40310-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="40310-809">String</span><span class="sxs-lookup"><span data-stu-id="40310-809">String</span></span>|<span data-ttu-id="40310-810">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="40310-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="40310-811">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-811">Requirements</span></span>

|<span data-ttu-id="40310-812">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-812">Requirement</span></span>| <span data-ttu-id="40310-813">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-814">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-814">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-815">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-815">1.0</span></span>|
|[<span data-ttu-id="40310-816">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-817">ReadItem</span></span>|
|[<span data-ttu-id="40310-818">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-819">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="40310-820">Retourne :</span><span class="sxs-lookup"><span data-stu-id="40310-820">Returns:</span></span>

<span data-ttu-id="40310-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="40310-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="40310-823">Type : Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="40310-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="40310-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="40310-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="40310-825">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="40310-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-826">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="40310-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="40310-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="40310-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="40310-830">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="40310-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="40310-831">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="40310-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="40310-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le texte. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du texte de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du texte d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du texte de l’élément.</span><span class="sxs-lookup"><span data-stu-id="40310-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="40310-835">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-835">Requirements</span></span>

|<span data-ttu-id="40310-836">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-836">Requirement</span></span>| <span data-ttu-id="40310-837">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-838">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-838">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-839">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-839">1.0</span></span>|
|[<span data-ttu-id="40310-840">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-841">ReadItem</span></span>|
|[<span data-ttu-id="40310-842">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-843">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="40310-844">Retourne :</span><span class="sxs-lookup"><span data-stu-id="40310-844">Returns:</span></span>

<span data-ttu-id="40310-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="40310-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="40310-847">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="40310-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="40310-848">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="40310-849">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-849">Example</span></span>

<span data-ttu-id="40310-850">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="40310-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="40310-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="40310-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="40310-852">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="40310-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-853">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="40310-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="40310-854">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="40310-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="40310-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de texte d’un élément, l’expression régulière doit filtrer davantage le texte. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du texte de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du texte d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="40310-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-857">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-857">Parameters:</span></span>

|<span data-ttu-id="40310-858">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-858">Name</span></span>| <span data-ttu-id="40310-859">Type</span><span class="sxs-lookup"><span data-stu-id="40310-859">Type</span></span>| <span data-ttu-id="40310-860">Description</span><span class="sxs-lookup"><span data-stu-id="40310-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="40310-861">String</span><span class="sxs-lookup"><span data-stu-id="40310-861">String</span></span>|<span data-ttu-id="40310-862">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="40310-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="40310-863">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-863">Requirements</span></span>

|<span data-ttu-id="40310-864">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-864">Requirement</span></span>| <span data-ttu-id="40310-865">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-866">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-866">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-867">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-867">1.0</span></span>|
|[<span data-ttu-id="40310-868">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-869">ReadItem</span></span>|
|[<span data-ttu-id="40310-870">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-871">Lecture</span><span class="sxs-lookup"><span data-stu-id="40310-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="40310-872">Retourne :</span><span class="sxs-lookup"><span data-stu-id="40310-872">Returns:</span></span>

<span data-ttu-id="40310-873">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="40310-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="40310-874">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="40310-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="40310-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="40310-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="40310-876">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="40310-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="40310-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="40310-878">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du texte d’un message.</span><span class="sxs-lookup"><span data-stu-id="40310-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="40310-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="40310-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-881">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-881">Parameters:</span></span>

|<span data-ttu-id="40310-882">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-882">Name</span></span>| <span data-ttu-id="40310-883">Type</span><span class="sxs-lookup"><span data-stu-id="40310-883">Type</span></span>| <span data-ttu-id="40310-884">Attributs</span><span class="sxs-lookup"><span data-stu-id="40310-884">Attributes</span></span>| <span data-ttu-id="40310-885">Description</span><span class="sxs-lookup"><span data-stu-id="40310-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="40310-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="40310-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="40310-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="40310-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="40310-890">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-890">Object</span></span>| <span data-ttu-id="40310-891">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-891">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-892">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="40310-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="40310-893">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-893">Object</span></span>| <span data-ttu-id="40310-894">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-894">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-895">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="40310-896">fonction</span><span class="sxs-lookup"><span data-stu-id="40310-896">function</span></span>||<span data-ttu-id="40310-897">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="40310-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="40310-898">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="40310-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="40310-899">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="40310-899">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="40310-900">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-900">Requirements</span></span>

|<span data-ttu-id="40310-901">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-901">Requirement</span></span>| <span data-ttu-id="40310-902">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-903">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-903">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-904">1.2</span><span class="sxs-lookup"><span data-stu-id="40310-904">1.2</span></span>|
|[<span data-ttu-id="40310-905">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="40310-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="40310-907">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-908">Composition</span><span class="sxs-lookup"><span data-stu-id="40310-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="40310-909">Retourne :</span><span class="sxs-lookup"><span data-stu-id="40310-909">Returns:</span></span>

<span data-ttu-id="40310-910">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="40310-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="40310-911">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="40310-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="40310-912">String</span><span class="sxs-lookup"><span data-stu-id="40310-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="40310-913">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-913">Example</span></span>

```
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="40310-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="40310-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="40310-915">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="40310-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="40310-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="40310-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-919">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-919">Parameters:</span></span>

|<span data-ttu-id="40310-920">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-920">Name</span></span>| <span data-ttu-id="40310-921">Type</span><span class="sxs-lookup"><span data-stu-id="40310-921">Type</span></span>| <span data-ttu-id="40310-922">Attributs</span><span class="sxs-lookup"><span data-stu-id="40310-922">Attributes</span></span>| <span data-ttu-id="40310-923">Description</span><span class="sxs-lookup"><span data-stu-id="40310-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="40310-924">fonction</span><span class="sxs-lookup"><span data-stu-id="40310-924">function</span></span>||<span data-ttu-id="40310-925">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="40310-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="40310-926">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="40310-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="40310-927">Cet objet peut être utilisé pour obtenir, définir et supprimer les propriétés personnalisées de l’élément et sauvegarder les modifications à la propriété personnalisée est redéfinie sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="40310-927">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="40310-928">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-928">Object</span></span>| <span data-ttu-id="40310-929">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-929">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-930">Les développeurs peuvent fournir n'importe quel objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-930">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="40310-931">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="40310-932">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-932">Requirements</span></span>

|<span data-ttu-id="40310-933">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-933">Requirement</span></span>| <span data-ttu-id="40310-934">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-935">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-936">1.0</span><span class="sxs-lookup"><span data-stu-id="40310-936">1.0</span></span>|
|[<span data-ttu-id="40310-937">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40310-938">ReadItem</span></span>|
|[<span data-ttu-id="40310-939">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-940">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="40310-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-941">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-941">Example</span></span>

<span data-ttu-id="40310-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="40310-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="40310-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="40310-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="40310-946">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="40310-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="40310-p165">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="40310-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-951">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-951">Parameters:</span></span>

|<span data-ttu-id="40310-952">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-952">Name</span></span>| <span data-ttu-id="40310-953">Type</span><span class="sxs-lookup"><span data-stu-id="40310-953">Type</span></span>| <span data-ttu-id="40310-954">Attributs</span><span class="sxs-lookup"><span data-stu-id="40310-954">Attributes</span></span>| <span data-ttu-id="40310-955">Description</span><span class="sxs-lookup"><span data-stu-id="40310-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="40310-956">String</span><span class="sxs-lookup"><span data-stu-id="40310-956">String</span></span>||<span data-ttu-id="40310-p166">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="40310-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="40310-959">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-959">Object</span></span>| <span data-ttu-id="40310-960">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-960">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-961">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="40310-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="40310-962">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-962">Object</span></span>| <span data-ttu-id="40310-963">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-963">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-964">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="40310-965">fonction</span><span class="sxs-lookup"><span data-stu-id="40310-965">function</span></span>| <span data-ttu-id="40310-966">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-966">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-967">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="40310-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="40310-968">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="40310-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="40310-969">Erreurs</span><span class="sxs-lookup"><span data-stu-id="40310-969">Errors</span></span>

| <span data-ttu-id="40310-970">Code d’erreur</span><span class="sxs-lookup"><span data-stu-id="40310-970">Error code</span></span> | <span data-ttu-id="40310-971">Description</span><span class="sxs-lookup"><span data-stu-id="40310-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="40310-972">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="40310-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="40310-973">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-973">Requirements</span></span>

|<span data-ttu-id="40310-974">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-974">Requirement</span></span>| <span data-ttu-id="40310-975">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-976">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-976">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-977">1.1</span><span class="sxs-lookup"><span data-stu-id="40310-977">1.1</span></span>|
|[<span data-ttu-id="40310-978">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="40310-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="40310-980">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-981">Composition</span><span class="sxs-lookup"><span data-stu-id="40310-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-982">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-982">Example</span></span>

<span data-ttu-id="40310-983">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="40310-983">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="40310-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="40310-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="40310-985">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="40310-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="40310-p167">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’identificateur de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="40310-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-989">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` pour utiliser avec EWS ou l’API REST, gardez à l’esprit que quand Outlook est en mode mis en cache, il peut prendre un certain temps avant que l’élément est réellement synchronisé avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="40310-989">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="40310-990">Jusqu'à ce que l’élément soit synchronisé, utiliser la propriété `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="40310-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="40310-p169">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="40310-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="40310-994">Les clients suivants ont un comportement différent pour `saveAsync` pour les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="40310-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="40310-995">Mac Outlook ne gère pas `saveAsync` pour une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="40310-995">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="40310-996">Faire appel à `saveAsync` pour une réunion dans Outlook Mac renverra une erreur.</span><span class="sxs-lookup"><span data-stu-id="40310-996">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="40310-997">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée pour un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="40310-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-998">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-998">Parameters:</span></span>

|<span data-ttu-id="40310-999">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-999">Name</span></span>| <span data-ttu-id="40310-1000">Type</span><span class="sxs-lookup"><span data-stu-id="40310-1000">Type</span></span>| <span data-ttu-id="40310-1001">Attributs</span><span class="sxs-lookup"><span data-stu-id="40310-1001">Attributes</span></span>| <span data-ttu-id="40310-1002">Description</span><span class="sxs-lookup"><span data-stu-id="40310-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="40310-1003">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-1003">Object</span></span>| <span data-ttu-id="40310-1004">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-1005">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="40310-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="40310-1006">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-1006">Object</span></span>| <span data-ttu-id="40310-1007">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-1008">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="40310-1009">fonction</span><span class="sxs-lookup"><span data-stu-id="40310-1009">function</span></span>||<span data-ttu-id="40310-1010">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="40310-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="40310-1011">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="40310-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="40310-1012">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-1012">Requirements</span></span>

|<span data-ttu-id="40310-1013">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-1013">Requirement</span></span>| <span data-ttu-id="40310-1014">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-1015">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-1015">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="40310-1016">1.3</span></span>|
|[<span data-ttu-id="40310-1017">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="40310-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="40310-1019">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-1020">Composition</span><span class="sxs-lookup"><span data-stu-id="40310-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="40310-1021">Exemples</span><span class="sxs-lookup"><span data-stu-id="40310-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="40310-p171">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="40310-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="40310-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="40310-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="40310-1025">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="40310-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="40310-p172">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le texte ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du texte ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="40310-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="40310-1029">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="40310-1029">Parameters:</span></span>

|<span data-ttu-id="40310-1030">Nom</span><span class="sxs-lookup"><span data-stu-id="40310-1030">Name</span></span>| <span data-ttu-id="40310-1031">Type</span><span class="sxs-lookup"><span data-stu-id="40310-1031">Type</span></span>| <span data-ttu-id="40310-1032">Attributs</span><span class="sxs-lookup"><span data-stu-id="40310-1032">Attributes</span></span>| <span data-ttu-id="40310-1033">Description</span><span class="sxs-lookup"><span data-stu-id="40310-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="40310-1034">String</span><span class="sxs-lookup"><span data-stu-id="40310-1034">String</span></span>||<span data-ttu-id="40310-p173">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="40310-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="40310-1038">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-1038">Object</span></span>| <span data-ttu-id="40310-1039">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-1040">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="40310-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="40310-1041">Objet</span><span class="sxs-lookup"><span data-stu-id="40310-1041">Object</span></span>| <span data-ttu-id="40310-1042">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-1043">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="40310-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="40310-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="40310-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="40310-1045">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="40310-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="40310-p174">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="40310-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="40310-p175">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="40310-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="40310-1050">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé. Si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="40310-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="40310-1051">fonction</span><span class="sxs-lookup"><span data-stu-id="40310-1051">function</span></span>||<span data-ttu-id="40310-1052">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="40310-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="40310-1053">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-1053">Requirements</span></span>

|<span data-ttu-id="40310-1054">Condition requise</span><span class="sxs-lookup"><span data-stu-id="40310-1054">Requirement</span></span>| <span data-ttu-id="40310-1055">Valeur</span><span class="sxs-lookup"><span data-stu-id="40310-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="40310-1056">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="40310-1056">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40310-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="40310-1057">1.2</span></span>|
|[<span data-ttu-id="40310-1058">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="40310-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40310-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="40310-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="40310-1060">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="40310-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40310-1061">Composition</span><span class="sxs-lookup"><span data-stu-id="40310-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="40310-1062">Exemple</span><span class="sxs-lookup"><span data-stu-id="40310-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```