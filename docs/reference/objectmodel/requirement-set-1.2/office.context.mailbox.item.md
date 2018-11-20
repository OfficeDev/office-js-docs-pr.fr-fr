
# <a name="item"></a><span data-ttu-id="190f4-101">élément</span><span class="sxs-lookup"><span data-stu-id="190f4-101">item</span></span>

### <span data-ttu-id="190f4-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="190f4-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="190f4-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="190f4-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-106">Requirements</span></span>

|<span data-ttu-id="190f4-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-107">Requirement</span></span>| <span data-ttu-id="190f4-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-110">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-110">1.0</span></span>|
|[<span data-ttu-id="190f4-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="190f4-112">Restricted</span></span>|
|[<span data-ttu-id="190f4-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="190f4-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-115">Example</span></span>

<span data-ttu-id="190f4-116">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="190f4-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="190f4-117">Membres</span><span class="sxs-lookup"><span data-stu-id="190f4-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="190f4-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="190f4-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="190f4-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-121">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="190f4-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="190f4-122">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="190f4-122">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-123">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-123">Type:</span></span>

*   <span data-ttu-id="190f4-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="190f4-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-125">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-125">Requirements</span></span>

|<span data-ttu-id="190f4-126">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-126">Requirement</span></span>| <span data-ttu-id="190f4-127">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-128">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-129">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-129">1.0</span></span>|
|[<span data-ttu-id="190f4-130">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-131">ReadItem</span></span>|
|[<span data-ttu-id="190f4-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-133">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-134">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-134">Example</span></span>

<span data-ttu-id="190f4-135">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="190f4-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="190f4-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="190f4-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="190f4-137">Obtient un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="190f4-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="190f4-138">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-139">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-139">Type:</span></span>

*   [<span data-ttu-id="190f4-140">Destinataires</span><span class="sxs-lookup"><span data-stu-id="190f4-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="190f4-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-141">Requirements</span></span>

|<span data-ttu-id="190f4-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-142">Requirement</span></span>| <span data-ttu-id="190f4-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-145">1.1</span><span class="sxs-lookup"><span data-stu-id="190f4-145">1.1</span></span>|
|[<span data-ttu-id="190f4-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-147">ReadItem</span></span>|
|[<span data-ttu-id="190f4-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-149">Composition</span><span class="sxs-lookup"><span data-stu-id="190f4-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="190f4-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="190f4-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="190f4-152">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="190f4-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-153">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-153">Type:</span></span>

*   [<span data-ttu-id="190f4-154">Corps</span><span class="sxs-lookup"><span data-stu-id="190f4-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="190f4-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-155">Requirements</span></span>

|<span data-ttu-id="190f4-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-156">Requirement</span></span>| <span data-ttu-id="190f4-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-159">1.1</span><span class="sxs-lookup"><span data-stu-id="190f4-159">1.1</span></span>|
|[<span data-ttu-id="190f4-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-161">ReadItem</span></span>|
|[<span data-ttu-id="190f4-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="190f4-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="190f4-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="190f4-165">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="190f4-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="190f4-166">Le type d’objet et le niveau d’accès varie selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="190f4-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="190f4-167">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-167">Read mode</span></span>

<span data-ttu-id="190f4-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="190f4-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="190f4-170">Mode composition</span><span class="sxs-lookup"><span data-stu-id="190f4-170">Compose mode</span></span>

<span data-ttu-id="190f4-171">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="190f4-171">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-172">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-172">Type:</span></span>

*   <span data-ttu-id="190f4-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="190f4-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-174">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-174">Requirements</span></span>

|<span data-ttu-id="190f4-175">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-175">Requirement</span></span>| <span data-ttu-id="190f4-176">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-177">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-178">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-178">1.0</span></span>|
|[<span data-ttu-id="190f4-179">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-180">ReadItem</span></span>|
|[<span data-ttu-id="190f4-181">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-182">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-183">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="190f4-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="190f4-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="190f4-185">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="190f4-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="190f4-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="190f4-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="190f4-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="190f4-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-190">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-190">Type:</span></span>

*   <span data-ttu-id="190f4-191">Chaîne</span><span class="sxs-lookup"><span data-stu-id="190f4-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-192">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-192">Requirements</span></span>

|<span data-ttu-id="190f4-193">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-193">Requirement</span></span>| <span data-ttu-id="190f4-194">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-195">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-196">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-196">1.0</span></span>|
|[<span data-ttu-id="190f4-197">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-198">ReadItem</span></span>|
|[<span data-ttu-id="190f4-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-200">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="190f4-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="190f4-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="190f4-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-204">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-204">Type:</span></span>

*   <span data-ttu-id="190f4-205">Date</span><span class="sxs-lookup"><span data-stu-id="190f4-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-206">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-206">Requirements</span></span>

|<span data-ttu-id="190f4-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-207">Requirement</span></span>| <span data-ttu-id="190f4-208">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-209">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-210">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-210">1.0</span></span>|
|[<span data-ttu-id="190f4-211">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-212">ReadItem</span></span>|
|[<span data-ttu-id="190f4-213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-214">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="190f4-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="190f4-216">dateTimeModified :Date</span></span>

<span data-ttu-id="190f4-p111">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-219">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="190f4-219">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-220">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-220">Type:</span></span>

*   <span data-ttu-id="190f4-221">Date</span><span class="sxs-lookup"><span data-stu-id="190f4-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-222">Requirements</span></span>

|<span data-ttu-id="190f4-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-223">Requirement</span></span>| <span data-ttu-id="190f4-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-226">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-226">1.0</span></span>|
|[<span data-ttu-id="190f4-227">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-228">ReadItem</span></span>|
|[<span data-ttu-id="190f4-229">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-230">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-231">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="190f4-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="190f4-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="190f4-233">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="190f4-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="190f4-p112">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="190f4-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="190f4-236">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-236">Read mode</span></span>

<span data-ttu-id="190f4-237">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="190f4-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="190f4-238">Mode composition</span><span class="sxs-lookup"><span data-stu-id="190f4-238">Compose mode</span></span>

<span data-ttu-id="190f4-239">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="190f4-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="190f4-240">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="190f4-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-241">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-241">Type:</span></span>

*   <span data-ttu-id="190f4-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="190f4-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-243">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-243">Requirements</span></span>

|<span data-ttu-id="190f4-244">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-244">Requirement</span></span>| <span data-ttu-id="190f4-245">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-246">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-247">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-247">1.0</span></span>|
|[<span data-ttu-id="190f4-248">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-249">ReadItem</span></span>|
|[<span data-ttu-id="190f4-250">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-251">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-252">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-252">Example</span></span>

<span data-ttu-id="190f4-253">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="190f4-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="190f4-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="190f4-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="190f4-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="190f4-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="190f4-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-259">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="190f4-259">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-260">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-260">Type:</span></span>

*   [<span data-ttu-id="190f4-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="190f4-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="190f4-262">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-262">Requirements</span></span>

|<span data-ttu-id="190f4-263">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-263">Requirement</span></span>| <span data-ttu-id="190f4-264">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-265">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-266">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-266">1.0</span></span>|
|[<span data-ttu-id="190f4-267">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-268">ReadItem</span></span>|
|[<span data-ttu-id="190f4-269">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-270">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="190f4-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="190f4-271">internetMessageId :String</span></span>

<span data-ttu-id="190f4-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-274">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-274">Type:</span></span>

*   <span data-ttu-id="190f4-275">Chaîne</span><span class="sxs-lookup"><span data-stu-id="190f4-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-276">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-276">Requirements</span></span>

|<span data-ttu-id="190f4-277">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-277">Requirement</span></span>| <span data-ttu-id="190f4-278">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-279">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-280">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-280">1.0</span></span>|
|[<span data-ttu-id="190f4-281">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-282">ReadItem</span></span>|
|[<span data-ttu-id="190f4-283">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-284">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-285">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="190f4-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="190f4-286">itemClass :String</span></span>

<span data-ttu-id="190f4-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="190f4-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="190f4-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="190f4-291">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-291">Type</span></span> | <span data-ttu-id="190f4-292">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-292">Description</span></span> | <span data-ttu-id="190f4-293">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="190f4-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="190f4-294">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="190f4-294">Appointment items</span></span> | <span data-ttu-id="190f4-295">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="190f4-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="190f4-296">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="190f4-296">Message items</span></span> | <span data-ttu-id="190f4-297">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="190f4-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="190f4-298">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="190f4-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-299">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-299">Type:</span></span>

*   <span data-ttu-id="190f4-300">Chaîne</span><span class="sxs-lookup"><span data-stu-id="190f4-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-301">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-301">Requirements</span></span>

|<span data-ttu-id="190f4-302">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-302">Requirement</span></span>| <span data-ttu-id="190f4-303">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-304">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-305">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-305">1.0</span></span>|
|[<span data-ttu-id="190f4-306">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-307">ReadItem</span></span>|
|[<span data-ttu-id="190f4-308">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-309">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-310">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="190f4-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="190f4-311">(nullable) itemId :String</span></span>

<span data-ttu-id="190f4-p118">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-314">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="190f4-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="190f4-315">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="190f4-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="190f4-316">Avant d’effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande `Office.context.mailbox.convertToRestId`, qui est disponible à partir de l’ensemble de conditions requises 1.3.</span><span class="sxs-lookup"><span data-stu-id="190f4-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="190f4-317">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="190f4-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-318">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-318">Type:</span></span>

*   <span data-ttu-id="190f4-319">Chaîne</span><span class="sxs-lookup"><span data-stu-id="190f4-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-320">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-320">Requirements</span></span>

|<span data-ttu-id="190f4-321">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-321">Requirement</span></span>| <span data-ttu-id="190f4-322">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-323">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-324">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-324">1.0</span></span>|
|[<span data-ttu-id="190f4-325">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-326">ReadItem</span></span>|
|[<span data-ttu-id="190f4-327">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-328">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-329">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-329">Example</span></span>

<span data-ttu-id="190f4-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="190f4-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="190f4-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="190f4-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="190f4-333">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="190f4-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="190f4-334">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="190f4-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-335">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-335">Type:</span></span>

*   [<span data-ttu-id="190f4-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="190f4-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="190f4-337">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-337">Requirements</span></span>

|<span data-ttu-id="190f4-338">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-338">Requirement</span></span>| <span data-ttu-id="190f4-339">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-340">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-340">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-341">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-341">1.0</span></span>|
|[<span data-ttu-id="190f4-342">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-343">ReadItem</span></span>|
|[<span data-ttu-id="190f4-344">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-345">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-346">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="190f4-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="190f4-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="190f4-348">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="190f4-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="190f4-349">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-349">Read mode</span></span>

<span data-ttu-id="190f4-350">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="190f4-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="190f4-351">Mode composition</span><span class="sxs-lookup"><span data-stu-id="190f4-351">Compose mode</span></span>

<span data-ttu-id="190f4-352">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="190f4-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-353">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-353">Type:</span></span>

*   <span data-ttu-id="190f4-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="190f4-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-355">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-355">Requirements</span></span>

|<span data-ttu-id="190f4-356">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-356">Requirement</span></span>| <span data-ttu-id="190f4-357">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-358">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-359">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-359">1.0</span></span>|
|[<span data-ttu-id="190f4-360">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-361">ReadItem</span></span>|
|[<span data-ttu-id="190f4-362">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-363">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-364">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="190f4-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="190f4-365">normalizedSubject :String</span></span>

<span data-ttu-id="190f4-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="190f4-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="190f4-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-370">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-370">Type:</span></span>

*   <span data-ttu-id="190f4-371">Chaîne</span><span class="sxs-lookup"><span data-stu-id="190f4-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-372">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-372">Requirements</span></span>

|<span data-ttu-id="190f4-373">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-373">Requirement</span></span>| <span data-ttu-id="190f4-374">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-375">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-376">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-376">1.0</span></span>|
|[<span data-ttu-id="190f4-377">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-378">ReadItem</span></span>|
|[<span data-ttu-id="190f4-379">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-380">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-381">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="190f4-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="190f4-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="190f4-383">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="190f4-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="190f4-384">Le type d’objet et le niveau d’accès varie selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="190f4-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="190f4-385">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-385">Read mode</span></span>

<span data-ttu-id="190f4-386">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="190f4-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="190f4-387">Mode composition</span><span class="sxs-lookup"><span data-stu-id="190f4-387">Compose mode</span></span>

<span data-ttu-id="190f4-388">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="190f4-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-389">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-389">Type:</span></span>

*   <span data-ttu-id="190f4-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="190f4-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-391">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-391">Requirements</span></span>

|<span data-ttu-id="190f4-392">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-392">Requirement</span></span>| <span data-ttu-id="190f4-393">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-394">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-394">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-395">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-395">1.0</span></span>|
|[<span data-ttu-id="190f4-396">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-397">ReadItem</span></span>|
|[<span data-ttu-id="190f4-398">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-399">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-400">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="190f4-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="190f4-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="190f4-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-404">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-404">Type:</span></span>

*   [<span data-ttu-id="190f4-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="190f4-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="190f4-406">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-406">Requirements</span></span>

|<span data-ttu-id="190f4-407">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-407">Requirement</span></span>| <span data-ttu-id="190f4-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-409">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-410">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-410">1.0</span></span>|
|[<span data-ttu-id="190f4-411">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-412">ReadItem</span></span>|
|[<span data-ttu-id="190f4-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-414">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="190f4-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="190f4-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="190f4-417">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="190f4-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="190f4-418">Le type d’objet et le niveau d’accès varie selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="190f4-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="190f4-419">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-419">Read mode</span></span>

<span data-ttu-id="190f4-420">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="190f4-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="190f4-421">Mode composition</span><span class="sxs-lookup"><span data-stu-id="190f4-421">Compose mode</span></span>

<span data-ttu-id="190f4-422">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="190f4-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-423">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-423">Type:</span></span>

*   <span data-ttu-id="190f4-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="190f4-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-425">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-425">Requirements</span></span>

|<span data-ttu-id="190f4-426">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-426">Requirement</span></span>| <span data-ttu-id="190f4-427">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-428">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-429">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-429">1.0</span></span>|
|[<span data-ttu-id="190f4-430">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-431">ReadItem</span></span>|
|[<span data-ttu-id="190f4-432">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-433">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-434">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="190f4-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="190f4-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="190f4-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="190f4-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="190f4-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="190f4-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-440">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="190f4-440">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-441">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-441">Type:</span></span>

*   [<span data-ttu-id="190f4-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="190f4-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="190f4-443">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-443">Requirements</span></span>

|<span data-ttu-id="190f4-444">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-444">Requirement</span></span>| <span data-ttu-id="190f4-445">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-446">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-447">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-447">1.0</span></span>|
|[<span data-ttu-id="190f4-448">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-449">ReadItem</span></span>|
|[<span data-ttu-id="190f4-450">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-451">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-452">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="190f4-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="190f4-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="190f4-454">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="190f4-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="190f4-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="190f4-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="190f4-457">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-457">Read mode</span></span>

<span data-ttu-id="190f4-458">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="190f4-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="190f4-459">Mode composition</span><span class="sxs-lookup"><span data-stu-id="190f4-459">Compose mode</span></span>

<span data-ttu-id="190f4-460">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="190f4-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="190f4-461">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="190f4-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-462">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-462">Type:</span></span>

*   <span data-ttu-id="190f4-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="190f4-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-464">Requirements</span></span>

|<span data-ttu-id="190f4-465">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-465">Requirement</span></span>| <span data-ttu-id="190f4-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-468">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-468">1.0</span></span>|
|[<span data-ttu-id="190f4-469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-470">ReadItem</span></span>|
|[<span data-ttu-id="190f4-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-472">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-473">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-473">Example</span></span>

<span data-ttu-id="190f4-474">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="190f4-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="190f4-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="190f4-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="190f4-476">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="190f4-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="190f4-477">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="190f4-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="190f4-478">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-478">Read mode</span></span>

<span data-ttu-id="190f4-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="190f4-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="190f4-481">Mode composition</span><span class="sxs-lookup"><span data-stu-id="190f4-481">Compose mode</span></span>

<span data-ttu-id="190f4-482">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="190f4-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="190f4-483">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-483">Type:</span></span>

*   <span data-ttu-id="190f4-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="190f4-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-485">Requirements</span></span>

|<span data-ttu-id="190f4-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-486">Requirement</span></span>| <span data-ttu-id="190f4-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-489">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-489">1.0</span></span>|
|[<span data-ttu-id="190f4-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-491">ReadItem</span></span>|
|[<span data-ttu-id="190f4-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-493">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="190f4-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="190f4-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="190f4-495">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="190f4-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="190f4-496">Le type d’objet et le niveau d’accès varie selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="190f4-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="190f4-497">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-497">Read mode</span></span>

<span data-ttu-id="190f4-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="190f4-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="190f4-500">Mode composition</span><span class="sxs-lookup"><span data-stu-id="190f4-500">Compose mode</span></span>

<span data-ttu-id="190f4-501">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="190f4-501">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="190f4-502">Type :</span><span class="sxs-lookup"><span data-stu-id="190f4-502">Type:</span></span>

*   <span data-ttu-id="190f4-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="190f4-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-504">Requirements</span></span>

|<span data-ttu-id="190f4-505">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-505">Requirement</span></span>| <span data-ttu-id="190f4-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-507">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-508">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-508">1.0</span></span>|
|[<span data-ttu-id="190f4-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-510">ReadItem</span></span>|
|[<span data-ttu-id="190f4-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-512">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-513">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="190f4-514">Méthodes</span><span class="sxs-lookup"><span data-stu-id="190f4-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="190f4-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="190f4-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="190f4-516">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="190f4-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="190f4-517">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="190f4-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="190f4-518">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="190f4-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-519">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-519">Parameters:</span></span>

|<span data-ttu-id="190f4-520">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-520">Name</span></span>| <span data-ttu-id="190f4-521">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-521">Type</span></span>| <span data-ttu-id="190f4-522">Attributs</span><span class="sxs-lookup"><span data-stu-id="190f4-522">Attributes</span></span>| <span data-ttu-id="190f4-523">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="190f4-524">String</span><span class="sxs-lookup"><span data-stu-id="190f4-524">String</span></span>||<span data-ttu-id="190f4-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="190f4-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="190f4-527">String</span><span class="sxs-lookup"><span data-stu-id="190f4-527">String</span></span>||<span data-ttu-id="190f4-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="190f4-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="190f4-530">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-530">Object</span></span>| <span data-ttu-id="190f4-531">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-531">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-532">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="190f4-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="190f4-533">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-533">Object</span></span>| <span data-ttu-id="190f4-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-534">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-535">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="190f4-536">fonction</span><span class="sxs-lookup"><span data-stu-id="190f4-536">function</span></span>| <span data-ttu-id="190f4-537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-537">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-538">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="190f4-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="190f4-539">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="190f4-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="190f4-540">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="190f4-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="190f4-541">Erreurs</span><span class="sxs-lookup"><span data-stu-id="190f4-541">Errors</span></span>

| <span data-ttu-id="190f4-542">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="190f4-542">Error code</span></span> | <span data-ttu-id="190f4-543">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="190f4-544">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="190f4-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="190f4-545">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="190f4-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="190f4-546">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="190f4-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="190f4-547">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-547">Requirements</span></span>

|<span data-ttu-id="190f4-548">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-548">Requirement</span></span>| <span data-ttu-id="190f4-549">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-550">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-551">1.1</span><span class="sxs-lookup"><span data-stu-id="190f4-551">1.1</span></span>|
|[<span data-ttu-id="190f4-552">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="190f4-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="190f4-554">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-555">Composition</span><span class="sxs-lookup"><span data-stu-id="190f4-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-556">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-556">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="190f4-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="190f4-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="190f4-558">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="190f4-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="190f4-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="190f4-562">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="190f4-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="190f4-563">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="190f4-563">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-564">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-564">Parameters:</span></span>

|<span data-ttu-id="190f4-565">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-565">Name</span></span>| <span data-ttu-id="190f4-566">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-566">Type</span></span>| <span data-ttu-id="190f4-567">Attributs</span><span class="sxs-lookup"><span data-stu-id="190f4-567">Attributes</span></span>| <span data-ttu-id="190f4-568">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="190f4-569">String</span><span class="sxs-lookup"><span data-stu-id="190f4-569">String</span></span>||<span data-ttu-id="190f4-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="190f4-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="190f4-572">String</span><span class="sxs-lookup"><span data-stu-id="190f4-572">String</span></span>||<span data-ttu-id="190f4-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="190f4-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="190f4-575">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-575">Object</span></span>| <span data-ttu-id="190f4-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-576">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-577">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="190f4-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="190f4-578">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-578">Object</span></span>| <span data-ttu-id="190f4-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-579">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-580">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="190f4-581">fonction</span><span class="sxs-lookup"><span data-stu-id="190f4-581">function</span></span>| <span data-ttu-id="190f4-582">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-582">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-583">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="190f4-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="190f4-584">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="190f4-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="190f4-585">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="190f4-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="190f4-586">Erreurs</span><span class="sxs-lookup"><span data-stu-id="190f4-586">Errors</span></span>

| <span data-ttu-id="190f4-587">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="190f4-587">Error code</span></span> | <span data-ttu-id="190f4-588">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="190f4-589">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="190f4-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="190f4-590">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-590">Requirements</span></span>

|<span data-ttu-id="190f4-591">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-591">Requirement</span></span>| <span data-ttu-id="190f4-592">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-593">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-593">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-594">1.1</span><span class="sxs-lookup"><span data-stu-id="190f4-594">1.1</span></span>|
|[<span data-ttu-id="190f4-595">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="190f4-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="190f4-597">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-598">Composition</span><span class="sxs-lookup"><span data-stu-id="190f4-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-599">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-599">Example</span></span>

<span data-ttu-id="190f4-600">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="190f4-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="190f4-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="190f4-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="190f4-602">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="190f4-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-603">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="190f4-603">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="190f4-604">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="190f4-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="190f4-605">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="190f4-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="190f4-p137">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="190f4-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-609">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-609">Parameters:</span></span>

|<span data-ttu-id="190f4-610">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-610">Name</span></span>| <span data-ttu-id="190f4-611">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-611">Type</span></span>| <span data-ttu-id="190f4-612">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="190f4-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="190f4-613">String &#124; Object</span></span>| |<span data-ttu-id="190f4-p138">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="190f4-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="190f4-616">**OU**</span><span class="sxs-lookup"><span data-stu-id="190f4-616">**OR**</span></span><br/><span data-ttu-id="190f4-p139">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="190f4-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="190f4-619">String</span><span class="sxs-lookup"><span data-stu-id="190f4-619">String</span></span> | <span data-ttu-id="190f4-620">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-620">&lt;optional&gt;</span></span> | <span data-ttu-id="190f4-p140">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="190f4-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="190f4-623">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="190f4-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-624">&lt;optional&gt;</span></span> | <span data-ttu-id="190f4-625">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="190f4-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="190f4-626">Chaîne</span><span class="sxs-lookup"><span data-stu-id="190f4-626">String</span></span> | | <span data-ttu-id="190f4-p141">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="190f4-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="190f4-629">String</span><span class="sxs-lookup"><span data-stu-id="190f4-629">String</span></span> | | <span data-ttu-id="190f4-630">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="190f4-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="190f4-631">Chaîne</span><span class="sxs-lookup"><span data-stu-id="190f4-631">String</span></span> | | <span data-ttu-id="190f4-p142">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="190f4-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="190f4-634">String</span><span class="sxs-lookup"><span data-stu-id="190f4-634">String</span></span> | | <span data-ttu-id="190f4-p143">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="190f4-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="190f4-638">function</span><span class="sxs-lookup"><span data-stu-id="190f4-638">function</span></span> | <span data-ttu-id="190f4-639">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-639">&lt;optional&gt;</span></span> | <span data-ttu-id="190f4-640">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="190f4-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="190f4-641">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-641">Requirements</span></span>

|<span data-ttu-id="190f4-642">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-642">Requirement</span></span>| <span data-ttu-id="190f4-643">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-644">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-645">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-645">1.0</span></span>|
|[<span data-ttu-id="190f4-646">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-647">ReadItem</span></span>|
|[<span data-ttu-id="190f4-648">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-649">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="190f4-650">Exemples</span><span class="sxs-lookup"><span data-stu-id="190f4-650">Examples</span></span>

<span data-ttu-id="190f4-651">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="190f4-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="190f4-652">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="190f4-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="190f4-653">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="190f4-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="190f4-654">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="190f4-654">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="190f4-655">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="190f4-655">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="190f4-656">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="190f4-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="190f4-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="190f4-658">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="190f4-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-659">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="190f4-659">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="190f4-660">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="190f4-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="190f4-661">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="190f4-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="190f4-p144">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="190f4-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-665">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-665">Parameters:</span></span>

|<span data-ttu-id="190f4-666">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-666">Name</span></span>| <span data-ttu-id="190f4-667">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-667">Type</span></span>| <span data-ttu-id="190f4-668">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="190f4-669">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="190f4-669">String &#124; Object</span></span>| | <span data-ttu-id="190f4-p145">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="190f4-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="190f4-672">**OU**</span><span class="sxs-lookup"><span data-stu-id="190f4-672">**OR**</span></span><br/><span data-ttu-id="190f4-p146">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="190f4-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="190f4-675">String</span><span class="sxs-lookup"><span data-stu-id="190f4-675">String</span></span> | <span data-ttu-id="190f4-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-676">&lt;optional&gt;</span></span> | <span data-ttu-id="190f4-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="190f4-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="190f4-679">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="190f4-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-680">&lt;optional&gt;</span></span> | <span data-ttu-id="190f4-681">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="190f4-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="190f4-682">Chaîne</span><span class="sxs-lookup"><span data-stu-id="190f4-682">String</span></span> | | <span data-ttu-id="190f4-p148">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="190f4-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="190f4-685">String</span><span class="sxs-lookup"><span data-stu-id="190f4-685">String</span></span> | | <span data-ttu-id="190f4-686">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="190f4-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="190f4-687">Chaîne</span><span class="sxs-lookup"><span data-stu-id="190f4-687">String</span></span> | | <span data-ttu-id="190f4-p149">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="190f4-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="190f4-690">String</span><span class="sxs-lookup"><span data-stu-id="190f4-690">String</span></span> | | <span data-ttu-id="190f4-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="190f4-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="190f4-694">function</span><span class="sxs-lookup"><span data-stu-id="190f4-694">function</span></span> | <span data-ttu-id="190f4-695">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-695">&lt;optional&gt;</span></span> | <span data-ttu-id="190f4-696">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="190f4-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="190f4-697">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-697">Requirements</span></span>

|<span data-ttu-id="190f4-698">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-698">Requirement</span></span>| <span data-ttu-id="190f4-699">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-700">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-700">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-701">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-701">1.0</span></span>|
|[<span data-ttu-id="190f4-702">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-703">ReadItem</span></span>|
|[<span data-ttu-id="190f4-704">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-705">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="190f4-706">Exemples</span><span class="sxs-lookup"><span data-stu-id="190f4-706">Examples</span></span>

<span data-ttu-id="190f4-707">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="190f4-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="190f4-708">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="190f4-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="190f4-709">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="190f4-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="190f4-710">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="190f4-710">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="190f4-711">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="190f4-711">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="190f4-712">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="190f4-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="190f4-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="190f4-714">Obtient les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="190f4-714">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-715">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="190f4-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-716">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-716">Requirements</span></span>

|<span data-ttu-id="190f4-717">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-717">Requirement</span></span>| <span data-ttu-id="190f4-718">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-719">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-719">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-720">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-720">1.0</span></span>|
|[<span data-ttu-id="190f4-721">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-722">ReadItem</span></span>|
|[<span data-ttu-id="190f4-723">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-724">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="190f4-725">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="190f4-725">Returns:</span></span>

<span data-ttu-id="190f4-726">Type : [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="190f4-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="190f4-727">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-727">Example</span></span>

<span data-ttu-id="190f4-728">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="190f4-728">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="190f4-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="190f4-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="190f4-730">Obtient un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="190f4-730">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-731">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="190f4-731">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-732">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-732">Parameters:</span></span>

|<span data-ttu-id="190f4-733">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-733">Name</span></span>| <span data-ttu-id="190f4-734">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-734">Type</span></span>| <span data-ttu-id="190f4-735">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="190f4-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="190f4-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="190f4-737">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="190f4-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="190f4-738">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-738">Requirements</span></span>

|<span data-ttu-id="190f4-739">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-739">Requirement</span></span>| <span data-ttu-id="190f4-740">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-741">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-742">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-742">1.0</span></span>|
|[<span data-ttu-id="190f4-743">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-744">Restreinte</span><span class="sxs-lookup"><span data-stu-id="190f4-744">Restricted</span></span>|
|[<span data-ttu-id="190f4-745">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-746">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="190f4-747">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="190f4-747">Returns:</span></span>

<span data-ttu-id="190f4-748">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="190f4-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="190f4-749">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="190f4-749">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="190f4-750">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="190f4-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="190f4-751">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="190f4-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="190f4-752">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="190f4-752">Value of `entityType`</span></span> | <span data-ttu-id="190f4-753">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="190f4-753">Type of objects in returned array</span></span> | <span data-ttu-id="190f4-754">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="190f4-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="190f4-755">String</span><span class="sxs-lookup"><span data-stu-id="190f4-755">String</span></span> | <span data-ttu-id="190f4-756">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="190f4-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="190f4-757">Contact</span><span class="sxs-lookup"><span data-stu-id="190f4-757">Contact</span></span> | <span data-ttu-id="190f4-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="190f4-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="190f4-759">String</span><span class="sxs-lookup"><span data-stu-id="190f4-759">String</span></span> | <span data-ttu-id="190f4-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="190f4-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="190f4-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="190f4-761">MeetingSuggestion</span></span> | <span data-ttu-id="190f4-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="190f4-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="190f4-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="190f4-763">PhoneNumber</span></span> | <span data-ttu-id="190f4-764">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="190f4-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="190f4-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="190f4-765">TaskSuggestion</span></span> | <span data-ttu-id="190f4-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="190f4-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="190f4-767">String</span><span class="sxs-lookup"><span data-stu-id="190f4-767">String</span></span> | <span data-ttu-id="190f4-768">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="190f4-768">**Restricted**</span></span> |

<span data-ttu-id="190f4-769">Type : Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="190f4-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="190f4-770">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-770">Example</span></span>

<span data-ttu-id="190f4-771">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="190f4-771">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="190f4-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="190f4-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="190f4-773">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="190f4-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-774">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="190f4-774">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="190f4-775">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="190f4-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-776">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-776">Parameters:</span></span>

|<span data-ttu-id="190f4-777">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-777">Name</span></span>| <span data-ttu-id="190f4-778">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-778">Type</span></span>| <span data-ttu-id="190f4-779">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="190f4-780">String</span><span class="sxs-lookup"><span data-stu-id="190f4-780">String</span></span>|<span data-ttu-id="190f4-781">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="190f4-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="190f4-782">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-782">Requirements</span></span>

|<span data-ttu-id="190f4-783">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-783">Requirement</span></span>| <span data-ttu-id="190f4-784">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-785">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-785">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-786">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-786">1.0</span></span>|
|[<span data-ttu-id="190f4-787">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-788">ReadItem</span></span>|
|[<span data-ttu-id="190f4-789">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-790">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="190f4-791">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="190f4-791">Returns:</span></span>

<span data-ttu-id="190f4-p152">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="190f4-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="190f4-794">Type : Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="190f4-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="190f4-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="190f4-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="190f4-796">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="190f4-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-797">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="190f4-797">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="190f4-p153">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="190f4-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="190f4-801">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="190f4-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="190f4-802">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="190f4-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="190f4-p154">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="190f4-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="190f4-805">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-805">Requirements</span></span>

|<span data-ttu-id="190f4-806">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-806">Requirement</span></span>| <span data-ttu-id="190f4-807">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-808">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-809">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-809">1.0</span></span>|
|[<span data-ttu-id="190f4-810">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-811">ReadItem</span></span>|
|[<span data-ttu-id="190f4-812">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-813">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="190f4-814">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="190f4-814">Returns:</span></span>

<span data-ttu-id="190f4-p155">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="190f4-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="190f4-817">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="190f4-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="190f4-818">Object</span><span class="sxs-lookup"><span data-stu-id="190f4-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="190f4-819">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-819">Example</span></span>

<span data-ttu-id="190f4-820">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="190f4-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="190f4-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="190f4-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="190f4-822">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="190f4-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="190f4-823">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="190f4-823">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="190f4-824">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="190f4-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="190f4-p156">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="190f4-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-827">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-827">Parameters:</span></span>

|<span data-ttu-id="190f4-828">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-828">Name</span></span>| <span data-ttu-id="190f4-829">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-829">Type</span></span>| <span data-ttu-id="190f4-830">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="190f4-831">String</span><span class="sxs-lookup"><span data-stu-id="190f4-831">String</span></span>|<span data-ttu-id="190f4-832">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="190f4-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="190f4-833">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-833">Requirements</span></span>

|<span data-ttu-id="190f4-834">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-834">Requirement</span></span>| <span data-ttu-id="190f4-835">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-836">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-836">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-837">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-837">1.0</span></span>|
|[<span data-ttu-id="190f4-838">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-839">ReadItem</span></span>|
|[<span data-ttu-id="190f4-840">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-841">Lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="190f4-842">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="190f4-842">Returns:</span></span>

<span data-ttu-id="190f4-843">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="190f4-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="190f4-844">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="190f4-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="190f4-845">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="190f4-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="190f4-846">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="190f4-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="190f4-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="190f4-848">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="190f4-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="190f4-p157">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="190f4-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-851">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-851">Parameters:</span></span>

|<span data-ttu-id="190f4-852">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-852">Name</span></span>| <span data-ttu-id="190f4-853">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-853">Type</span></span>| <span data-ttu-id="190f4-854">Attributs</span><span class="sxs-lookup"><span data-stu-id="190f4-854">Attributes</span></span>| <span data-ttu-id="190f4-855">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="190f4-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="190f4-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="190f4-p158">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="190f4-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="190f4-860">Object</span><span class="sxs-lookup"><span data-stu-id="190f4-860">Object</span></span>| <span data-ttu-id="190f4-861">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-861">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-862">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="190f4-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="190f4-863">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-863">Object</span></span>| <span data-ttu-id="190f4-864">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-864">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-865">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="190f4-866">fonction</span><span class="sxs-lookup"><span data-stu-id="190f4-866">function</span></span>||<span data-ttu-id="190f4-867">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="190f4-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="190f4-868">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="190f4-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="190f4-869">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="190f4-869">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="190f4-870">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-870">Requirements</span></span>

|<span data-ttu-id="190f4-871">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-871">Requirement</span></span>| <span data-ttu-id="190f4-872">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-873">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-874">1.2</span><span class="sxs-lookup"><span data-stu-id="190f4-874">1.2</span></span>|
|[<span data-ttu-id="190f4-875">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="190f4-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="190f4-877">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-878">Composition</span><span class="sxs-lookup"><span data-stu-id="190f4-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="190f4-879">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="190f4-879">Returns:</span></span>

<span data-ttu-id="190f4-880">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="190f4-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="190f4-881">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="190f4-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="190f4-882">String</span><span class="sxs-lookup"><span data-stu-id="190f4-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="190f4-883">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-883">Example</span></span>

```JavaScript
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="190f4-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="190f4-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="190f4-885">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="190f4-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="190f4-p160">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="190f4-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-889">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-889">Parameters:</span></span>

|<span data-ttu-id="190f4-890">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-890">Name</span></span>| <span data-ttu-id="190f4-891">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-891">Type</span></span>| <span data-ttu-id="190f4-892">Attributs</span><span class="sxs-lookup"><span data-stu-id="190f4-892">Attributes</span></span>| <span data-ttu-id="190f4-893">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="190f4-894">function</span><span class="sxs-lookup"><span data-stu-id="190f4-894">function</span></span>||<span data-ttu-id="190f4-895">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="190f4-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="190f4-896">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="190f4-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="190f4-897">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="190f4-897">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="190f4-898">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-898">Object</span></span>| <span data-ttu-id="190f4-899">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-899">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-900">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-900">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="190f4-901">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="190f4-902">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-902">Requirements</span></span>

|<span data-ttu-id="190f4-903">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-903">Requirement</span></span>| <span data-ttu-id="190f4-904">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-905">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-906">1.0</span><span class="sxs-lookup"><span data-stu-id="190f4-906">1.0</span></span>|
|[<span data-ttu-id="190f4-907">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="190f4-908">ReadItem</span></span>|
|[<span data-ttu-id="190f4-909">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-910">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="190f4-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-911">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-911">Example</span></span>

<span data-ttu-id="190f4-p163">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="190f4-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="190f4-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="190f4-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="190f4-916">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="190f4-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="190f4-p164">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="190f4-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-921">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-921">Parameters:</span></span>

|<span data-ttu-id="190f4-922">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-922">Name</span></span>| <span data-ttu-id="190f4-923">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-923">Type</span></span>| <span data-ttu-id="190f4-924">Attributs</span><span class="sxs-lookup"><span data-stu-id="190f4-924">Attributes</span></span>| <span data-ttu-id="190f4-925">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="190f4-926">String</span><span class="sxs-lookup"><span data-stu-id="190f4-926">String</span></span>||<span data-ttu-id="190f4-p165">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="190f4-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="190f4-929">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-929">Object</span></span>| <span data-ttu-id="190f4-930">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-930">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-931">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="190f4-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="190f4-932">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-932">Object</span></span>| <span data-ttu-id="190f4-933">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-933">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-934">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="190f4-935">fonction</span><span class="sxs-lookup"><span data-stu-id="190f4-935">function</span></span>| <span data-ttu-id="190f4-936">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-936">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-937">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="190f4-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="190f4-938">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="190f4-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="190f4-939">Erreurs</span><span class="sxs-lookup"><span data-stu-id="190f4-939">Errors</span></span>

| <span data-ttu-id="190f4-940">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="190f4-940">Error code</span></span> | <span data-ttu-id="190f4-941">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="190f4-942">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="190f4-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="190f4-943">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-943">Requirements</span></span>

|<span data-ttu-id="190f4-944">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-944">Requirement</span></span>| <span data-ttu-id="190f4-945">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-946">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-946">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-947">1.1</span><span class="sxs-lookup"><span data-stu-id="190f4-947">1.1</span></span>|
|[<span data-ttu-id="190f4-948">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="190f4-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="190f4-950">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-951">Composition</span><span class="sxs-lookup"><span data-stu-id="190f4-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-952">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-952">Example</span></span>

<span data-ttu-id="190f4-953">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="190f4-953">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="190f4-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="190f4-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="190f4-955">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="190f4-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="190f4-p166">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="190f4-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="190f4-959">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="190f4-959">Parameters:</span></span>

|<span data-ttu-id="190f4-960">Nom</span><span class="sxs-lookup"><span data-stu-id="190f4-960">Name</span></span>| <span data-ttu-id="190f4-961">Type</span><span class="sxs-lookup"><span data-stu-id="190f4-961">Type</span></span>| <span data-ttu-id="190f4-962">Attributs</span><span class="sxs-lookup"><span data-stu-id="190f4-962">Attributes</span></span>| <span data-ttu-id="190f4-963">Description</span><span class="sxs-lookup"><span data-stu-id="190f4-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="190f4-964">String</span><span class="sxs-lookup"><span data-stu-id="190f4-964">String</span></span>||<span data-ttu-id="190f4-p167">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="190f4-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="190f4-968">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-968">Object</span></span>| <span data-ttu-id="190f4-969">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-969">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-970">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="190f4-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="190f4-971">Objet</span><span class="sxs-lookup"><span data-stu-id="190f4-971">Object</span></span>| <span data-ttu-id="190f4-972">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-972">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-973">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="190f4-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="190f4-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="190f4-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="190f4-975">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="190f4-975">&lt;optional&gt;</span></span>|<span data-ttu-id="190f4-p168">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="190f4-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="190f4-p169">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="190f4-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="190f4-980">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="190f4-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="190f4-981">fonction</span><span class="sxs-lookup"><span data-stu-id="190f4-981">function</span></span>||<span data-ttu-id="190f4-982">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="190f4-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="190f4-983">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="190f4-983">Requirements</span></span>

|<span data-ttu-id="190f4-984">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="190f4-984">Requirement</span></span>| <span data-ttu-id="190f4-985">Valeur</span><span class="sxs-lookup"><span data-stu-id="190f4-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="190f4-986">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="190f4-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="190f4-987">1.2</span><span class="sxs-lookup"><span data-stu-id="190f4-987">1.2</span></span>|
|[<span data-ttu-id="190f4-988">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="190f4-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="190f4-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="190f4-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="190f4-990">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="190f4-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="190f4-991">Composition</span><span class="sxs-lookup"><span data-stu-id="190f4-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="190f4-992">Exemple</span><span class="sxs-lookup"><span data-stu-id="190f4-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```